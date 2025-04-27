from flask import Flask, request, render_template, send_file, redirect, url_for, session, jsonify, flash
import pandas as pd
from docxtpl import DocxTemplate
import os
import re
import json
from dotenv import load_dotenv
from werkzeug.utils import secure_filename
from datetime import datetime
from flask_mail import Mail, Message

app = Flask(__name__)
app.secret_key = 'your_secret_key'

load_dotenv('.env')

UPLOAD_FOLDER = 'uploads'
GENERATED_FOLDER = 'generated'
SETTINGS_PATH = os.path.join(UPLOAD_FOLDER, 'settings.json')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(GENERATED_FOLDER, exist_ok=True)

cached_docx_path = None
cached_csv_path = None
cached_dataframe = None
date_column = None
date_min = None
date_max = None

app.config['MAIL_SERVER'] = 'smtp.gmail.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USE_SSL'] = False
app.config['MAIL_USERNAME'] = os.getenv('MAIL_USERNAME')
app.config['MAIL_PASSWORD'] = os.getenv('MAIL_PASSWORD')
app.config['MAIL_DEFAULT_SENDER'] = os.getenv('MAIL_USERNAME')

mail = Mail(app)

def extract_template_variables(template_path):
    doc = DocxTemplate(template_path)
    full_text = ""
    for paragraph in doc.get_docx().paragraphs:
        full_text += paragraph.text + "\n"
    matches = re.findall(r"\{\{\s*(.*?)\s*\}\}", full_text)
    return list(set([match.strip() for match in matches]))

def analyze_csv(csv_path):
    global cached_dataframe, date_column, date_min, date_max

    df = pd.read_csv(csv_path)
    cached_dataframe = df

    for col in df.columns:
        try:
            parsed_dates = pd.to_datetime(df[col], errors='coerce')
            if parsed_dates.notna().sum() > 0:
                date_column = col
                df[col] = parsed_dates
                date_min = parsed_dates.min().date()
                date_max = parsed_dates.max().date()
                break
        except:
            continue

    if not date_column:
        raise Exception("沒有找到日期欄位")

    return df

def convert_docx_to_html(template_path):
    from docx import Document
    document = Document(template_path)
    html = ""

    for para in document.paragraphs:
        line = para.text
        line = re.sub(r"\{\{\s*(.*?)\s*\}\}", r'<span class="variable editable" data-variable="\1">{{\1}}</span>', line)
        html += f"<p>{line}</p>\n"

    return html

def evaluate_formula(formula: str, df: pd.DataFrame, context: dict = None, formulas: dict = None) -> float:
    import numpy as np

    formula = formula.strip()
    if context is None:
        context = {}
    if formulas is None:
        formulas = {}

    try:
        if formula.startswith("'") and formula.endswith("'"):
            return formula.strip("'")

        col_map = {col.lower(): col for col in df.columns}
        variables_in_formula = re.findall(r"[a-zA-Z_][a-zA-Z0-9_]*", formula)
        
        # 首先處理公式中的變數依賴
        processed_vars = set()
        for var in variables_in_formula:
            if var not in context and var in formulas and var not in processed_vars:
                # 避免循環依賴
                processed_vars.add(var)
                try:
                    # 遞迴計算依賴變數的值
                    context[var] = evaluate_formula(formulas[var], df, context, formulas)
                    print(f"已計算變數 {var} = {context[var]}")
                except Exception as e:
                    print(f"計算變數 {var} 時發生錯誤: {str(e)}")
                    raise ValueError(f"變數 {var} 計算錯誤: {str(e)}")

        # 處理列名大小寫不敏感
        for var in variables_in_formula:
            var_lower = var.lower()
            if var_lower in col_map:
                correct_col = col_map[var_lower]
                formula = re.sub(r'\b' + re.escape(var) + r'\b', correct_col, formula)

        # 處理特殊函數
        formula = re.sub(r'COUNT\s*\(\s*DISTINCT\s*\((.*?)\)\s*\)', r'df["\1"].nunique()', formula, flags=re.I)
        formula = re.sub(r'MODE\s*\(\s*(.*?)\s*\)', r'df["\1"].mode().iloc[0] if len(df["\1"].mode()) > 0 else None', formula, flags=re.I)
        # 處理特殊的 COUNT(condition) 情況
        count_pattern = re.compile(r'COUNT\s*\(\s*(.+?)\s*\)', re.I)
        for match in count_pattern.finditer(formula):
            content = match.group(1)
            if '==' in content or '!=' in content or '>' in content or '<' in content or '>=' in content or '<=' in content:
                # 條件表達式
                replacement = f'(df.query("{content.replace("==", "==")}", engine="python").shape[0])'
                formula = formula.replace(match.group(0), replacement)
            else:
                # 普通列計數
                replacement = f'df["{content}"].count()'
                formula = formula.replace(match.group(0), replacement)

        formula = re.sub(r'SUM\s*\(\s*(.*?)\s*\)', r'df.eval("\1").sum()', formula, flags=re.I)
        formula = re.sub(r'MEAN\s*\(\s*(.*?)\s*\)', r'df.eval("\1").mean()', formula, flags=re.I)
        formula = re.sub(r'MAX\s*\(\s*(.*?)\s*\)', r'df.eval("\1").max()', formula, flags=re.I)
        formula = re.sub(r'MIN\s*\(\s*(.*?)\s*\)', r'df.eval("\1").min()', formula, flags=re.I)
        formula = re.sub(r'MEDIAN\s*\(\s*(.*?)\s*\)', r'df.eval("\1").median()', formula, flags=re.I)
        formula = re.sub(r'STD\s*\(\s*(.*?)\s*\)', r'df.eval("\1").std()', formula, flags=re.I)
        formula = re.sub(r'VAR\s*\(\s*(.*?)\s*\)', r'df.eval("\1").var()', formula, flags=re.I)

        if formula.upper().startswith("PERCENT_CHANGE("):
            inside = re.findall(r'PERCENT_CHANGE\((.*?)\)', formula, flags=re.I)[0]
            values = df.eval(inside.strip())
            return (values.iloc[-1] - values.iloc[0]) / values.iloc[0] * 100

        if formula.upper().startswith("DIFF("):
            inside = re.findall(r'DIFF\((.*?)\)', formula, flags=re.I)[0]
            values = df.eval(inside.strip())
            return values.iloc[-1] - values.iloc[0]

        if formula.upper().startswith("CAGR("):
            inside = re.findall(r'CAGR\((.*?)\)', formula, flags=re.I)[0]
            parts = [p.strip() for p in inside.split(",")]
            start_val = df.eval(parts[0]).iloc[0]
            end_val = df.eval(parts[1]).iloc[-1]
            periods = float(parts[2])
            return (end_val / start_val) ** (1/periods) - 1

        # 將變數替換為實際值
        for var_name, var_value in context.items():
            if isinstance(var_value, str):
                # 如果是字串，需要加引號
                formula = re.sub(r'\b' + re.escape(var_name) + r'\b', f"'{var_value}'", formula)
            else:
                # 如果是數值，直接替換
                formula = re.sub(r'\b' + re.escape(var_name) + r'\b', str(var_value), formula)

        print(f"最終評估公式: {formula}")
        eval_globals = {"df": df, "np": np}
        eval_globals.update(context)

        result = eval(formula, eval_globals)
        return result

    except Exception as e:
        raise ValueError(f"公式錯誤：{str(e)}")

# 新增首頁路由
@app.route('/')
def home():
    message = request.args.get('message', '')
    return render_template('home.html', message=message)

# 上傳頁面 - 對應原本的 index，保持原來的功能
@app.route('/upload', methods=['GET', 'POST'])
def upload():
    global cached_docx_path, cached_csv_path

    if request.method == 'POST':
        if 'docx_file' in request.files:
            docx_file = request.files['docx_file']
            if docx_file.filename != '':
                filename = datetime.now().strftime('%Y%m%d%H%M%S_') + secure_filename(docx_file.filename)
                cached_docx_path = os.path.join(UPLOAD_FOLDER, filename)
                docx_file.save(cached_docx_path)

        if 'csv_file' in request.files:
            csv_file = request.files['csv_file']
            if csv_file.filename != '':
                filename = datetime.now().strftime('%Y%m%d%H%M%S_') + secure_filename(csv_file.filename)
                cached_csv_path = os.path.join(UPLOAD_FOLDER, filename)
                csv_file.save(cached_csv_path)

        if 'settings_file' in request.files:
            settings_file = request.files['settings_file']
            if settings_file.filename != '' and settings_file.filename.endswith('.json'):
                settings_file.save(SETTINGS_PATH)

        if cached_docx_path and cached_csv_path:
            return redirect(url_for('preview'))

    return render_template('index.html')

# 保留原來的 index 路由，但現在轉發到 upload 函數
@app.route('/index', methods=['GET', 'POST'])
def index():
    return upload()

@app.route('/preview', methods=['GET'])
def preview():
    global cached_docx_path, cached_csv_path

    if not cached_docx_path or not cached_csv_path:
        return redirect(url_for('upload'))

    try:
        variables = extract_template_variables(cached_docx_path)
        analyze_csv(cached_csv_path)
        html_content = convert_docx_to_html(cached_docx_path)

        return render_template('preview.html',
                            html_content=html_content,
                            variables=variables,
                            date_column=date_column,
                            date_min=date_min,
                            date_max=date_max)
    except Exception as e:
        print(f"預覽生成錯誤: {str(e)}")
        return render_template('index.html', error=f"預覽生成失敗: {str(e)}")

@app.route('/filter_data', methods=['POST'])
def filter_data():
    global cached_dataframe, date_column

    data = request.json
    start_date = pd.to_datetime(data.get('start_date'))
    end_date = pd.to_datetime(data.get('end_date'))

    if not date_column:
        return jsonify([])

    filtered = cached_dataframe[
        (cached_dataframe[date_column] >= start_date) &
        (cached_dataframe[date_column] <= end_date)
    ]

    return jsonify(filtered.to_dict(orient='records'))

@app.route('/render_preview', methods=['POST'])
def render_preview():
    data = request.json.get('data', [])
    formulas = request.json.get('formulas', {})

    if not data:
        return jsonify({})

    df = pd.DataFrame(data)
    results = {}
    calculated_context = {}

    # 計算公式的依賴關係
    formula_graph = {}
    for var, formula in formulas.items():
        dependencies = set()
        for other_var in formulas:
            if other_var != var and re.search(r'\b' + re.escape(other_var) + r'\b', formula):
                dependencies.add(other_var)
        formula_graph[var] = dependencies

    # 對公式進行拓撲排序
    visited = set()
    temp_visited = set()
    order = []

    def topo_sort(node):
        if node in temp_visited:
            # 發現循環依賴
            print(f"循環依賴: {node}")
            return
        if node in visited:
            return

        temp_visited.add(node)
        for dep in formula_graph.get(node, []):
            topo_sort(dep)
        temp_visited.remove(node)
        visited.add(node)
        order.append(node)

    for var in formulas:
        if var not in visited:
            topo_sort(var)

    # 按照拓撲排序順序計算公式
    order.reverse()  # 反轉，使依賴項先計算
    print(f"公式計算順序: {order}")

    # 先計算沒有依賴的公式
    for var in formulas:
        if var not in order:
            try:
                value = evaluate_formula(formulas[var], df, context=calculated_context, formulas=formulas)
                if hasattr(value, "item"):
                    value = value.item()
                if isinstance(value, (int, float)):
                    value = round(value, 2)
                results[var] = value
                calculated_context[var] = value
            except Exception as e:
                print(f"⚡錯誤發生在變數 [{var}]，公式 [{formulas[var]}]，錯誤訊息：{e}")
                results[var] = f"錯誤: {str(e)}"

    # 按依賴順序計算其餘公式
    for var in order:
        try:
            value = evaluate_formula(formulas[var], df, context=calculated_context, formulas=formulas)
            if hasattr(value, "item"):
                value = value.item()
            if isinstance(value, (int, float)):
                value = round(value, 2)
            results[var] = value
            calculated_context[var] = value
        except Exception as e:
            print(f"⚡錯誤發生在變數 [{var}]，公式 [{formulas[var]}]，錯誤訊息：{e}")
            results[var] = f"錯誤: {str(e)}"

    return jsonify(results)

@app.route('/render', methods=['POST'])
def render_word():
    global cached_docx_path

    formulas = request.json.get('formulas', {})
    data = request.json.get('data', [])

    if not data:
        return "資料錯誤", 400

    filtered_df = pd.DataFrame(data)
    context = {}
    calculated_context = {}

    # 計算公式的依賴關係
    formula_graph = {}
    for var, formula in formulas.items():
        dependencies = set()
        for other_var in formulas:
            if other_var != var and re.search(r'\b' + re.escape(other_var) + r'\b', formula):
                dependencies.add(other_var)
        formula_graph[var] = dependencies

    # 對公式進行拓撲排序
    visited = set()
    temp_visited = set()
    order = []

    def topo_sort(node):
        if node in temp_visited:
            # 發現循環依賴
            print(f"循環依賴: {node}")
            return
        if node in visited:
            return

        temp_visited.add(node)
        for dep in formula_graph.get(node, []):
            topo_sort(dep)
        temp_visited.remove(node)
        visited.add(node)
        order.append(node)

    for var in formulas:
        if var not in visited:
            topo_sort(var)

    # 按照拓撲排序順序計算公式
    order.reverse()  # 反轉，使依賴項先計算
    print(f"公式計算順序: {order}")

    # 先計算沒有依賴的公式
    for var in formulas:
        if var not in order:
            try:
                value = evaluate_formula(formulas[var], filtered_df, context=calculated_context, formulas=formulas)
                if hasattr(value, "item"):
                    value = value.item()
                if isinstance(value, (int, float)):
                    value = round(value, 2)
                context[var] = value
                calculated_context[var] = value
            except Exception as e:
                context[var] = f"錯誤: {str(e)}"

    # 按依賴順序計算其餘公式
    for var in order:
        try:
            value = evaluate_formula(formulas[var], filtered_df, context=calculated_context, formulas=formulas)
            if hasattr(value, "item"):
                value = value.item()
            if isinstance(value, (int, float)):
                value = round(value, 2)
            context[var] = value
            calculated_context[var] = value
        except Exception as e:
            context[var] = f"錯誤: {str(e)}"

    doc = DocxTemplate(cached_docx_path)
    doc.render(context)

    output_path = os.path.join(GENERATED_FOLDER, 'final_report.docx')
    doc.save(output_path)

    return send_file(output_path, as_attachment=True)

@app.route('/save_settings', methods=['POST'])
def save_settings():
    settings = request.json
    with open(SETTINGS_PATH, 'w', encoding='utf-8') as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)
    return "設定已儲存", 200

@app.route('/load_settings', methods=['GET'])
def load_settings():
    if not os.path.exists(SETTINGS_PATH):
        return jsonify({}), 404

    with open(SETTINGS_PATH, 'r', encoding='utf-8') as f:
        settings = json.load(f)

    return jsonify(settings)

# 聯絡表單處理路由
@app.route('/contact', methods=['POST'])
def contact():
    name = request.form.get('name')
    email = request.form.get('email')
    message = request.form.get('message')

    if not name or not email or not message:
        flash('請完整填寫所有欄位喔～', 'warning')
        return redirect(url_for('home'))

    try:
        msg = Message('網站聯絡表單新訊息 ✉️', recipients=[app.config['MAIL_USERNAME']], charset='utf-8')
        msg.body = f"""
收到一封新的聯絡表單：

姓名：{name}
電子郵件：{email}
訊息內容：
{message}
        """
        mail.send(msg)
        flash('感謝您的聯絡！我們已收到您的訊息！', 'success')
    except Exception as e:
        print(f"寄信失敗：{e}")
        flash('寄信時發生錯誤，請稍後再試！', 'warning')

    return redirect(url_for('home'))

# 處理示例模板下載的路由
@app.route('/download_sample', methods=['GET'])
def download_sample():
    sample_path = os.path.join('static', 'samples', 'sample_template.docx')
    # 確保目錄存在
    os.makedirs(os.path.dirname(sample_path), exist_ok=True)
    
    # 如果文件不存在，創建一個簡單的樣本文件
    if not os.path.exists(sample_path):
        from docx import Document
        doc = Document()
        doc.add_heading('範例模板', 0)
        doc.add_paragraph('這是一個範例模板，您可以在其中使用 {{ 變數名稱 }} 格式來定義需要替換的變數。')
        p = doc.add_paragraph('例如：本月銷售總額為 ')
        p.add_run('{{ 銷售總額 }}').bold = True
        p.add_run('，較上月')
        p.add_run('{{ 銷售增長率 }}').bold = True
        doc.add_paragraph('系統將自動計算並填充這些變數的值。')
        doc.save(sample_path)
    
    return send_file(sample_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port="8000")