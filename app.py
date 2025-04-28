import io
import base64
import os
import re
import json
from flask import Flask, request, render_template, send_file, redirect, url_for, jsonify, flash, send_from_directory
import pandas as pd
import numpy as np
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import plotly.express as px      
from docxtpl import DocxTemplate
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

def generate_chart(df, x_col, y_col, chart_type, output_path, chart_title=None, dpi_scale=2):
    fig = None

    title = chart_title if chart_title else f"{y_col} by {x_col}"

    if chart_type == 'line':
        df[x_col] = pd.to_datetime(df[x_col], errors='coerce')
        daily_df = df.groupby(df[x_col].dt.date).agg({y_col: 'sum'}).reset_index()
        daily_df.rename(columns={daily_df.columns[0]: x_col}, inplace=True)
        fig = px.line(daily_df, x=x_col, y=y_col, title=title, template='simple_white')
    
    elif chart_type == 'bar':
        bar_df = df.groupby(x_col).agg({y_col: 'sum'}).reset_index()
        fig = px.bar(bar_df, x=x_col, y=y_col, title=title, template='simple_white')
    
    elif chart_type == 'hist':
        fig = px.histogram(df, x=y_col, nbins=20, title=title, template='simple_white')
    
    elif chart_type == 'pie':
        pie_data = df.groupby(x_col)[y_col].sum().reset_index()
        fig = px.pie(pie_data, values=y_col, names=x_col, title=title)
    
    else:
        raise ValueError(f"不支援的圖表類型: {chart_type}")

    fig.update_layout(
        template='none',
        font_family=os.path.join("fonts", "cwTeXQYuan-Medium.ttf"),
        title_font_size=20,
        xaxis_title_font_size=16,
        yaxis_title_font_size=16,
        legend_title_font_size=16,
        font_size=14,
        paper_bgcolor='white',
        plot_bgcolor='white'
    )

    fig.write_image(output_path, width=800, height=600, scale=dpi_scale)

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
        html += f"<div>{line}</div>\n"

    return html

def evaluate_formula(formula: str, df: pd.DataFrame, context: dict = None, formulas: dict = None) -> float:
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
                    next_formula = formulas[var]
                    if isinstance(next_formula, dict) and next_formula.get('type') == 'formula':
                        next_formula = next_formula.get('value', '')
                    context[var] = evaluate_formula(next_formula, df, context, formulas)
                    # print(f"已計算變數 {var} = {context[var]}")
                except Exception as e:
                    # print(f"計算變數 {var} 時發生錯誤: {str(e)}")
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

        eval_globals = {"df": df, "np": np}
        eval_globals.update(context)

        result = eval(formula, eval_globals)
        if any(name in formula for name in ["top_region", "top_product"]):
            print(f"⚡ DEBUG: 變數評估後結果 type={type(result)}, value={result}")

        if isinstance(result, (list, pd.Series)):
            if len(result) > 0:
                result = result[0]
            else:
                result = None
        elif isinstance(result, dict):
            if result:
                result = list(result.values())[0]
            else:
                result = None
        return result

    except Exception as e:
        raise ValueError(f"公式錯誤：{str(e)}")

# 新增首頁路由
@app.route('/')
def home():
    message = request.args.get('message', '')
    return render_template('home.html', message=message)

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
    special_vars = ['start_date', 'end_date']

    # 建立公式依賴關係圖
    formula_graph = {}
    for var, formula_info in formulas.items():
        if isinstance(formula_info, dict) and formula_info.get('type') == 'formula':
            formula = formula_info.get('value', '')
        elif isinstance(formula_info, dict) and formula_info.get('type') == 'chart':
            formula = ''
        else:
            formula = formula_info

        dependencies = set()
        for other_var in formulas:
            if other_var != var and isinstance(formula, str) and re.search(r'\b' + re.escape(other_var) + r'\b', formula):
                dependencies.add(other_var)
        formula_graph[var] = dependencies

    visited = set()
    temp_visited = set()
    order = []

    def topo_sort(node):
        if node in temp_visited:
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

    order.reverse()

    for var in formulas:
        if var not in order:
            try:
                setting = formulas[var]
                if var in special_vars and 'date' in df.columns:
                    if var == 'start_date':
                        value = df['date'].min()
                    else:
                        value = df['date'].max()
                    if pd.isna(value):
                        value = ''
                    else:
                        value = pd.to_datetime(value).strftime('%Y/%m/%d')
                    results[var] = value
                    calculated_context[var] = value
                elif isinstance(setting, dict):
                    if setting.get('type') == 'formula':
                        expr = setting.get('value')
                        value = evaluate_formula(expr, df, context=calculated_context, formulas=formulas)
                        if hasattr(value, "item"):
                            value = value.item()
                        if isinstance(value, (int, float)):
                            value = round(value, 2)
                        results[var] = value
                        calculated_context[var] = value
                    elif setting.get('type') == 'chart':
                        results[var] = ''
                    else:
                        results[var] = setting.get('value', '')
                        calculated_context[var] = setting.get('value', '')
                else:
                    results[var] = setting
                    calculated_context[var] = setting
            except Exception as e:
                results[var] = f"錯誤: {str(e)}"

    for var in order:
        try:
            setting = formulas[var]
            if isinstance(setting, dict):
                if setting.get('type') == 'chart':
                    x_col = setting.get('x')
                    y_col = setting.get('y')
                    chart_type = setting.get('chartType')

                    if not all([x_col, y_col, chart_type]):
                        results[var] = '錯誤：缺少圖表設定'
                        continue

                    fig = None
                    if chart_type == 'line':
                        df[x_col] = pd.to_datetime(df[x_col], errors='coerce')
                        daily_df = df.groupby(df[x_col].dt.date).agg({y_col: 'sum'}).reset_index()
                        daily_df.rename(columns={daily_df.columns[0]: x_col}, inplace=True)
                        fig = px.line(daily_df, x=x_col, y=y_col, title=f"{y_col} by {x_col}")
                    
                    elif chart_type == 'bar':
                        bar_df = df.groupby(x_col).agg({y_col: 'sum'}).reset_index()
                        fig = px.bar(bar_df, x=x_col, y=y_col, title=f"{y_col} by {x_col}")
                    
                    elif chart_type == 'hist':
                        fig = px.histogram(df, x=y_col, nbins=20, title=f"{y_col} Histogram")
                    
                    elif chart_type == 'pie':
                        pie_data = df.groupby(x_col)[y_col].sum().reset_index()
                        fig = px.pie(pie_data, values=y_col, names=x_col, title=f"{y_col} 分佈")
                    
                    else:
                        results[var] = '錯誤：不支援的圖表類型'
                        continue

                    save_path = os.path.join('generated', f'{var}.png')
                    fig.write_image(save_path, width=800, height=600, scale=2)

                    results[var] = ''

                elif setting.get('type') == 'formula':
                    expr = setting.get('value')
                    value = evaluate_formula(expr, df, context=calculated_context, formulas=formulas)
                    if hasattr(value, "item"):
                        value = value.item()
                    if isinstance(value, (int, float)):
                        value = round(value, 2)
                    results[var] = value
                    calculated_context[var] = value

                else:
                    results[var] = setting.get('value', '')
                    calculated_context[var] = setting.get('value', '')
            else:
                results[var] = setting
                calculated_context[var] = setting

        except Exception as e:
            results[var] = f"錯誤: {str(e)}"

    return jsonify(results)

@app.route('/render', methods=['POST'])
def render_word():
    global cached_docx_path

    formulas = request.json.get('formulas', {})
    data = request.json.get('data', [])
    filename = request.json.get('filename', 'final_report.docx')

    if not data:
        return "資料錯誤", 400

    filtered_df = pd.DataFrame(data)
    context = {}
    calculated_context = {}

    # 計算公式的依賴關係
    formula_graph = {}
    for var, formula_info in formulas.items():
        if isinstance(formula_info, dict) and formula_info.get('type') == 'formula':
            formula = formula_info.get('value', '')
        else:
            formula = formula_info
        dependencies = set()
        for other_var in formulas:
            if other_var != var and isinstance(formula, str) and re.search(r'\b' + re.escape(other_var) + r'\b', formula):
                dependencies.add(other_var)
        formula_graph[var] = dependencies

    # 拓撲排序
    visited = set()
    temp_visited = set()
    order = []

    def topo_sort(node):
        if node in temp_visited:
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

    order.reverse()
    print(f"公式計算順序: {order}")

    doc = DocxTemplate(cached_docx_path)

    # 先計算沒有依賴的公式
    for var, setting in formulas.items():
        if var not in order:
            try:
                if isinstance(setting, dict):
                    if setting.get('type') == 'formula':
                        expr = setting.get('value', '')
                        value = evaluate_formula(expr, filtered_df, context=calculated_context, formulas=formulas)
                        if hasattr(value, "item"):
                            value = value.item()
                        if isinstance(value, (int, float)):
                            value = round(value, 2)
                        context[var] = value
                        calculated_context[var] = value
                    elif setting.get('type') == 'chart':
                        x_col = setting.get('x')
                        y_col = setting.get('y')
                        chart_type = setting.get('chartType')
                        chart_title = setting.get('chartTitle')
                        if not x_col or not y_col or not chart_type:
                            continue
                        img_path = os.path.join(GENERATED_FOLDER, f"{var}.png")
                        generate_chart(filtered_df, x_col, y_col, chart_type, chart_title=chart_title, output_path=img_path)
                        context[var] = InlineImage(doc, img_path, width=Mm(120))
                else:
                    # 舊格式
                    expr = setting
                    value = evaluate_formula(expr, filtered_df, context=calculated_context, formulas=formulas)
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
            setting = formulas[var]
            if isinstance(setting, dict):
                if setting.get('type') == 'formula':
                    expr = setting.get('value', '')
                    value = evaluate_formula(expr, filtered_df, context=calculated_context, formulas=formulas)
                    if hasattr(value, "item"):
                        value = value.item()
                    if isinstance(value, (int, float)):
                        value = round(value, 2)
                    context[var] = value
                    calculated_context[var] = value
                elif setting.get('type') == 'chart':
                    x_col = setting.get('x')
                    y_col = setting.get('y')
                    chart_type = setting.get('chartType')
                    chart_title = setting.get('chartTitle')
                    if not x_col or not y_col or not chart_type:
                        continue
                    img_path = os.path.join(GENERATED_FOLDER, f"{var}.png")
                    generate_chart(filtered_df, x_col, y_col, chart_type, chart_title=chart_title, output_path=img_path)
                    context[var] = InlineImage(doc, img_path, width=Mm(120))
            else:
                expr = setting
                value = evaluate_formula(expr, filtered_df, context=calculated_context, formulas=formulas)
                if hasattr(value, "item"):
                    value = value.item()
                if isinstance(value, (int, float)):
                    value = round(value, 2)
                context[var] = value
                calculated_context[var] = value
        except Exception as e:
            context[var] = f"錯誤: {str(e)}"

    doc.render(context)

    # output_path = os.path.join(GENERATED_FOLDER, 'final_report.docx')
    output_path = os.path.join(GENERATED_FOLDER, filename)
    doc.save(output_path)

    return send_file(output_path, as_attachment=True, download_name=filename)

@app.route('/save_settings', methods=['POST'])
def save_settings():
    settings = request.json
    formulas = settings.get('formulas', {})

    # 自動正規化 formulas
    new_formulas = {}
    for var, setting in formulas.items():
        if isinstance(setting, dict) and 'type' in setting:
            new_formulas[var] = setting
        else:
            # 舊格式，補成 type=formula
            new_formulas[var] = {
                "type": "formula",
                "value": setting
            }

    # 重新組成完整 settings
    new_settings = {
        "formulas": new_formulas
    }

    with open(SETTINGS_PATH, 'w', encoding='utf-8') as f:
        json.dump(new_settings, f, ensure_ascii=False, indent=2)
    return "設定已儲存", 200

@app.route('/load_settings', methods=['GET'])
def load_settings():
    if not os.path.exists(SETTINGS_PATH):
        return jsonify({}), 404

    with open(SETTINGS_PATH, 'r', encoding='utf-8') as f:
        settings = json.load(f)

    formulas = settings.get('formulas', {})
    new_formulas = {}
    for var, setting in formulas.items():
        if isinstance(setting, dict) and 'type' in setting:
            new_formulas[var] = setting
        else:
            # 舊格式，補成 type=formula
            new_formulas[var] = {
                "type": "formula",
                "value": setting
            }

    settings['formulas'] = new_formulas
    return jsonify(settings)

@app.route('/get_columns', methods=['GET'])
def get_columns():
    global cached_dataframe
    if cached_dataframe is not None:
        return jsonify({"columns": list(cached_dataframe.columns)})
    else:
        return jsonify({"columns": []})

@app.route('/generated/<path:filename>')
def serve_generated_file(filename):
    return send_from_directory('generated', filename)

@app.route('/regenerate_chart', methods=['POST'])
def regenerate_chart():
    content = request.get_json()

    var_name = content.get('varName')
    x_col = content.get('x')
    y_col = content.get('y')
    chart_type = content.get('chartType')
    chart_title = content.get('chartTitle')
    data = content.get('data')
    dpi_scale = content.get('dpi', 2)  # 🔥 預設 dpi_scale 2

    if not all([var_name, x_col, y_col, chart_type, data]):
        return jsonify({'error': '缺少必要參數'}), 400

    df = pd.DataFrame(data)

    if df.empty:
        return jsonify({'error': '沒有資料'}), 400

    try:
        save_path = os.path.join('generated', f'{var_name}.png')
        generate_chart(df, x_col, y_col, chart_type, save_path, chart_title=chart_title, dpi_scale=dpi_scale)  # 🔥 使用上面的 generate_chart
        return jsonify({'message': '重新產生完成'})
    except Exception as e:
        return jsonify({'error': str(e)})

@app.route('/get_chart_preview', methods=['POST'])
def get_chart_preview():
    content = request.get_json()
    var_name = content.get('varName')  # 🔥 傳進來變數名
    
    if not var_name:
        return jsonify({'error': '缺少變數名稱'}), 400

    filename = f"{var_name}.png"  # 圖片檔案命名規則
    filepath = os.path.join('generated', filename)

    if not os.path.exists(filepath):
        return jsonify({'error': '找不到圖片'}), 404

    # 🔥 回傳圖片 URL，而不是 base64
    return jsonify({'image_url': url_for('static', filename=f'../generated/{filename}')})


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