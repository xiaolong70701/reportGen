// static/main.js

function showLoading() {
    document.getElementById('loadingOverlay').style.display = 'flex';
}

function hideLoading() {
    document.getElementById('loadingOverlay').style.display = 'none';
}

let currentFilteredData = [];    // 篩選後的小資料集
let formulas = {};               // 每個變數的公式
let currentVariable = '';        // 當前選擇的變數
let startDateInput = '';
let endDateInput = '';

// ====== 1. 套用篩選日期 ======
document.getElementById('filterForm').addEventListener('submit', function(e) {
    e.preventDefault();
    showLoading();

    startDateInput = document.getElementById('startDate').value;
    endDateInput = document.getElementById('endDate').value;

    if (!startDateInput || !endDateInput) {
        hideLoading();
        alert('請選擇起始和結束日期！');
        return;
    }

    fetch('/filter_data', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ start_date: startDateInput, end_date: endDateInput })
    })
    .then(response => response.json())
    .then(data => {
        currentFilteredData = data;
        hideLoading();
        alert(`✅ 成功套用篩選條件！目前資料量：${data.length} 筆`);
        document.getElementById('generateSection').style.display = 'block';

        // 更新 start_date, end_date 的顯示
        const startSpan = document.querySelector('[data-variable="start_date"]');
        const endSpan = document.querySelector('[data-variable="end_date"]');
        if (startSpan) startSpan.textContent = startDateInput;
        if (endSpan) endSpan.textContent = endDateInput;

        // 儲存到 formulas，作為固定文字
        formulas['start_date'] = `'${startDateInput}'`;
        formulas['end_date'] = `'${endDateInput}'`;

        // 篩選完成後自動計算
        calculateAndRender();
    })
    .catch(err => {
        hideLoading();
        console.error('篩選資料失敗', err);
    });
});

// ====== 2. 點擊變數打開設定公式小視窗 ======
document.addEventListener('click', function(e) {
    if (e.target.classList.contains('editable')) {
        currentVariable = e.target.getAttribute('data-variable');
        document.getElementById('variableName').value = currentVariable;
        document.getElementById('formulaInput').value = formulas[currentVariable] || '';

        const modal = new bootstrap.Modal(document.getElementById('formulaModal'));
        modal.show();
    }
});

// ====== 3. 儲存公式設定，並即時計算並更新畫面 ======
document.getElementById('saveFormulaBtn').addEventListener('click', function() {
    const formula = document.getElementById('formulaInput').value.trim();

    if (!formula) {
        alert('請輸入公式！');
        return;
    }

    formulas[currentVariable] = formula;

    const modal = bootstrap.Modal.getInstance(document.getElementById('formulaModal'));
    modal.hide();

    calculateAndRender();
});

// ====== 4. 真正的「即時計算＋更新預覽」 ======
function calculateAndRender() {
    if (currentFilteredData.length === 0) {
        console.warn('尚未篩選資料！');
        return;
    }
    if (Object.keys(formulas).length === 0) {
        console.warn('尚未設定公式！');
        return;
    }

    showLoading();

    fetch('/render_preview', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            formulas: formulas,
            data: currentFilteredData
        })
    })
    .then(response => response.json())
    .then(results => {
        hideLoading();
        document.querySelectorAll('.variable').forEach(span => {
            const varName = span.getAttribute('data-variable');
            if (results[varName] !== undefined) {
                const value = results[varName];

                if (typeof value === 'string' && value.startsWith('錯誤')) {
                    span.textContent = value;
                    span.classList.add('error');
                    span.setAttribute('title', value);
                } else {
                    span.textContent = value;
                    span.classList.remove('error');
                    span.removeAttribute('title');
                }
            }
        });
    })
    .catch(err => {
        hideLoading();
        console.error('即時計算失敗', err);
    });
}

// ====== 5. 最後一鍵產生 Word 報告 ======
document.getElementById('generateForm').addEventListener('submit', function(e) {
    e.preventDefault();
    showLoading();

    fetch('/render', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            formulas: formulas,
            data: currentFilteredData
        })
    })
    .then(response => response.blob())
    .then(blob => {
        hideLoading();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = "weekly_report.docx";
        document.body.appendChild(a);
        a.click();
        a.remove();
    })
    .catch(err => {
        hideLoading();
        console.error('產生 Word 失敗', err);
    });
});

// ====== 6. 儲存公式設定成 settings.json ======
document.getElementById('saveSettingsBtn').addEventListener('click', function() {
    fetch('/save_settings', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            formulas: formulas
        })
    })
    .then(response => {
        if (response.ok) {
            alert('✅ 設定已成功儲存！');
            calculateAndRender();
        } else {
            alert('❌ 儲存失敗！');
        }
    })
    .catch(err => {
        console.error('儲存設定失敗', err);
    });
});

// ====== 7. 載入設定檔 settings.json ======
document.getElementById('loadSettingsBtn').addEventListener('click', function() {
    fetch('/load_settings')
    .then(response => {
        if (!response.ok) throw new Error('找不到設定檔');
        return response.json();
    })
    .then(data => {
        if (data.formulas) {
            formulas = data.formulas;
            alert('✅ 成功載入公式設定！');
            calculateAndRender();
        } else {
            alert('⚠️ 設定檔格式錯誤');
        }
    })
    .catch(err => {
        console.error('載入設定失敗', err);
        alert('❌ 載入失敗！');
    });
});

// ====== 8. 頁面一進來自動嘗試載入 settings.json ======
window.addEventListener('DOMContentLoaded', function() {
    fetch('/load_settings')
    .then(response => {
        if (!response.ok) {
            console.log('ℹ️ 無設定檔，跳過');
            return;
        }
        return response.json();
    })
    .then(data => {
        if (data && data.formulas) {
            formulas = data.formulas;
            console.log('✅ 自動載入設定');

            setTimeout(() => {
                document.getElementById('filterForm').dispatchEvent(new Event('submit'));
            }, 300);
        }
    })
    .catch(err => {
        console.log('ℹ️ 沒有設定檔，不處理');
    });
});
