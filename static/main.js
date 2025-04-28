function showLoading() {
    document.getElementById('loadingOverlay').style.display = 'flex';
}

function hideLoading() {
    document.getElementById('loadingOverlay').style.display = 'none';
}

let currentFilteredData = [];    // 篩選後的小資料集
let formulas = {};               // 每個變數的公式或圖表設定
let currentVariable = '';        // 當前選擇的變數
let startDateInput = '';
let endDateInput = '';

// 輔助函數 - 取得圖表類型中文名稱
function getChartTypeName(chartType) {
    switch(chartType) {
        case 'line': return '折線圖';
        case 'bar': return '長條圖';
        case 'hist': return '直方圖';
        case 'pie': return '圓餅圖';
        default: return chartType;
    }
}

// 處理圖表切換顯示/隱藏
function setupChartToggles() {
    document.querySelectorAll('.chart-preview-toggle').forEach(toggle => {
        toggle.addEventListener('click', function() {
            const chartContainer = this.nextElementSibling;
            
            if (chartContainer.style.display === 'none' || !chartContainer.style.display) {
                chartContainer.style.display = 'block';
                this.innerHTML = '<i class="bi bi-chevron-up"></i> 收起圖表';
                this.classList.remove('btn-primary');
                this.classList.add('btn-outline-primary');
            } else {
                chartContainer.style.display = 'none';
                this.innerHTML = '<i class="bi bi-chevron-down"></i> 查看圖表';
                this.classList.remove('btn-outline-primary');
                this.classList.add('btn-primary');
            }
        });
    });
}

function generateChartPreview(varName, xAxis, yAxis, chartType, filteredData, chartTitle) {
    console.log(`【開始生成圖表】變數: ${varName}, X軸: ${xAxis}, Y軸: ${yAxis}, 類型: ${chartType}, 標題: "${chartTitle}"`);
    console.log(`formulas[${varName}]:`, JSON.stringify(formulas[varName]));
    const chartContainer = document.getElementById(`chart-${varName}`);
    chartContainer.innerHTML = '';

    if (!filteredData || filteredData.length === 0) {
        chartContainer.innerHTML = '<p class="text-danger">沒有資料可顯示圖表</p>';
        return;
    }

    const plotArea = document.createElement('div');
    plotArea.id = `plot-area-${varName}`;
    plotArea.style.width = '90%';
    plotArea.style.maxWidth = '800px';
    plotArea.style.margin = '0 auto';
    chartContainer.appendChild(plotArea);
    
    // 對標題進行特殊處理 - 徹底解決問題
    // 如果有提供標題就使用它，否則使用變數名稱
    // 注意：空字串 "" 應被視為有效標題（用戶想要清空標題）
    const finalTitle = (chartTitle !== undefined && chartTitle !== null) ? chartTitle : varName;
    
    console.log(`處理變數 ${varName} 的圖表，標題: "${finalTitle}"`);
    console.log(`圖表標題參數 (${varName}): "${chartTitle}"`);
    console.log(`最終設定的標題: "${finalTitle}"`);

    let plotData = [];
    
    // 創建 layout 時徹底修改標題設定方式
    let layout = {
        margin: { t: 80, b: 60, l: 60, r: 60 },
        height: 400,
        autosize: true
    };
    
    // 只有當標題非空時才設置標題
    if (finalTitle !== "") {
        layout.title = {
            text: finalTitle,
            font: {
                size: 20
            },
            x: 0.5,
            xanchor: 'center'
        };
    }
    
    // 記錄 layout 設定
    console.log('Plotly 使用的 layout:', JSON.stringify(layout));

    if (chartType === 'line') {
        const grouped = {};
        filteredData.forEach(row => {
            const dateStr = new Date(row[xAxis]).toISOString().slice(0, 10);
            grouped[dateStr] = (grouped[dateStr] || 0) + (Number(row[yAxis]) || 0);
        });
        const x = Object.keys(grouped).sort();
        const y = x.map(date => grouped[date]);
        plotData = [{ x, y, mode: 'lines+markers', type: 'scatter' }];
    } else if (chartType === 'bar') {
        const grouped = {};
        filteredData.forEach(row => {
            const key = row[xAxis];
            grouped[key] = (grouped[key] || 0) + (Number(row[yAxis]) || 0);
        });
        const x = Object.keys(grouped);
        const y = x.map(k => grouped[k]);
        plotData = [{ x, y, type: 'bar' }];
    } else if (chartType === 'hist') {
        const values = filteredData.map(row => Number(row[yAxis])).filter(v => !isNaN(v));
        plotData = [{ x: values, type: 'histogram', nbinsx: 20 }];
    } else if (chartType === 'pie') {
        const grouped = {};
        filteredData.forEach(row => {
            const key = row[xAxis];
            grouped[key] = (grouped[key] || 0) + (Number(row[yAxis]) || 0);
        });
        const labels = Object.keys(grouped);
        const values = labels.map(k => grouped[k]);
        plotData = [{ labels, values, type: 'pie' }];
    } else {
        chartContainer.innerHTML = '<p class="text-danger">不支援的圖表類型</p>';
        return;
    }

    const config = {
        responsive: true,
        displaylogo: false,
        modeBarButtonsToAdd: ['toImage'],
        toImageButtonOptions: {
            format: 'png',
            filename: varName,
            height: 600,
            width: 800,
            scale: 2
        }
    };
    
    // 清除先前的圖表
    Plotly.purge(plotArea);
    
    // 繪製新圖表，確保標題已正確設定
    Plotly.newPlot(plotArea, plotData, layout, config);
}

// 新增函數：設置快速修改按鈕功能
function setupQuickEditButtons() {
    document.querySelectorAll('.quick-edit-chart-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            const varName = this.getAttribute('data-variable');
            if (formulas[varName] && formulas[varName].type === 'chart') {
                currentVariable = varName;
                openQuickEditModal(varName, formulas[varName]);
            }
        });
    });
}

// 新增函數：快速修改彈出窗口
// 新增函數：快速修改彈出窗口
function openQuickEditModal(varName, chartSetting) {
    // 創建快速修改的 Modal HTML
    const modalHTML = `
        <div class="modal fade" id="quickEditModal" tabindex="-1" aria-labelledby="quickEditModalLabel" aria-hidden="true">
            <div class="modal-dialog modal-dialog-centered">
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title" id="quickEditModalLabel">
                            <i class="bi bi-graph-up me-2"></i>快速修改圖表
                        </h5>
                        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="關閉"></button>
                    </div>
                    <div class="modal-body">
                        <div class="mb-3">
                            <label for="quickEditXAxis" class="form-label">X 軸欄位</label>
                            <select id="quickEditXAxis" class="form-select"></select>
                        </div>
                        <div class="mb-3">
                            <label for="quickEditYAxis" class="form-label">Y 軸欄位</label>
                            <select id="quickEditYAxis" class="form-select"></select>
                        </div>
                        <div class="mb-3">
                            <label for="quickEditChartTitle" class="form-label">圖表標題</label>
                            <input type="text" id="quickEditChartTitle" class="form-control" placeholder="請輸入圖表標題">
                        </div>
                        <div class="mb-3">
                            <label for="quickEditChartType" class="form-label">圖表類型</label>
                            <select id="quickEditChartType" class="form-select">
                                <option value="line">折線圖（Line）</option>
                                <option value="bar">長條圖（Bar）</option>
                                <option value="hist">直方圖（Histogram）</option>
                                <option value="pie">圓餅圖（Pie）</option>
                            </select>
                        </div>
                        <div class="mb-3">
                            <label for="quickEditDpi" class="form-label">輸出解析度 (DPI)</label>
                            <select id="quickEditDpi" class="form-select">
                                <option value="1">預設</option>
                                <option value="1.5">150 DPI</option>
                                <option value="3">300 DPI</option>
                                <option value="6">600 DPI</option>
                            </select>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">取消</button>
                        <button type="button" class="btn btn-primary" id="saveQuickEditBtn">套用修改</button>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    // 如果頁面上已經有 quickEditModal 則移除
    const existingModal = document.getElementById('quickEditModal');
    if (existingModal) {
        existingModal.remove();
    }
    
    // 添加 Modal 到頁面
    const modalContainer = document.createElement('div');
    modalContainer.innerHTML = modalHTML;
    document.body.appendChild(modalContainer.firstElementChild);
    
    // 載入 X、Y 軸選項
    fetch('/get_columns')
        .then(response => response.json())
        .then(data => {
            const xAxisSelect = document.getElementById('quickEditXAxis');
            const yAxisSelect = document.getElementById('quickEditYAxis');
            xAxisSelect.innerHTML = '';
            yAxisSelect.innerHTML = '';

            if (data.columns && data.columns.length > 0) {
                data.columns.forEach(col => {
                    const xOption = document.createElement('option');
                    xOption.value = col;
                    xOption.textContent = col;
                    xAxisSelect.appendChild(xOption);

                    const yOption = document.createElement('option');
                    yOption.value = col;
                    yOption.textContent = col;
                    yAxisSelect.appendChild(yOption);
                });
                
                // 設置當前值
                xAxisSelect.value = chartSetting.x;
                yAxisSelect.value = chartSetting.y;
                document.getElementById('quickEditChartType').value = chartSetting.chartType;
                
                // 修正：正確設定圖表標題到輸入框
                const titleInput = document.getElementById('quickEditChartTitle');
                titleInput.value = chartSetting.chartTitle || '';
                
                // 記錄用於除錯
                console.log(`打開快速編輯模態視窗: ${varName}, 載入標題: "${chartSetting.chartTitle}"`);
            }
        });
    
    // 顯示 Modal
    const quickEditModal = new bootstrap.Modal(document.getElementById('quickEditModal'));
    quickEditModal.show();
    
    // 綁定保存按鈕事件
    document.getElementById('saveQuickEditBtn').addEventListener('click', function() {
        const newXAxis = document.getElementById('quickEditXAxis').value;
        const newYAxis = document.getElementById('quickEditYAxis').value;
        const newChartType = document.getElementById('quickEditChartType').value;
        
        // 獲取標題輸入框的值，但不進行 trim 處理
        const newChartTitle = document.getElementById('quickEditChartTitle').value;
        
        if (!newXAxis || !newYAxis) {
            alert('請選擇 X 軸和 Y 軸！');
            return;
        }
    
        // 除錯：記錄標題值
        console.log(`儲存按鈕 - 變數: ${currentVariable}, 設定標題值: "${newChartTitle}"`);
    
        // 更新 formulas 物件
        formulas[currentVariable] = {
            type: 'chart',
            x: newXAxis,
            y: newYAxis,
            chartType: newChartType,
            chartTitle: newChartTitle  // 直接使用輸入值，不進行加工
        };
    
        // 除錯：確認 formulas 已更新
        console.log(`formulas[${currentVariable}] 已更新:`, JSON.stringify(formulas[currentVariable]));
    
        // 呼叫後端重新生成圖
        fetch('/regenerate_chart', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                varName: currentVariable,
                x: newXAxis,
                y: newYAxis,
                chartType: newChartType,
                chartTitle: newChartTitle,  // 確保標題被傳遞到後端
                dpi: parseFloat(document.getElementById('quickEditDpi').value), 
                data: currentFilteredData.map(row => ({
                    [newXAxis]: row[newXAxis],
                    [newYAxis]: row[newYAxis]
                }))
            })
        })
        .then(response => response.json())
        .then(result => {
            if (result.error) {
                alert('❌ 重繪圖表失敗: ' + result.error);
            } else {
                alert('✅ 成功重新產生圖表');
                // 關閉 Modal
                const quickEditModal = bootstrap.Modal.getInstance(document.getElementById('quickEditModal'));
                quickEditModal.hide();
                
                // 重新計算並渲染所有內容
                calculateAndRender();
            }
        })
        .catch(err => {
            console.error('快速重繪失敗', err);
        });
    });    
}

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
        alert(`成功套用篩選條件！目前資料量：${data.length} 筆`);
        document.getElementById('generateSection').style.display = 'block';

        // 更新 start_date, end_date 的顯示
        const startSpan = document.querySelector('[data-variable="start_date"]');
        const endSpan = document.querySelector('[data-variable="end_date"]');
        if (startSpan) startSpan.textContent = startDateInput;
        if (endSpan) endSpan.textContent = endDateInput;

        // 儲存到 formulas，作為固定文字
        formulas['start_date'] = { type: 'fixed', value: `${startDateInput}` };
        formulas['end_date'] = { type: 'fixed', value: `${endDateInput}` };

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

        const setting = formulas[currentVariable];

        document.getElementById('typeFormula').checked = true;
        document.getElementById('formulaFields').style.display = 'block';
        document.getElementById('chartFields').style.display = 'none';
        document.getElementById('formulaInput').value = '';

        fetch('/get_columns')
            .then(response => response.json())
            .then(data => {
                const xAxisSelect = document.getElementById('xAxisSelect');
                const yAxisSelect = document.getElementById('yAxisSelect');
                const chartTypeSelect = document.getElementById('chartTypeSelect');
                
                xAxisSelect.innerHTML = '';
                yAxisSelect.innerHTML = '';

                if (data.columns && data.columns.length > 0) {
                    data.columns.forEach(col => {
                        const optionX = document.createElement('option');
                        optionX.value = col;
                        optionX.textContent = col;
                        xAxisSelect.appendChild(optionX);

                        const optionY = document.createElement('option');
                        optionY.value = col;
                        optionY.textContent = col;
                        yAxisSelect.appendChild(optionY);
                    });
                }

                if (setting) {
                    if (setting.type === 'chart') {
                        document.getElementById('typeChart').checked = true;
                        document.getElementById('formulaFields').style.display = 'none';
                        document.getElementById('chartFields').style.display = 'block';

                        xAxisSelect.value = setting.x || '';
                        yAxisSelect.value = setting.y || '';
                        chartTypeSelect.value = setting.chartType || 'line';
                    } else if (setting.type === 'formula') {
                        document.getElementById('typeFormula').checked = true;
                        document.getElementById('formulaFields').style.display = 'block';
                        document.getElementById('chartFields').style.display = 'none';

                        document.getElementById('formulaInput').value = setting.value || '';
                    }
                }
            });

        const modal = new bootstrap.Modal(document.getElementById('formulaModal'));
        modal.show();
    }
});

// 切換 公式 / 圖表 顯示不同欄位
document.getElementsByName('variableType').forEach(radio => {
    radio.addEventListener('change', function() {
        if (this.value === 'formula') {
            document.getElementById('formulaFields').style.display = 'block';
            document.getElementById('chartFields').style.display = 'none';
        } else {
            document.getElementById('formulaFields').style.display = 'none';
            document.getElementById('chartFields').style.display = 'block';
        }
    });
});

// ====== 3. 儲存公式設定，並即時計算並更新畫面 ======
document.getElementById('saveFormulaBtn').addEventListener('click', function() {
    const selectedType = document.querySelector('input[name="variableType"]:checked').value;

    if (!currentVariable) {
        alert('未選取變數，請重新操作');
        return;
    }

    if (selectedType === 'formula') {
        const formula = document.getElementById('formulaInput').value.trim();
        if (!formula) {
            alert('請輸入公式！');
            return;
        }
        formulas[currentVariable] = {
            type: 'formula',
            value: formula
        };
    } else if (selectedType === 'chart') {
        const xAxis = document.getElementById('xAxisSelect').value;
        const yAxis = document.getElementById('yAxisSelect').value;
        const chartType = document.getElementById('chartTypeSelect').value;
        const chartTitle = document.getElementById('chartTitleInput').value.trim();

        if (!xAxis || !yAxis || !chartType) {
            alert('請選擇完整的圖表設定 (X 軸, Y 軸, 圖表類型)！');
            return;
        }

        formulas[currentVariable] = {
            type: 'chart',
            x: xAxis,
            y: yAxis,
            chartType: chartType,
            chartTitle: chartTitle
        };
    }

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

    // 除錯用：顯示所有圖表設定
    Object.keys(formulas).forEach(key => {
        if (formulas[key].type === 'chart') {
            console.log(`圖表設定檢查 (${key}):`, JSON.stringify(formulas[key]));
        }
    });

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
            span.innerHTML = '';

            if (formulas[varName] && formulas[varName].type === 'chart') {
                const setting = formulas[varName];
                
                // 除錯用：記錄在處理圖表時的標題值
                console.log(`處理變數 ${varName} 的圖表，標題設定:`, JSON.stringify(setting));

                const filteredData = currentFilteredData
                    .filter(row => row[setting.x] !== undefined && row[setting.y] !== undefined)
                    .map(row => ({
                        [setting.x]: row[setting.x],
                        [setting.y]: row[setting.y]
                    }));

                const chartHTML = `
                    <div class="chart-container" style="width:1000px; height:600px; max-width:90vw; margin-left:auto; margin-right:auto;">
                        <button class="btn btn-outline-primary btn-sm chart-preview-toggle mb-2">
                            <i class="bi bi-chevron-up"></i> 收起圖表
                        </button>
                        <div class="chart-preview">
                            <div class="d-flex justify-content-between align-items-center mb-2">
                                <div class="chart-info d-flex gap-3">
                                    <div><strong>X 軸:</strong> ${setting.x}</div>
                                    <div><strong>Y 軸:</strong> ${setting.y}</div>
                                    <div><strong>類型:</strong> ${getChartTypeName(setting.chartType)}</div>
                                    ${setting.chartTitle ? `<div><strong>標題:</strong> ${setting.chartTitle}</div>` : ''}
                                </div>
                                <button class="btn btn-sm btn-outline-secondary quick-edit-chart-btn" data-variable="${varName}">
                                    <i class="bi bi-pencil"></i> 快速修改
                                </button>
                            </div>
                            <div id="chart-${varName}" class="chart-image-container" style="width:100%; height:300px; background-color:#f8f9fa; border-radius:8px; display:flex; justify-content:center; align-items:center;">
                                <div class="text-center">
                                    <div class="spinner-border text-primary" role="status">
                                        <span class="visually-hidden">Loading...</span>
                                    </div>
                                    <p class="mt-2">正在生成圖表預覽...</p>
                                </div>
                            </div>
                        </div>
                    </div>
                `;
                span.innerHTML = chartHTML;

                // 關鍵修正：確保標題正確傳遞！
                generateChartPreview(
                    varName, 
                    setting.x, 
                    setting.y, 
                    setting.chartType, 
                    filteredData,
                    setting.chartTitle  // 直接傳遞 setting.chartTitle，不做任何處理
                );

            } else if (results[varName] !== undefined) {
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

        setupChartToggles();
        setupQuickEditButtons();
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
    showLoading();
    
    fetch('/save_settings', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            formulas: formulas
        })
    })
    .then(response => {
        hideLoading();
        if (response.ok) {
            alert('✅ 設定已成功儲存！');
            calculateAndRender();
        } else {
            alert('❌ 儲存失敗！');
        }
    })
    .catch(err => {
        hideLoading();
        console.error('儲存設定失敗', err);
    });
});

// ====== 7. 載入設定檔 settings.json ======
document.getElementById('loadSettingsBtn').addEventListener('click', function() {
    showLoading();
    
    fetch('/load_settings')
    .then(response => {
        if (!response.ok) throw new Error('找不到設定檔');
        return response.json();
    })
    .then(data => {
        hideLoading();
        if (data.formulas) {
            formulas = data.formulas;
            alert('✅ 成功載入公式設定！');
            calculateAndRender();
        } else {
            alert('⚠️ 設定檔格式錯誤');
        }
    })
    .catch(err => {
        hideLoading();
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
    
    // 確保在一開始就設置好圖表切換功能（如果頁面上已有圖表）
    setupChartToggles();
});