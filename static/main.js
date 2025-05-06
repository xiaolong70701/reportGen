function showLoading() {
    document.getElementById('loadingOverlay').style.display = 'flex';
}

function hideLoading() {
    document.getElementById('loadingOverlay').style.display = 'none';
}

function showNotification(message, type = 'info', delay = 3000) { // type 可以是 success, warning, danger, info 等 Bootstrap 背景色
    // 建立通知元素
    const notification = document.createElement('div');
    // 確保使用有效的 Bootstrap 背景色類別
    const validTypes = ['primary', 'secondary', 'success', 'danger', 'warning', 'info', 'light', 'dark'];
    const bgType = validTypes.includes(type) ? type : 'info'; // 預設為 info
    const textClass = (bgType === 'light' || bgType === 'warning') ? 'text-dark' : 'text-white'; // 淺色和警告色背景用深色文字

    notification.className = `toast align-items-center ${textClass} bg-${bgType} border-0`;
    notification.setAttribute('role', 'alert');
    notification.setAttribute('aria-live', 'assertive');
    notification.setAttribute('aria-atomic', 'true');

    // 判斷圖標
    let iconClass = 'bi-info-circle-fill';
    switch(type) {
        case 'success': iconClass = 'bi-check-circle-fill'; break;
        case 'warning': iconClass = 'bi-exclamation-triangle-fill'; break;
        case 'danger': iconClass = 'bi-x-octagon-fill'; break;
    }

    // 建立通知內容
    notification.innerHTML = `
        <div class="d-flex">
            <div class="toast-body">
                <i class="bi ${iconClass} me-2"></i>
                ${message}
            </div>
            <button type="button" class="btn-close btn-close-${textClass === 'text-white' ? 'white' : 'dark'} me-2 m-auto" data-bs-dismiss="toast" aria-label="關閉"></button>
        </div>
    `;

    // 建立或取得通知容器
    let toastContainer = document.querySelector('.toast-container');
    if (!toastContainer) {
        toastContainer = document.createElement('div');
        // 確保容器有 position-fixed 或 position-absolute 以便顯示在頂層
        toastContainer.className = 'toast-container position-fixed top-0 end-0 p-3';
        toastContainer.style.zIndex = '1100'; // 確保在 loading overlay 之上 (如果需要) 或其他元素之上
        document.body.appendChild(toastContainer);
    }

    // 添加通知到容器
    toastContainer.appendChild(notification);

    // 顯示通知
    const toast = new bootstrap.Toast(notification, {
        delay: delay // 使用傳入的延遲時間
    });
    toast.show();

    // 當 Toast 關閉時，從 DOM 中移除，避免累積
    notification.addEventListener('hidden.bs.toast', function () {
        notification.remove();
    });
}

let currentFilteredData = [];    //Filtered data set
let formulas = {};               //Formula or chart setting for each variable
let currentVariable = '';        //The currently selected variable
let startDateInput = '';
let endDateInput = '';

//Helper function -Get the Chinese name of the chart type
function getChartTypeName(chartType) {
    switch(chartType) {
        case 'line': return '折線圖';
        case 'bar': return '長條圖';
        case 'hist': return '直方圖';
        case 'pie': return '圓餅圖';
        default: return chartType;
    }
}

//Handle chart switching to show/hide
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

    const finalTitle = (chartTitle !== undefined && chartTitle !== null) ? chartTitle : varName;

    let plotData = [];
    
    //Completely modify the title setting method when creating layout
    let layout = {
        margin: { t: 80, b: 60, l: 60, r: 60 },
        height: 400,
        autosize: true
    };
    
    //Set the title only if it is not empty
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
    
    //Clear previous chart
    Plotly.purge(plotArea);
    
    //Draw a new chart to make sure the title is set correctly
    Plotly.newPlot(plotArea, plotData, layout, config);
}

//New function: Set the quick modification button function
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

//New function: Quickly modify pop-up window
//New function: Quickly modify pop-up window
function openQuickEditModal(varName, chartSetting) {
    //Create fast modified Modal HTML
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
    
    //If there is already quickEditModal on the page, remove it
    const existingModal = document.getElementById('quickEditModal');
    if (existingModal) {
        existingModal.remove();
    }
    
    //Add Modal to the page
    const modalContainer = document.createElement('div');
    modalContainer.innerHTML = modalHTML;
    document.body.appendChild(modalContainer.firstElementChild);
    
    //Load X and Y axis options
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
                
                //Set the current value
                xAxisSelect.value = chartSetting.x;
                yAxisSelect.value = chartSetting.y;
                document.getElementById('quickEditChartType').value = chartSetting.chartType;
                
                //Fixed: Correctly set the chart title to the input box
                const titleInput = document.getElementById('quickEditChartTitle');
                titleInput.value = chartSetting.chartTitle || '';
            }
        });
    
    //Show Modal
    const quickEditModal = new bootstrap.Modal(document.getElementById('quickEditModal'));
    quickEditModal.show();
    
    //Bind Save Button Event
    document.getElementById('saveQuickEditBtn').addEventListener('click', function() {
        const newXAxis = document.getElementById('quickEditXAxis').value;
        const newYAxis = document.getElementById('quickEditYAxis').value;
        const newChartType = document.getElementById('quickEditChartType').value;
        
        //Get the value of the title input box, but does not perform trim processing
        const newChartTitle = document.getElementById('quickEditChartTitle').value;
        
        if (!newXAxis || !newYAxis) {
            alert('請選擇 X 軸和 Y 軸！');
            return;
        }
    
        //Update formulas objects
        formulas[currentVariable] = {
            type: 'chart',
            x: newXAxis,
            y: newYAxis,
            chartType: newChartType,
            chartTitle: newChartTitle  //Use the input value directly without processing
        };
    
        //Regenerate the image of the call backend
        fetch('/regenerate_chart', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                varName: currentVariable,
                x: newXAxis,
                y: newYAxis,
                chartType: newChartType,
                chartTitle: newChartTitle,  //Make sure the title is passed to the backend
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
                //Close Modal
                const quickEditModal = bootstrap.Modal.getInstance(document.getElementById('quickEditModal'));
                quickEditModal.hide();
                
                //Recalculate and render everything
                calculateAndRender();
            }
        })
        .catch(err => {
            console.error('快速重繪失敗', err);
        });
    });    
}

//====== 1. Apply filter date =======
document.getElementById('filterForm').addEventListener('submit', function(e) {
    e.preventDefault();
    showLoading();

    startDateInput = document.getElementById('startDate').value;
    endDateInput = document.getElementById('endDate').value;

    if (!startDateInput || !endDateInput) {
        hideLoading();
        // alert('請選擇起始和結束日期！');
        showNotification('請選擇起始和結束日期！', 'warning');
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
        // alert(`成功套用篩選條件！目前資料量：${data.length} 筆`);
        showNotification(`成功套用篩選條件！目前資料量：${data.length} 筆`, 'primary');
        document.getElementById('generateSection').style.display = 'block';

        //Update the display of start_date, end_date
        const startSpan = document.querySelector('[data-variable="start_date"]');
        const endSpan = document.querySelector('[data-variable="end_date"]');
        if (startSpan) startSpan.textContent = startDateInput;
        if (endSpan) endSpan.textContent = endDateInput;

        //Save to formulas as fixed text
        formulas['start_date'] = { type: 'fixed', value: `${startDateInput}` };
        formulas['end_date'] = { type: 'fixed', value: `${endDateInput}` };

        //Automatic calculation after filtering is completed
        calculateAndRender();
    })
    .catch(err => {
        hideLoading();
        showNotification('篩選資料失敗，請檢查網路或後端服務。', 'danger');
        console.error('篩選資料失敗', err);
    });
});

//===== 2. Click on the variable to open the small window for setting formulas =======
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

//Switch Formula/Graph Show different columns
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

//===== 3. Save formula settings, calculate and update the screen instantly ======
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

//===== 4. The real "instant calculation + update preview" ======
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

    //Default: Show all chart settings
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
                
                //Decryption: Record the title value when processing the chart
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

                //Key correction: Make sure the title is delivered correctly!
                generateChartPreview(
                    varName, 
                    setting.x, 
                    setting.y, 
                    setting.chartType, 
                    filteredData,
                    setting.chartTitle  //Pass setting.chartTitle directly without any processing
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

//===== 5. The last button generates a Word report =======
document.getElementById('generateForm').addEventListener('submit', function(e) {
    e.preventDefault();
    showLoading();

    const filename = document.getElementById('docxFileName').value.trim() || 'weekly_report.docx';

    fetch('/render', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            formulas: formulas,
            data: currentFilteredData,
            filename: filename
        })
    })
    .then(response => response.blob())
    .then(blob => {
        hideLoading();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        const filenameInput = document.getElementById('docxFileName').value.trim() || 'weekly_report.docx';
        let finalFilename = filenameInput;
        if (!finalFilename.toLowerCase().endsWith('.docx')) {
            finalFilename += '.docx';  //Automatically fill in .docx
        }
        a.download = finalFilename;
        document.body.appendChild(a);
        a.click();
        a.remove();
    })
    .catch(err => {
        hideLoading();
        console.error('產生 Word 失敗', err);
    });
});

//===== 6. Set the storage formula to settings.json ======
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

//====== 7. Load settings.json ======
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

//====== 8. As soon as the page comes in, it automatically tries to load settings.json ======
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
    
    //Make sure to set the chart switching function at the beginning (if there is a chart on the page)
    setupChartToggles();
});