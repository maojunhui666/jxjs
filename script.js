document.addEventListener('DOMContentLoaded', function() {
    const workOrderType = document.getElementById('workOrderType');
    const modificationFields = document.getElementById('modificationFields');
    const modificationItems = document.getElementById('modificationItems');
    const addModificationItem = document.getElementById('addModificationItem');
    const externalDevicesFields = document.getElementById('externalDevicesFields');
    const externalDeviceItems = document.getElementById('externalDeviceItems');
    const addExternalDevice = document.getElementById('addExternalDevice');
    const specialInstallationFields = document.getElementById('specialInstallationFields');
    const specialInstrumentType = document.getElementById('specialInstrumentType');
    const maintenanceFields = document.getElementById('maintenanceFields');
    const maintenanceType = document.getElementById('maintenanceType');
    const calculateButton = document.getElementById('calculate');
    const resultDiv = document.getElementById('result');
    const fileInput = document.getElementById('excelFile');
    const uploadStatus = document.getElementById('uploadStatus');
    const performanceRatio = document.getElementById('performanceRatio');
    const performanceRatioValue = document.getElementById('performanceRatioValue');
    const specialBonus = document.getElementById('specialBonus');
    const specialBonusValue = document.getElementById('specialBonusValue');
    const daysFactor = document.getElementById('daysFactor');
    const paramsToggle = document.getElementById('paramsToggle');
    const paramsCollapseEl = document.getElementById('paramsCollapse');
    let instrumentData = []; // 存储所有仪器配置数据
    const toggleFormBtn = document.getElementById('toggleFormBtn');
    const formContainer = document.getElementById('formContainer');

    // 初始化显示或隐藏字段
    modificationFields.classList.add('hidden');
    externalDevicesFields.classList.add('hidden');
    specialInstallationFields.classList.add('hidden');
    maintenanceFields.classList.add('hidden');

    // 根据默认工作单类型显示相应字段
    if (workOrderType.value === 'installation') {
        externalDevicesFields.classList.remove('hidden');
    }

    // 初始化参数调整折叠状态（默认收起），并记忆状态
    (function initParamsCollapse() {
        if (!paramsCollapseEl) return;
        const saved = localStorage.getItem('paramsCollapseOpen');
        const shouldOpen = saved === '1';
        const collapse = new bootstrap.Collapse(paramsCollapseEl, { toggle: false });
        if (shouldOpen) {
            collapse.show();
            if (paramsToggle) paramsToggle.setAttribute('aria-expanded', 'true');
        } else {
            collapse.hide();
            if (paramsToggle) paramsToggle.setAttribute('aria-expanded', 'false');
        }
        paramsCollapseEl.addEventListener('shown.bs.collapse', () => {
            localStorage.setItem('paramsCollapseOpen', '1');
            if (paramsToggle) paramsToggle.querySelector('i').className = 'bi bi-chevron-up';
        });
        paramsCollapseEl.addEventListener('hidden.bs.collapse', () => {
            localStorage.setItem('paramsCollapseOpen', '0');
            if (paramsToggle) paramsToggle.querySelector('i').className = 'bi bi-chevron-down';
        });
    })();
    
    // 滑块值实时显示
    performanceRatio.addEventListener('input', function() {
        performanceRatioValue.textContent = this.value;
    });
    
    // 设置偏远地区上浮的初始值为0%
    specialBonus.value = 0;
    specialBonusValue.textContent = 0;
    
    specialBonus.addEventListener('input', function() {
        specialBonusValue.textContent = this.value;
    });
    
    // 创建历史记录区域
    const historySection = document.createElement('div');
    historySection.className = 'history-section mt-4';
    historySection.innerHTML = `
        <div class="d-flex justify-content-between align-items-center mb-3">
            <h3 class="mb-0">历史记录</h3>
            <div class="d-flex align-items-center gap-3">
                <div class="engineer-input-group">
                    
                    <input type="text" class="form-control modern-input engineer-name-input" id="engineerName" placeholder="请输入工程师姓名" style="width: 220px;">
                </div>
                <div class="btn-group">
                            <button id="screenshotHistory" class="btn btn-info">
            <i class="bi bi-camera me-1"></i>绩效截图
        </button>
                        <button id="exportHistory" class="btn btn-success">
                            <i class="bi bi-download me-1"></i>导出记录
                        </button>
                        <button id="importHistory" class="btn btn-secondary">
                            <i class="bi bi-upload me-1"></i>导入记录
                        </button>
                        <input id="importHistoryFile" type="file" accept=".csv" style="display: none;" />
                        <button id="clearHistory" class="btn btn-warning">
                            <i class="bi bi-trash me-1"></i>清除记录
                        </button>
                    </div>
                </div>
            </div>
            <div id="historyRecords"></div>
        `;
    const targetContainer = document.querySelector('.main-container .container') || document.querySelector('.container');
    const welcomeCard = targetContainer ? targetContainer.querySelector('.welcome-card') : null;
    if (welcomeCard && welcomeCard.parentNode) {
        welcomeCard.parentNode.insertBefore(historySection, welcomeCard.nextSibling);
    } else if (targetContainer) {
        targetContainer.appendChild(historySection);
    }
    
    const historyRecords = document.getElementById('historyRecords');
    const clearHistoryBtn = document.getElementById('clearHistory');
    const exportHistoryBtn = document.getElementById('exportHistory');
    const screenshotHistoryBtn = document.getElementById('screenshotHistory');
    const engineerNameInput = document.getElementById('engineerName');
    const importHistoryBtn = document.getElementById('importHistory');
    const importHistoryFile = document.getElementById('importHistoryFile');
    
    // 从本地存储加载历史记录
    let calculationHistory = JSON.parse(localStorage.getItem('performanceHistory')) || [];
    
    // 显示历史记录
    function displayHistory() {
        historyRecords.innerHTML = '';
        if (calculationHistory.length === 0) {
            historyRecords.innerHTML = '<div class="alert alert-info">暂无历史记录</div>';
            return;
        }
        
        // 创建一个响应式表格容器
        const tableContainer = document.createElement('div');
        tableContainer.className = 'table-responsive';
        
        const table = document.createElement('table');
        table.className = 'table table-striped table-hover table-sm';
        table.innerHTML = `
            <thead class="table-light">
                <tr>
                    <th>时间</th>
                    <th>仪器编号</th>
                    <th>工单类型</th>
                    <th>绩效值</th>
                    <th>绩效比例</th>
                    <th>偏远上浮</th>
                    <th>天数因子</th>
                    <th>操作</th>
                </tr>
            </thead>
            <tbody></tbody>
        `;
        
        const tbody = table.querySelector('tbody');
        
        calculationHistory.forEach((record, index) => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${record.timestamp}</td>
                <td>${record.instrumentId}</td>
                <td>${record.workOrderType}</td>
                <td>
                    <div class="performance-editable" data-index="${index}">
                        <span class="performance-display">${record.performance.toFixed(2)} 元</span>
                        <div class="performance-edit d-none">
                            <input type="number" class="form-control form-control-sm performance-input" value="${record.performance}" step="0.01" min="0">
                            <div class="edit-actions mt-1">
                                <button class="btn btn-sm btn-success btn-save-performance" data-index="${index}">
                                    <i class="bi bi-check"></i>
                                </button>
                                <button class="btn btn-sm btn-secondary btn-cancel-performance" data-index="${index}">
                                    <i class="bi bi-x"></i>
                                </button>
                            </div>
                        </div>
                    </div>
                </td>
                <td>${record.performanceRatio}%</td>
                <td>${record.specialBonus}%</td>
                <td>${record.daysFactor}</td>
                <td>
                    <div class="btn-group btn-group-sm">
                        <button class="btn btn-sm btn-outline-warning edit-performance" data-index="${index}">
                            <i class="bi bi-pencil"></i>
                        </button>
                        <button class="btn btn-sm btn-outline-info view-details" data-index="${index}">
                            <i class="bi bi-info-circle"></i>
                        </button>
                        <button class="btn btn-sm btn-outline-danger delete-record" data-index="${index}">
                            <i class="bi bi-trash"></i>
                        </button>
                    </div>
                </td>
            `;
            tbody.appendChild(row);
        });
        
        tableContainer.appendChild(table);
        historyRecords.appendChild(tableContainer);
        
        // 添加编辑绩效值的事件监听
        document.querySelectorAll('.edit-performance').forEach(button => {
            button.addEventListener('click', function() {
                const index = parseInt(this.getAttribute('data-index'));
                const performanceEditable = document.querySelector(`.performance-editable[data-index="${index}"]`);
                const display = performanceEditable.querySelector('.performance-display');
                const edit = performanceEditable.querySelector('.performance-edit');
                
                display.classList.add('d-none');
                edit.classList.remove('d-none');
            });
        });
        
        // 添加保存绩效值的事件监听
        document.querySelectorAll('.btn-save-performance').forEach(button => {
            button.addEventListener('click', function() {
                const index = parseInt(this.getAttribute('data-index'));
                const performanceEditable = document.querySelector(`.performance-editable[data-index="${index}"]`);
                const display = performanceEditable.querySelector('.performance-display');
                const edit = performanceEditable.querySelector('.performance-edit');
                const input = edit.querySelector('.performance-input');
                
                const newValue = parseFloat(input.value);
                if (isNaN(newValue) || newValue < 0) {
                    alert('请输入有效的绩效值！');
                    return;
                }
                
                // 记录手动修改信息
                const previousValue = Number(calculationHistory[index].performance) || 0;
                if (calculationHistory[index].originalPerformance === undefined) {
                    calculationHistory[index].originalPerformance = previousValue;
                }
                calculationHistory[index].manualEdit = true;
                calculationHistory[index].manualEditInfo = {
                    previous: previousValue,
                    current: newValue,
                    editedAt: new Date().toLocaleString()
                };
                
                // 更新数据
                calculationHistory[index].performance = newValue;
                localStorage.setItem('performanceHistory', JSON.stringify(calculationHistory));
                
                // 更新显示
                display.textContent = `${newValue.toFixed(2)} 元`;
                display.classList.remove('d-none');
                edit.classList.add('d-none');
                
                // 更新统计摘要
                updateStatsSummary();
                
                // 显示成功提示
                showToast('绩效值已更新！', 'success');
            });
        });
        
        // 添加取消编辑的事件监听
        document.querySelectorAll('.btn-cancel-performance').forEach(button => {
            button.addEventListener('click', function() {
                const index = parseInt(this.getAttribute('data-index'));
                const performanceEditable = document.querySelector(`.performance-editable[data-index="${index}"]`);
                const display = performanceEditable.querySelector('.performance-display');
                const edit = performanceEditable.querySelector('.performance-edit');
                const input = edit.querySelector('.performance-input');
                
                // 恢复原值
                input.value = calculationHistory[index].performance;
                display.classList.remove('d-none');
                edit.classList.add('d-none');
            });
        });
        
        // 添加删除单条记录的事件监听
        document.querySelectorAll('.delete-record').forEach(button => {
            button.addEventListener('click', function() {
                const index = parseInt(this.getAttribute('data-index'));
                calculationHistory.splice(index, 1);
                localStorage.setItem('performanceHistory', JSON.stringify(calculationHistory));
                displayHistory();
                updateStatsSummary();
            });
        });
        
        // 添加查看详情的事件监听
        document.querySelectorAll('.view-details').forEach(button => {
            button.addEventListener('click', function() {
                const index = parseInt(this.getAttribute('data-index'));
                showRecordDetails(calculationHistory[index]);
            });
        });
    }
    
    // 添加截图功能的事件监听
    if (screenshotHistoryBtn) {
        screenshotHistoryBtn.addEventListener('click', function() {
            takeScreenshot();
        });
    }
    
    // 添加工程师姓名输入框的事件监听
    if (engineerNameInput) {
        engineerNameInput.addEventListener('input', function() {
            // 保存工程师姓名到本地存储
            localStorage.setItem('engineerName', this.value);
        });
        
        // 从本地存储加载工程师姓名
        const savedName = localStorage.getItem('engineerName');
        if (savedName) {
            engineerNameInput.value = savedName;
        }
    }
    
    // 初始显示历史记录
    displayHistory();
    
    // 导出历史记录
    exportHistoryBtn.addEventListener('click', function() {
        if (calculationHistory.length === 0) {
            alert('没有历史记录可导出');
            return;
        }
        
        // 创建CSV内容（新增外接设备、改装项目、维修信息、计算过程 JSON 列）
        let csvContent = "时间,仪器编号,工单类型,绩效值(元),绩效比例(%),偏远上浮(%),天数因子,外接设备,改装项目,维修信息,计算过程,手动修改,原始绩效(元),手动修改信息\n";
        
        calculationHistory.forEach(record => {
            const externalDevicesJson = JSON.stringify(record.externalDevices || []).replace(/"/g, '""');
            const modificationItemsJson = JSON.stringify(record.modificationItems || []).replace(/"/g, '""');
            const maintenanceJson = JSON.stringify(record.maintenance || null).replace(/"/g, '""');
            const calcProcessJson = JSON.stringify(record.calculationProcess || null).replace(/"/g, '""');
            const manualEditFlag = record.manualEdit ? '是' : '否';
            const originalPerf = record.originalPerformance !== undefined ? Number(record.originalPerformance).toFixed(2) : '';
            const manualEditInfoJson = JSON.stringify(record.manualEditInfo || null).replace(/"/g, '""');
            csvContent += `"${record.timestamp}","${record.instrumentId}","${record.workOrderType}","${Number(record.performance).toFixed(2)}","${record.performanceRatio}","${record.specialBonus}","${record.daysFactor}","${externalDevicesJson}","${modificationItemsJson}","${maintenanceJson}","${calcProcessJson}","${manualEditFlag}","${originalPerf}","${manualEditInfoJson}"\n`;
        });
        
        // 创建下载链接
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.setAttribute('href', url);
        const engineerNameForExport = (document.getElementById('engineerName')?.value || '未知工程师').trim() || '未知工程师';
        link.setAttribute('download', `${engineerNameForExport}_绩效计算记录_${new Date().toISOString().slice(0,10)}.csv`);
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    });

    // 导入历史记录（CSV）
    if (importHistoryBtn && importHistoryFile) {
        importHistoryBtn.addEventListener('click', function() {
            importHistoryFile.click();
        });

        importHistoryFile.addEventListener('change', function(e) {
            const file = e.target.files && e.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = function(ev) {
                try {
                    const text = ev.target.result;
                    const importedCount = importHistoryFromCSV(text);
                    showToast(`成功导入 ${importedCount} 条记录`, 'success');
                } catch (err) {
                    console.error(err);
                    showToast('导入失败：文件格式不正确或内容为空', 'danger');
                } finally {
                    importHistoryFile.value = '';
                }
            };
            reader.onerror = function() {
                showToast('读取文件失败，请重试', 'danger');
                importHistoryFile.value = '';
            };
            reader.readAsText(file, 'utf-8');
        });
    }
    
    // 清除所有历史记录
    clearHistoryBtn.addEventListener('click', function() {
        if (confirm('确定要清除所有历史记录吗？')) {
            calculationHistory = [];
            localStorage.setItem('performanceHistory', JSON.stringify(calculationHistory));
            displayHistory();
        }
    });
    
    // 创建模态框用于显示详细信息
    const detailsModal = document.createElement('div');
    detailsModal.className = 'modal fade';
    detailsModal.id = 'recordDetailsModal';
    detailsModal.tabIndex = '-1';
    detailsModal.setAttribute('aria-labelledby', 'recordDetailsModalLabel');
    detailsModal.setAttribute('aria-hidden', 'true');
    
    detailsModal.innerHTML = `
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="recordDetailsModalLabel">计算详情</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body" id="recordDetailsContent">
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">关闭</button>
                </div>
            </div>
        </div>
    `;
    
    document.body.appendChild(detailsModal);
    
    // 显示记录详情的函数
    function showRecordDetails(record) {
        const detailsContent = document.getElementById('recordDetailsContent');
        
        let content = `
            <div class="card mb-3">
                <div class="card-header bg-primary text-white">
                    <h5 class="mb-0">基本信息</h5>
                </div>
                <div class="card-body">
                    <div class="row">
                        <div class="col-md-6">
                            <p><strong>计算时间：</strong> ${record.timestamp}</p>
                            <p><strong>仪器编号：</strong> ${record.instrumentId}</p>
                            <p><strong>工单类型：</strong> ${record.workOrderType}</p>
                        </div>
                        <div class="col-md-6">
                            <p><strong>绩效值：</strong> ${Number(record.performance).toFixed(2)} 元</p>
                            <p><strong>绩效比例：</strong> ${record.performanceRatio}%</p>
                            <p><strong>偏远地区上浮：</strong> ${record.specialBonus}%</p>
                            <p><strong>天数因子：</strong> ${record.daysFactor}</p>
                            ${record.workOrderType === '维修' && record.maintenance ? `<p><strong>维修类型：</strong> ${record.maintenance.label}</p>` : ''}
                            ${record.manualEdit ? `<p class="text-warning mb-0"><strong>已手动修改：</strong> 原始 ${Number(record.originalPerformance ?? record.performance).toFixed(2)} 元 → 当前 ${Number(record.performance).toFixed(2)} 元（${record.manualEditInfo?.editedAt || ''}）</p>` : ''}
                        </div>
                    </div>
                </div>
            </div>
        `;
        
        // 如果有外接设备信息
        if (record.workOrderType === '安装' && record.externalDevices && record.externalDevices.length > 0) {
            content += `
                <div class="card mb-3">
                    <div class="card-header bg-info text-white">
                        <h5 class="mb-0">外接设备信息</h5>
                    </div>
                    <div class="card-body">
                        <table class="table table-sm">
                            <thead>
                                <tr>
                                    <th>设备类型</th>
                                    <th>数量</th>
                                </tr>
                            </thead>
                            <tbody>
            `;
            
            record.externalDevices.forEach(device => {
                if (device.quantity > 0) {
                    let deviceName = '';
                    switch (device.type) {
                        case 'HS930': deviceName = 'HS930顶空进样器'; break;
                        case 'HS950': deviceName = 'HS950顶空进样器'; break;
                        case 'otherHS': deviceName = '其他顶空进样器'; break;
                        case 'PT800': deviceName = 'PT800吹扫捕集'; break;
                        case 'TDA': deviceName = 'TDA热解析仪'; break;
                        case 'otherTD': deviceName = '其他热解析仪'; break;
                        case 'airConcentrator': deviceName = '大气浓缩仪'; break;
                        case 'flashEvaporator': deviceName = '闪蒸仪'; break;
                        case 'massSpectrometer': deviceName = '质谱仪'; break;
                        default: deviceName = device.type;
                    }
                    
                    content += `
                        <tr>
                            <td>${deviceName}</td>
                            <td>${device.quantity}</td>
                        </tr>
                    `;
                }
            });
            
            content += `
                            </tbody>
                        </table>
                    </div>
                </div>
            `;
        }
        
        // 如果有改装项目信息
        if (record.workOrderType === '改装' && record.modificationItems && record.modificationItems.length > 0) {
            content += `
                <div class="card mb-3">
                    <div class="card-header bg-success text-white">
                        <h5 class="mb-0">改装项目信息</h5>
                    </div>
                    <div class="card-body">
                        <table class="table table-sm">
                            <thead>
                                <tr>
                                    <th>项目类型</th>
                                    <th>数量</th>
                                </tr>
                            </thead>
                            <tbody>
            `;
            
            record.modificationItems.forEach(item => {
                if (item.quantity > 0) {
                    let itemName = '';
                    switch (item.type) {
                        case 'manualSixValve': itemName = '手动六通阀'; break;
                        case 'manualTenValve': itemName = '手动十通阀'; break;
                        case 'pneumaticSixValve': itemName = '气动六通阀'; break;
                        case 'pneumaticTenValve': itemName = '气动十通阀'; break;
                        case 'fourWayValve': itemName = '四通阀'; break;
                        case 'temperatureControlBox': itemName = '温控箱'; break;
                        default: itemName = item.type;
                    }
                    
                    content += `
                        <tr>
                            <td>${itemName}</td>
                            <td>${item.quantity}</td>
                        </tr>
                    `;
                }
            });
            
            content += `
                            </tbody>
                        </table>
                    </div>
                </div>
            `;
        }
        
        // 计算过程展示
        if (record.calculationProcess) {
            const p = record.calculationProcess;
            const currency = (v) => Number(v || 0).toFixed(2);
            let processHtml = '';
            if (p.steps) {
                if (p.steps.installation) {
                    const inst = p.steps.installation;
                    processHtml += `
                        <div class="mb-2"><strong>安装项：</strong></div>
                        <ul class="mb-2">
                            ${inst.model && inst.model.amount ? `<li>型号贡献：+${currency(inst.model.amount)} 元${inst.model.label ? `（${inst.model.label}）` : ''}</li>` : ''}
                            ${Array.isArray(inst.externalDevices) && inst.externalDevices.length ? `<li>外接设备：<div class="mt-1">
                                <table class="table table-sm mb-0"><thead><tr><th>设备</th><th>数量</th><th>单价</th><th>小计</th></tr></thead><tbody>
                                    ${inst.externalDevices.map(d => `<tr><td>${d.label || d.type}</td><td>${d.quantity}</td><td>${currency(d.unitPrice)}</td><td>${currency(d.subtotal)}</td></tr>`).join('')}
                                </tbody></table></div></li>` : ''}
                            ${Array.isArray(inst.equipment) && inst.equipment.length ? `<li>设备清单：<div class="mt-1">
                                <table class="table table-sm mb-0"><thead><tr><th>项目</th><th>数量</th><th>单价</th><th>小计</th></tr></thead><tbody>
                                    ${inst.equipment.map(e => `<tr><td>${e.label}</td><td>${e.quantity}</td><td>${currency(e.unitPrice)}</td><td>${currency(e.subtotal)}</td></tr>`).join('')}
                                </tbody></table></div></li>` : ''}
                        </ul>
                    `;
                }
                if (p.steps.maintenance) {
                    const m = p.steps.maintenance;
                    processHtml += `<div class="mb-2"><strong>维修项：</strong></div><ul class="mb-2"><li>${m.label || ''}：${currency(m.base)} 元</li></ul>`;
                }
                if (p.steps.modification) {
                    const mod = p.steps.modification;
                    processHtml += `
                        <div class="mb-2"><strong>改装项：</strong></div>
                        <div class="mt-1">
                            <table class="table table-sm mb-2"><thead><tr><th>项目</th><th>数量</th><th>单价</th><th>小计</th></tr></thead><tbody>
                                ${Array.isArray(mod.items) ? mod.items.map(e => `<tr><td>${e.label}</td><td>${e.quantity}</td><td>${currency(e.unitPrice)}</td><td>${currency(e.subtotal)}</td></tr>`).join('') : ''}
                            </tbody></table>
                        </div>
                    `;
                }
                if (p.steps.specialInstallation) {
                    const s = p.steps.specialInstallation;
                    processHtml += `<div class="mb-2"><strong>专用安装：</strong></div><ul class="mb-2"><li>${s.label || ''}：${currency(s.base)} 元</li></ul>`;
                }
            }
            processHtml += `
                <div class="mb-2"><strong>加成过程：</strong></div>
                <ul class="mb-2">
                    <li>基础合计：${currency(p.preAdjustTotal)} 元</li>
                    <li>绩效比例：× ${p.adjustments?.ratio || 100}% → ${currency(p.afterRatio)}</li>
                    <li>偏远上浮：× (1 + ${(p.adjustments?.bonus || 0)}%) → ${currency(p.afterBonus)}</li>
                    <li>天数因子：+ ${currency(p.daysAdd)} (${p.adjustments?.days || 0} 天 × 140)</li>
                    <li><strong>计算结果：</strong> ${currency(p.final)} 元</li>
                    ${record.manualEdit ? `<li class="text-warning"><strong>手动修改为：</strong> ${currency(record.performance)} 元（原始 ${currency(record.originalPerformance ?? p.final)} 元）</li>` : ''}
                </ul>
            `;
            content += `
                <div class="card mb-3">
                    <div class="card-header bg-secondary text-white">
                        <h5 class="mb-0">计算过程</h5>
                    </div>
                    <div class="card-body">
                        ${processHtml}
                    </div>
                </div>
            `;
        }
        
        detailsContent.innerHTML = content;
        
        // 显示模态框
        const modal = new bootstrap.Modal(document.getElementById('recordDetailsModal'));
        modal.show();
    }

    // 显示或隐藏改装项目和外接设备输入区域
    workOrderType.addEventListener('change', function() {
        // 隐藏所有可选字段
        modificationFields.classList.add('hidden');
        externalDevicesFields.classList.add('hidden');
        specialInstallationFields.classList.add('hidden');
        maintenanceFields.classList.add('hidden');
        
        // 根据选择显示相应字段
        if (this.value === 'modification') {
            modificationFields.classList.remove('hidden');
        } else if (this.value === 'installation') {
            externalDevicesFields.classList.remove('hidden');
        } else if (this.value === 'specialInstallation') {
            specialInstallationFields.classList.remove('hidden');
        } else if (this.value === 'maintenance') {
            maintenanceFields.classList.remove('hidden');
        }
    });

    // 添加改装项目
    addModificationItem.addEventListener('click', function() {
        const itemDiv = document.createElement('div');
        itemDiv.className = 'modification-item';
        itemDiv.innerHTML = `
            <select class="form-select modification-type">
                <option value="manualSixValve">手动六通阀</option>
                <option value="manualTenValve">手动十通阀</option>
                <option value="pneumaticSixValve">气动六通阀</option>
                <option value="pneumaticTenValve">气动十通阀</option>
                <option value="fourWayValve">四通阀</option>
                <option value="temperatureControlBox">温控箱</option>
                <option value="conversionFurnace">转化炉</option>
                <option value="filledInlet">填充进样口</option>
                <option value="capillaryInlet">毛细进样口</option>
                <option value="fidDetector">FID检测器</option>
                <option value="tcdDetector">TCD检测器</option>
                <option value="ecdDetector">ECD检测器</option>
                <option value="fpdDetector">FPD检测器</option>
                <option value="npdDetector">NPD检测器</option>
                <option value="sideHangingBox">侧挂箱</option>
                <option value="solenoidValve">电磁阀</option>
            </select>
            <input type="number" class="form-control modification-quantity" placeholder="数量" min="1" value="1">
            <button class="btn btn-sm btn-outline-danger remove-modification-item">
                <i class="bi bi-trash"></i>
            </button>
        `;
        modificationItems.appendChild(itemDiv);

        // 删除改装项目
        itemDiv.querySelector('.remove-modification-item').addEventListener('click', function() {
            itemDiv.remove();
        });
        
        // 优化下拉选择器的用户体验
        const selectElement = itemDiv.querySelector('.modification-type');
        selectElement.addEventListener('focus', function() {
            this.style.zIndex = '1000';
        });
        
        selectElement.addEventListener('blur', function() {
            this.style.zIndex = 'auto';
        });
    });

    // 添加外接设备
    addExternalDevice.addEventListener('click', function() {
        const itemDiv = document.createElement('div');
        itemDiv.className = 'external-device-item';
        itemDiv.innerHTML = `
            <select class="form-select external-device-type">
                <option value="HS930">HS930顶空进样器</option>
                <option value="HS950">HS950顶空进样器</option>
                <option value="otherHS">其他顶空进样器</option>
                <option value="PT800">PT800吹扫捕集</option>
                <option value="TDA">TDA热解析仪</option>
                <option value="otherTD">其他热解析仪</option>
                <option value="airConcentrator">大气浓缩仪</option>
                <option value="flashEvaporator">闪蒸仪</option>
                <option value="massSpectrometer">质谱仪</option>
                <option value="other">其他</option>
            </select>
            <input type="number" class="form-control external-device-quantity" placeholder="数量" min="1" value="1">
            <button class="btn btn-sm btn-outline-danger remove-external-device">
                <i class="bi bi-trash"></i>
            </button>
        `;
        externalDeviceItems.appendChild(itemDiv);

        // 删除外接设备
        itemDiv.querySelector('.remove-external-device').addEventListener('click', function() {
            itemDiv.remove();
        });
        
        // 优化下拉选择器的用户体验
        const selectElement = itemDiv.querySelector('.external-device-type');
        selectElement.addEventListener('focus', function() {
            this.style.zIndex = '1000';
        });
        
        selectElement.addEventListener('blur', function() {
            this.style.zIndex = 'auto';
        });
    });

    // 解析用户上传的Excel文件
    fileInput.addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (!file) return;

        // 显示加载中状态
        uploadStatus.innerHTML = `
            <div class="spinner-border spinner-border-sm text-primary" role="status">
                <span class="visually-hidden">加载中...</span>
            </div>
            <span class="ms-2">正在解析文件...</span>
        `;

        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // 假设第一个工作表包含所有仪器配置数据
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            instrumentData = XLSX.utils.sheet_to_json(firstSheet);
            
            // 更新上传状态
            uploadStatus.innerHTML = `
                <div class="alert alert-success mb-0">
                    <i class="bi bi-check-circle-fill me-2"></i>成功加载 ${instrumentData.length} 条仪器配置！
                </div>
            `;
            
            // 3秒后自动关闭模态框
            setTimeout(() => {
                const dataSourceModal = bootstrap.Modal.getInstance(document.getElementById('dataSourceModal'));
                if (dataSourceModal) {
                    dataSourceModal.hide();
                }
            }, 3000);
        };
        
        reader.onerror = function() {
            uploadStatus.innerHTML = `
                <div class="alert alert-danger mb-0">
                    <i class="bi bi-exclamation-triangle-fill me-2"></i>文件解析失败，请检查文件格式！
                </div>
            `;
        };
        
        reader.readAsArrayBuffer(file);
    });

    // 计算绩效
    calculateButton.addEventListener('click', function() {
        const instrumentId = document.getElementById('instrumentId').value;
        const performanceRatioValue = parseInt(performanceRatio.value);
        const specialBonusValue = parseInt(specialBonus.value);
        const daysFactorValue = parseInt(daysFactor.value) || 0;
        const workOrderTypeValue = workOrderType.value;
        let wasInstrumentMissingButExternalDevicesPresent = false; // 当安装工单找不到仪器编号但添加了外接设备时标记

        if (!instrumentId) {
            resultDiv.textContent = '请输入仪器编号！';
            return;
        }

        // 检查是否有外接设备
        const hasExternalDevices = document.querySelectorAll('.external-device-item').length > 0;
        
        // 专用仪器安装和维修不需要Excel配置
        if (workOrderTypeValue !== 'specialInstallation' && workOrderTypeValue !== 'maintenance' && instrumentData.length === 0) {
            resultDiv.textContent = '请先上传Excel配置文件！';
            return;
        }

        let instrumentConfig = null;
        
        // 如果不是专用仪器安装或维修，需要查找仪器配置
        if (workOrderTypeValue !== 'specialInstallation' && workOrderTypeValue !== 'maintenance') {
            // 1. 尝试精确匹配编号字段
            const possibleIdFields = ['编号', 'ID', '仪器编号', 'instrumentId', 'id'];
            for (const field of possibleIdFields) {
                if (instrumentConfig) break;
                instrumentConfig = instrumentData.find(item => item[field] === instrumentId);
            }
            
            // 2. 对于安装工单，检查是否需要严格匹配
            if (!instrumentConfig && workOrderTypeValue === 'installation') {
                // 如果添加了外接设备，允许单独计算外接设备绩效
                if (hasExternalDevices) {
                    // 使用空的仪器配置，只计算外接设备绩效
                    instrumentConfig = {
                        '手动六通阀': 0,
                        '手动十通阀': 0,
                        '气动六通阀': 0,
                        '气动十通阀': 0,
                        '四通阀': 0,
                        'FL1092': 0,
                        'FL1090B': 0,
                        '填充进样口': 0,
                        '毛细进样口': 0,
                        'FID检测器': 0,
                        'TCD检测器': 0,
                        'ECD检测器': 0,
                        'FPD检测器': 0,
                        'NPD检测器': 0,
                        'HID检测器': 0,
                        'SCD检测器': 0,
                        '温控箱': 0,
                        '转化炉': 0,
                        '侧挂箱': 0,
                        '电磁阀': 0
                    };
                    wasInstrumentMissingButExternalDevicesPresent = true;
                    
                    // 使用 toast 提示，不再写入 result 容器，避免后续被清空
                    showToast('未找到仪器配置，已检测到外接设备：将仅计算型号贡献与外接设备绩效', 'info');
                } else {
                    // 没有外接设备，要求严格匹配
                    resultDiv.classList.remove('d-none');
                    resultDiv.className = 'alert alert-danger mt-4';
                    resultDiv.innerHTML = `
                        <div class="d-flex align-items-center">
                            <i class="bi bi-exclamation-triangle-fill me-2 text-danger"></i>
                            <div>
                                <h5 class="mb-1">错误：仪器编号不匹配</h5>
                                <p class="mb-0">您输入的仪器编号 <strong>"${instrumentId}"</strong> 在配置文件中未找到精确匹配。</p>
                                <p class="mb-0 small text-muted">请检查编号是否完整正确，或联系管理员确认仪器配置信息。</p>
                                <p class="mb-0 small text-muted">或者您可以添加外接设备来单独计算外接设备绩效。</p>
                            </div>
                        </div>
                    `;
                    return; // 直接返回，不继续计算
                }
            }
            
            // 3. 对于改装工单，如果没找到精确匹配，尝试部分匹配
            if (!instrumentConfig && workOrderTypeValue === 'modification') {
                for (const field of possibleIdFields) {
                    if (instrumentConfig) break;
                    instrumentConfig = instrumentData.find(item => 
                        item[field] && item[field].toString().includes(instrumentId)
                    );
                }
            }
            
            // 4. 如果还是没找到，对于改装工单，使用默认配置
            if (!instrumentConfig && workOrderTypeValue === 'modification') {
                instrumentConfig = {
                    '手动六通阀': 0,
                    '手动十通阀': 0,
                    '气动六通阀': 0,
                    '气动十通阀': 0,
                    '四通阀': 0,
                    'FL1092': 0,
                    'FL1090B': 0,
                    '填充进样口': 0,
                    '毛细进样口': 0,
                    'FID检测器': 0,
                    'TCD检测器': 0,
                    'ECD检测器': 0,
                    'FPD检测器': 0,
                    'NPD检测器': 0,
                    'HID检测器': 0,
                    'SCD检测器': 0,
                    '温控箱': 0,
                    '转化炉': 0,
                    '侧挂箱': 0,
                    '电磁阀': 0
                };
                resultDiv.classList.remove('d-none');
                resultDiv.className = 'alert alert-warning mt-4';
                resultDiv.innerHTML = `
                    <div class="d-flex align-items-center">
                        <i class="bi bi-exclamation-triangle-fill me-2 text-warning"></i>
                        <div>
                            <h5 class="mb-1">警告：使用默认配置</h5>
                            <p class="mb-0">未找到编号为 <strong>"${instrumentId}"</strong> 的仪器配置，将使用默认配置计算改装绩效。</p>
                        </div>
                    </div>
                `;
            }
        }

        let totalPerformance = 0;
        
        if (workOrderTypeValue === 'installation') {
            // 动态计算安装绩效
            totalPerformance = calculateInstallationPerformance(instrumentId, instrumentConfig);
        } else if (workOrderTypeValue === 'maintenance') {
            // 计算维修绩效
            totalPerformance = calculateMaintenancePerformance();
        } else if (workOrderTypeValue === 'modification') {
            // 计算改装维修绩效
            totalPerformance = calculateModificationPerformance();
        } else if (workOrderTypeValue === 'specialInstallation') {
            // 计算专用仪器安装绩效
            totalPerformance = calculateSpecialInstallationPerformance();
        }

        // 在应用比例/上浮/天数前的基础合计（用于计算过程展示）
        let preAdjustTotal = totalPerformance;

        // 应用绩效比例和加成
        totalPerformance = totalPerformance * (performanceRatioValue / 100);
        const afterRatio = totalPerformance;
        totalPerformance = totalPerformance * (1 + specialBonusValue / 100);
        const afterBonus = totalPerformance;
        
        // 添加天数因子的影响（每天增加140元）
        totalPerformance += 140 * daysFactorValue;
        const daysAdd = 140 * daysFactorValue;

        // 显示结果，包含仪器编号和天数因子信息
        resultDiv.classList.remove('d-none');
        resultDiv.className = wasInstrumentMissingButExternalDevicesPresent ? 'alert alert-warning mt-4' : 'alert alert-success mt-4';
        
        let resultText = `
            <div class="d-flex align-items-center gap-2 flex-wrap">
                <h4 class="mb-0">总绩效：<strong>${totalPerformance.toFixed(2)}</strong> 元</h4>
                <span class="badge bg-secondary">仪器编号: ${instrumentId}</span>
            </div>
        `;
        
        // 如果天数因子大于0，显示天数因子的贡献
        if (daysFactorValue > 0) {
            resultText += `
                <div class="mt-2 small">
                    <i class="bi bi-info-circle me-1"></i>其中包含天数因子(${daysFactorValue}天)贡献: 
                    <strong>${(140 * daysFactorValue).toFixed(2)}</strong> 元
                </div>
            `;
        }
        
        // 若仅按外接设备计算（未匹配到仪器配置），在结果中追加醒目提示
        if (wasInstrumentMissingButExternalDevicesPresent) {
            resultText += `
                <div class="mt-2 small text-warning">
                    <i class="bi bi-exclamation-triangle-fill me-1"></i>
                    该编号仪器在数据源中未找到，仅计算了型号贡献与外接设备绩效，请检查仪器内部配置后确认总绩效是否准确。
                </div>
            `;
        }

        resultDiv.innerHTML = resultText;
        
        // 生成计算过程分解并随记录保存
        function getExternalLabel(type) {
            switch (type) {
                case 'HS930': return 'HS930顶空进样器';
                case 'HS950': return 'HS950顶空进样器';
                case 'otherHS': return '其他顶空进样器';
                case 'PT800': return 'PT800吹扫捕集';
                case 'TDA': return 'TDA热解析仪';
                case 'otherTD': return '其他热解析仪';
                case 'airConcentrator': return '大气浓缩仪';
                case 'flashEvaporator': return '闪蒸仪';
                case 'massSpectrometer': return '质谱仪';
                case 'other': return '其他';
                default: return type;
            }
        }
        function getExternalUnitPrice(type) {
            switch (type) {
                case 'HS930': return 210;
                case 'HS950': return 250;
                case 'otherHS': return 210;
                case 'PT800': return 280;
                case 'TDA': return 280;
                case 'otherTD': return 280;
                case 'airConcentrator': return 280;
                case 'flashEvaporator': return 140;
                case 'massSpectrometer': return 500;
                case 'other': return 0;
                default: return 0;
            }
        }
        function getModificationLabel(type) {
            switch (type) {
                case 'manualSixValve': return '手动六通阀';
                case 'manualTenValve': return '手动十通阀';
                case 'pneumaticSixValve': return '气动六通阀';
                case 'pneumaticTenValve': return '气动十通阀';
                case 'fourWayValve': return '四通阀';
                case 'temperatureControlBox': return '温控箱';
                case 'conversionFurnace': return '转化炉';
                case 'filledInlet': return '填充进样口';
                case 'capillaryInlet': return '毛细进样口';
                case 'fidDetector': return 'FID检测器';
                case 'tcdDetector': return 'TCD检测器';
                case 'ecdDetector': return 'ECD检测器';
                case 'fpdDetector': return 'FPD检测器';
                case 'npdDetector': return 'NPD检测器';
                case 'sideHangingBox': return '侧挂箱';
                case 'solenoidValve': return '电磁阀';
                default: return type;
            }
        }
        function getModificationUnitPrice(type) {
            switch (type) {
                case 'manualSixValve': return 70;
                case 'manualTenValve': return 110;
                case 'pneumaticSixValve': return 140;
                case 'pneumaticTenValve': return 170;
                case 'fourWayValve': return 55;
                case 'temperatureControlBox': return 70;
                case 'conversionFurnace': return 85;
                case 'filledInlet': return 55;
                case 'capillaryInlet': return 110;
                case 'fidDetector': return 170;
                case 'tcdDetector': return 170;
                case 'ecdDetector': return 170;
                case 'fpdDetector': return 170;
                case 'npdDetector': return 170;
                case 'sideHangingBox': return 70;
                case 'solenoidValve': return 30;
                default: return 0;
            }
        }

        const calculationProcess = {
            preAdjustTotal: preAdjustTotal,
            afterRatio: afterRatio,
            afterBonus: afterBonus,
            daysAdd: daysAdd,
            final: totalPerformance,
            adjustments: { ratio: performanceRatioValue, bonus: specialBonusValue, days: daysFactorValue },
            steps: {}
        };

        if (workOrderTypeValue === 'installation') {
            const externalDevices = Array.from(document.querySelectorAll('.external-device-item')).map(device => {
                const type = device.querySelector('.external-device-type').value;
                const quantity = parseInt(device.querySelector('.external-device-quantity').value) || 0;
                const unit = getExternalUnitPrice(type);
                return { type, label: getExternalLabel(type), quantity, unitPrice: unit, subtotal: unit * quantity };
            }).filter(d => d.quantity > 0);

            const equipmentPriceMap = [
                { key: '手动六通阀', unit: 30 },
                { key: '手动十通阀', unit: 30 },
                { key: '气动六通阀', unit: 40 },
                { key: '气动十通阀', unit: 40 },
                { key: '四通阀', unit: 30 },
                { key: 'FL1092', unit: 70 },
                { key: '填充进样口', unit: 30 },
                { key: '毛细进样口', unit: 50 },
                { key: 'FID检测器', unit: 70 },
                { key: 'TCD检测器', unit: 80 },
                { key: 'ECD检测器', unit: 80 },
                { key: 'FPD检测器', unit: 80 },
                { key: 'NPD检测器', unit: 80 },
                { key: 'HID检测器', unit: 280 },
                { key: 'SCD检测器', unit: 170 },
                { key: '温控箱', unit: 15 },
                { key: '转化炉', unit: 30 }
            ];
            const equipment = [];
            if (instrumentConfig) {
                // 计算备注驱动的 FL1092 数量规则
                const remarkFields = ['备注', '备注信息', '备注说明', '说明'];
                let remarksRaw = '';
                for (const f of remarkFields) {
                    if (instrumentConfig[f]) { remarksRaw = String(instrumentConfig[f]); break; }
                }
                const normalizedRemarks = remarksRaw.replace(/（/g, '(').replace(/）/g, ')');
                const fl1092Base = Number(instrumentConfig['FL1092'] || 0);
                // 互斥匹配：先判断“进口日本”规则（基础+1），否则再判断双塔=3，其次150位=2，否则用基础值
                let fl1092Effective = fl1092Base;
                if (normalizedRemarks.includes('进口日本150位自动进样器') ||
                    normalizedRemarks.includes('进口日本双塔自动进样器(2T+1C)')) {
                    fl1092Effective = fl1092Base + 1;
                } else if (normalizedRemarks.includes('双塔自动进样器(2T+1C)')) {
                    fl1092Effective = 3;
                } else if (normalizedRemarks.includes('150位自动进样器')) {
                    fl1092Effective = 2;
                }

                equipmentPriceMap.forEach(e => {
                    let qty = Number(instrumentConfig[e.key] || 0);
                    if (e.key === 'FL1092') {
                        qty = fl1092Effective;
                    }
                    if (qty > 0) {
                        equipment.push({ label: e.key, quantity: qty, unitPrice: e.unit, subtotal: e.unit * qty });
                    }
                });
            }
            let modelAmount = 0;
            let modelLabel = '';
            if (instrumentId.includes('F60') || instrumentId.includes('97902')) {
                modelAmount = 80; modelLabel = instrumentId.includes('F60') ? 'F60' : '97902';
            } else if (instrumentId.includes('F70') || instrumentId.includes('9790P')) {
                modelAmount = 100; modelLabel = instrumentId.includes('F70') ? 'F70' : '9790P';
            } else if (instrumentId.includes('F80') || instrumentId.includes('9720P')) {
                modelAmount = 120; modelLabel = instrumentId.includes('F80') ? 'F80' : '9720P';
            }

            calculationProcess.steps.installation = {
                model: { amount: modelAmount, label: modelLabel },
                externalDevices: externalDevices,
                equipment: equipment
            };
        } else if (workOrderTypeValue === 'maintenance') {
            const labelMap = {
                inWarranty: '保内维修',
                outWarranty: '保外维修',
                dismantle: '维修拆机',
                '3qCertification': '3Q认证',
                training: '培训服务',
                valueAdded: '增值服务',
                relocation: '仪器移机',
                techSupport: '技术支持',
                visit: '上门回访',
                maintenance: '维保服务'
            };
            const base = calculateMaintenancePerformance();
            calculationProcess.steps.maintenance = { label: labelMap[maintenanceType.value] || maintenanceType.value, base: base };
        } else if (workOrderTypeValue === 'modification') {
            const items = Array.from(document.querySelectorAll('.modification-item')).map(item => {
                const type = item.querySelector('.modification-type').value;
                const quantity = parseInt(item.querySelector('.modification-quantity').value) || 0;
                const unit = getModificationUnitPrice(type);
                return { type, label: getModificationLabel(type), quantity, unitPrice: unit, subtotal: unit * quantity };
            }).filter(e => e.quantity > 0);
            calculationProcess.steps.modification = { items };
        } else if (workOrderTypeValue === 'specialInstallation') {
            const labelMap = { NMHC: 'NMHC分析仪', oxygenCompound: '含氧化合物分析仪', powerChromatography: '电力色谱仪', portableChromatography: '便携色谱仪' };
            const base = calculateSpecialInstallationPerformance();
            calculationProcess.steps.specialInstallation = { label: labelMap[specialInstrumentType.value] || specialInstrumentType.value, base: base };
        }

        // 添加到历史记录
        const now = new Date();
        const timestamp = now.toLocaleString();
        
        // 获取工单类型的显示名称
        let workOrderTypeName = '';
        switch (workOrderTypeValue) {
            case 'installation':
                workOrderTypeName = '安装';
                break;
            case 'maintenance':
                workOrderTypeName = '维修';
                break;
            case 'modification':
                workOrderTypeName = '改装';
                break;
            case 'specialInstallation':
                workOrderTypeName = '专用仪器安装';
                break;
        }
        
        // 创建历史记录
        const historyRecord = {
            timestamp: timestamp,
            instrumentId: instrumentId,
            workOrderType: workOrderTypeName,
            performance: totalPerformance,
            daysFactor: daysFactorValue,
            performanceRatio: performanceRatio.value,
            specialBonus: specialBonus.value,
            externalDevices: workOrderTypeValue === 'installation' ? Array.from(document.querySelectorAll('.external-device-item')).map(device => ({
                type: device.querySelector('.external-device-type').value,
                quantity: parseInt(device.querySelector('.external-device-quantity').value) || 0
            })) : [],
            modificationItems: workOrderTypeValue === 'modification' ? Array.from(document.querySelectorAll('.modification-item')).map(item => ({
                type: item.querySelector('.modification-type').value,
                quantity: parseInt(item.querySelector('.modification-quantity').value) || 0
            })) : [],
            maintenance: workOrderTypeValue === 'maintenance' ? {
                type: maintenanceType.value,
                label: (function() {
                    switch (maintenanceType.value) {
                        case 'inWarranty': return '保内维修';
                        case 'outWarranty': return '保外维修';
                        case 'dismantle': return '维修拆机';
                        case '3qCertification': return '3Q认证';
                        case 'training': return '培训服务';
                        case 'valueAdded': return '增值服务';
                        case 'relocation': return '仪器移机';
                        case 'techSupport': return '技术支持';
                        case 'visit': return '上门回访';
                        case 'maintenance': return '维保服务';
                        default: return maintenanceType.value;
                    }
                })()
            } : null,
            calculationProcess: calculationProcess,
            originalPerformance: totalPerformance,
            manualEdit: false
        };
        
        // 添加到历史记录数组
        calculationHistory.unshift(historyRecord); // 添加到数组开头
        
        // 限制历史记录数量，最多保存50条
        if (calculationHistory.length > 50) {
            calculationHistory = calculationHistory.slice(0, 50);
        }
        
        // 保存到本地存储
        localStorage.setItem('performanceHistory', JSON.stringify(calculationHistory));
        
        // 更新累计计算统计
        updateStatsSummary();
        
        // 更新显示
        displayHistory();
    });

    // 专用仪器安装绩效计算函数
    function calculateSpecialInstallationPerformance() {
        const type = specialInstrumentType.value;
        let performance = 0;
        
        switch (type) {
            case 'NMHC':
                performance = 550;
                break;
            case 'oxygenCompound':
                performance = 550;
                break;
            case 'powerChromatography':
                performance = 550;
                break;
            case 'portableChromatography':
                performance = 500;
                break;
        }
        
        return performance;
    }

    // 维修绩效计算函数
    function calculateMaintenancePerformance() {
        const type = maintenanceType.value;
        let performance = 0;
        
        switch (type) {
            case 'inWarranty':
                performance = 140;
                break;
            case 'outWarranty':
                performance = 170;
                break;
            case 'dismantle':
                performance = 80;
                break;
            case '3qCertification':
                performance = 500;
                break;
            case 'training':
                performance = 140;
                break;
            case 'valueAdded':
                performance = 140;
                break;
            case 'relocation':
                performance = 140;
                break;
            case 'techSupport':
                performance = 140;
                break;
            case 'visit':
                performance = 140;
                break;
            case 'maintenance':
                performance = 140;
                break;
        }
        
        return performance;
    }

    // 安装绩效计算函数（动态关联Excel配置）
    function calculateInstallationPerformance(instrumentId, config) {
        let performance = 0;

        // 仪器型号绩效（根据编号或特定型号码匹配）
        if (instrumentId.includes('F60') || instrumentId.includes('97902')) performance += 80;
        else if (instrumentId.includes('F70') || instrumentId.includes('9790P')) performance += 100;
        else if (instrumentId.includes('F80') || instrumentId.includes('9720P')) performance += 120;
        
        // 外接设备绩效
        const externalDevices = document.querySelectorAll('.external-device-item');
        externalDevices.forEach(device => {
            const type = device.querySelector('.external-device-type').value;
            const quantity = parseInt(device.querySelector('.external-device-quantity').value) || 0;
            
            switch (type) {
                case 'HS930':
                    performance += quantity * 210;
                    break;
                case 'HS950':
                    performance += quantity * 250;
                    break;
                case 'otherHS':
                    performance += quantity * 210;
                    break;
                case 'PT800':
                    performance += quantity * 280;
                    break;
                case 'TDA':
                    performance += quantity * 280;
                    break;
                case 'otherTD':
                    performance += quantity * 280;
                    break;
                case 'airConcentrator':
                    performance += quantity * 280;
                    break;
                case 'flashEvaporator':
                    performance += quantity * 140;
                    break;
                case 'massSpectrometer':
                    performance += quantity * 500;
                    break;
                case 'other':
                    performance += quantity * 0;
                    break;
            }
        });

        // 设备绩效（从Excel配置中动态读取数量）
        performance += (config['手动六通阀'] || 0) * 30;
        performance += (config['手动十通阀'] || 0) * 30;
        performance += (config['气动六通阀'] || 0) * 40;
        performance += (config['气动十通阀'] || 0) * 40;
        performance += (config['四通阀'] || 0) * 30;
        // 读取备注并按规则调整 FL1092 的数量；去掉 FL1090B 的计算
        (function() {
            const remarkFields = ['备注', '备注信息', '备注说明', '说明'];
            let remarksRaw = '';
            for (const f of remarkFields) {
                if (config && config[f]) { remarksRaw = String(config[f]); break; }
            }
            const normalizedRemarks = remarksRaw.replace(/（/g, '(').replace(/）/g, ')');
            const fl1092Base = Number(config['FL1092'] || 0);
            let fl1092Effective = fl1092Base;
            if (normalizedRemarks.includes('进口日本150位自动进样器') ||
                normalizedRemarks.includes('进口日本双塔自动进样器(2T+1C)')) {
                fl1092Effective = fl1092Base + 1;
            } else if (normalizedRemarks.includes('双塔自动进样器(2T+1C)')) {
                fl1092Effective = 3;
            } else if (normalizedRemarks.includes('150位自动进样器')) {
                fl1092Effective = 2;
            }
            performance += fl1092Effective * 70;
        })();
        // 移除：performance += (config['FL1090B'] || 0) * 140;
        performance += (config['填充进样口'] || 0) * 30;
        performance += (config['毛细进样口'] || 0) * 50;
        performance += (config['FID检测器'] || 0) * 70;
        performance += (config['TCD检测器'] || 0) * 80;
        performance += (config['ECD检测器'] || 0) * 80;
        performance += (config['FPD检测器'] || 0) * 80;
        performance += (config['NPD检测器'] || 0) * 80;
        performance += (config['HID检测器'] || 0) * 280;
        performance += (config['SCD检测器'] || 0) * 170;
        performance += (config['温控箱'] || 0) * 15;
        performance += (config['转化炉'] || 0) * 30;

        return performance;
    }

    // 改装维修绩效计算函数
    function calculateModificationPerformance() {
        const items = document.querySelectorAll('.modification-item');
        let performance = 0;

        items.forEach(item => {
            const type = item.querySelector('.modification-type').value;
            const quantity = parseInt(item.querySelector('.modification-quantity').value) || 0;

            switch (type) {
                case 'manualSixValve':
                    performance += quantity * 70;
                    break;
                case 'manualTenValve':
                    performance += quantity * 110;
                    break;
                case 'pneumaticSixValve':
                    performance += quantity * 140;
                    break;
                case 'pneumaticTenValve':
                    performance += quantity * 170;
                    break;
                case 'fourWayValve':
                    performance += quantity * 55;
                    break;
                case 'temperatureControlBox':
                    performance += quantity * 70;
                    break;
                case 'conversionFurnace':
                    performance += quantity * 85;
                    break;
                case 'filledInlet':
                    performance += quantity * 55;
                    break;
                case 'capillaryInlet':
                    performance += quantity * 110;
                    break;
                case 'fidDetector':
                    performance += quantity * 170;
                    break;
                case 'tcdDetector':
                    performance += quantity * 170;
                    break;
                case 'ecdDetector':
                    performance += quantity * 170;
                    break;
                case 'fpdDetector':
                    performance += quantity * 170;
                    break;
                case 'npdDetector':
                    performance += quantity * 170;
                    break;
                case 'sideHangingBox':
                    performance += quantity * 70;
                    break;
                case 'solenoidValve':
                    performance += quantity * 30;
                    break;
            }
        });

        return performance;
    }
    
    // 更新统计摘要（累计次数与累计金额）
    function updateStatsSummary() {
        const totalCalculationsElement = document.getElementById('totalCalculations');
        const totalAmountElement = document.getElementById('totalAmount');
        
        if (totalCalculationsElement) {
            const total = calculationHistory.length;
            totalCalculationsElement.textContent = total;
            totalCalculationsElement.style.transform = 'scale(1.1)';
            setTimeout(() => { totalCalculationsElement.style.transform = 'scale(1)'; }, 200);
        }
        
        if (totalAmountElement) {
            const sum = calculationHistory.reduce((acc, r) => acc + (Number(r.performance) || 0), 0);
            totalAmountElement.textContent = sum.toFixed(2);
            totalAmountElement.style.transform = 'scale(1.1)';
            setTimeout(() => { totalAmountElement.style.transform = 'scale(1)'; }, 200);
        }
    }
    
    // 页面加载时初始化统计
    updateStatsSummary();
    
    // 截图功能
    function takeScreenshot() {
        // 检查是否加载了html2canvas
        if (typeof html2canvas === 'undefined') {
            showToast('截图功能加载失败，请刷新页面重试', 'danger');
            return;
        }

        // 显示加载状态
        screenshotHistoryBtn.innerHTML = '<i class="bi bi-hourglass-split me-1"></i>生成中...';
        screenshotHistoryBtn.disabled = true;

        // 仅截取欢迎卡片 + 历史记录区域（不包含表单）
        const mainContainer = document.querySelector('.main-container .container');
        const welcomeCard = mainContainer ? mainContainer.querySelector('.welcome-card') : null;
        const historySectionEl = document.querySelector('.history-section');

        if (!welcomeCard || !historySectionEl) {
            screenshotHistoryBtn.innerHTML = '<i class="bi bi-camera me-1"></i>绩效截图';
            screenshotHistoryBtn.disabled = false;
            showToast('未找到可截图的区域', 'danger');
            return;
        }

        // 创建离屏容器，克隆必要区域，避免包含表单
        const offscreen = document.createElement('div');
        offscreen.className = 'container';
        offscreen.style.position = 'fixed';
        offscreen.style.left = '-99999px';
        offscreen.style.top = '0';
        offscreen.style.backgroundColor = '#ffffff';
        offscreen.style.padding = '0 0 24px';

        const welcomeClone = welcomeCard.cloneNode(true);
        const historyClone = historySectionEl.cloneNode(true);

        // 同步工程师姓名输入框的值到克隆节点，避免截图中为空
        const originalEngineerInput = historySectionEl.querySelector('#engineerName');
        const clonedEngineerInput = historyClone.querySelector('#engineerName');
        if (originalEngineerInput && clonedEngineerInput) {
            clonedEngineerInput.value = originalEngineerInput.value;
        }

        offscreen.appendChild(welcomeClone);
        offscreen.appendChild(historyClone);
        document.body.appendChild(offscreen);

        // 创建截图（仅包含欢迎卡片与历史记录）
        html2canvas(offscreen, {
            backgroundColor: '#ffffff',
            scale: 2, // 提高截图质量
            useCORS: true,
            allowTaint: true,
            logging: false
        }).then(canvas => {
            // 创建下载链接
            const link = document.createElement('a');
            const engineerName = document.getElementById('engineerName')?.value || '未知工程师';
            const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
            link.download = `${engineerName}_绩效报告_${timestamp}.png`;
            link.href = canvas.toDataURL();

            // 触发下载
            link.click();

            // 显示成功提示
            showToast('绩效报告截图已生成并下载！', 'success');
        }).catch(error => {
            console.error('截图生成失败:', error);
            // 显示错误提示
            showToast('截图生成失败，请重试', 'danger');
        }).finally(() => {
            // 清理临时容器并恢复按钮状态
            if (offscreen && offscreen.parentNode) {
                offscreen.parentNode.removeChild(offscreen);
            }
            screenshotHistoryBtn.innerHTML = '<i class="bi bi-camera me-1"></i>绩效截图';
            screenshotHistoryBtn.disabled = false;
        });
    }
    
    // 显示提示信息
    function showToast(message, type = 'info') {
        // 创建提示元素
        const toast = document.createElement('div');
        toast.className = `toast-notification toast-${type}`;
        toast.innerHTML = `
            <div class="toast-content">
                <i class="bi bi-${type === 'success' ? 'check-circle' : 'info-circle'} me-2"></i>
                <span>${message}</span>
            </div>
        `;
        
        // 添加到页面
        document.body.appendChild(toast);
        
        // 显示动画
        setTimeout(() => toast.classList.add('show'), 100);
        
        // 自动隐藏
        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => document.body.removeChild(toast), 300);
        }, 3000);
    }

    // 解析CSV为二维数组并导入到历史记录
    function importHistoryFromCSV(text) {
        const rows = parseCSV(text);
        if (!rows || rows.length < 2) throw new Error('empty');
        const headers = rows[0].map(h => h.trim());
        const idx = {
            time: headers.indexOf('时间'),
            instrumentId: headers.indexOf('仪器编号'),
            workOrderType: headers.indexOf('工单类型'),
            performance: headers.indexOf('绩效值(元)'),
            performanceRatio: headers.indexOf('绩效比例(%)'),
            specialBonus: headers.indexOf('偏远上浮(%)') >= 0 ? headers.indexOf('偏远上浮(%)') : headers.indexOf('偏远地区上浮(%)'),
            daysFactor: headers.indexOf('天数因子'),
            externalDevices: headers.indexOf('外接设备'),
            modificationItems: headers.indexOf('改装项目'),
            maintenance: headers.indexOf('维修信息'),
            calculationProcess: headers.indexOf('计算过程'),
            manualEdit: headers.indexOf('手动修改'),
            originalPerformance: headers.indexOf('原始绩效(元)'),
            manualEditInfo: headers.indexOf('手动修改信息')
        };
        // 旧版CSV没有新增两列时，允许通过
        const required = ['time','instrumentId','workOrderType','performance','performanceRatio','specialBonus','daysFactor'];
        if (required.some(k => idx[k] === -1)) throw new Error('headers');

        const existing = new Set(
            calculationHistory.map(r => `${r.timestamp}|${r.instrumentId}|${r.workOrderType}|${Number(r.performance).toFixed(2)}`)
        );

        let added = 0;
        for (let i = 1; i < rows.length; i++) {
            const cols = rows[i];
            if (!cols || cols.length === 0 || cols.every(c => !c || String(c).trim() === '')) continue;
            const timestamp = cols[idx.time] || '';
            const instId = cols[idx.instrumentId] || '';
            const orderType = cols[idx.workOrderType] || '';
            const perf = parseFloat(cols[idx.performance] || '0');
            const ratio = cols[idx.performanceRatio] || '100';
            const bonus = cols[idx.specialBonus] || '0';
            const days = parseInt(cols[idx.daysFactor] || '0');

            // 新增：解析外接设备与改装项目（若列不存在或为空则置空数组）
            let externalDevicesArr = [];
            if (idx.externalDevices >= 0) {
                const raw = cols[idx.externalDevices] || '';
                if (raw && raw.trim()) {
                    try { externalDevicesArr = JSON.parse(raw); } catch(e) { externalDevicesArr = []; }
                }
            }
            let modificationItemsArr = [];
            if (idx.modificationItems >= 0) {
                const raw = cols[idx.modificationItems] || '';
                if (raw && raw.trim()) {
                    try { modificationItemsArr = JSON.parse(raw); } catch(e) { modificationItemsArr = []; }
                }
            }
            let maintenanceObj = null;
            if (idx.maintenance >= 0) {
                const raw = cols[idx.maintenance] || '';
                if (raw && raw.trim()) {
                    try { maintenanceObj = JSON.parse(raw); } catch(e) { maintenanceObj = null; }
                }
            }
            let calcProcessObj = null;
            if (idx.calculationProcess >= 0) {
                const raw = cols[idx.calculationProcess] || '';
                if (raw && raw.trim()) {
                    try { calcProcessObj = JSON.parse(raw); } catch(e) { calcProcessObj = null; }
                }
            }
            // 解析手动修改信息
            let manualEditFlag = false;
            if (idx.manualEdit >= 0) {
                const v = (cols[idx.manualEdit] || '').trim();
                manualEditFlag = v === '是' || v.toLowerCase() === 'yes' || v === '1' || v === 'true';
            }
            let originalPerfParsed = undefined;
            if (idx.originalPerformance >= 0) {
                const v = cols[idx.originalPerformance];
                if (v !== undefined && v !== null && String(v).trim() !== '') originalPerfParsed = parseFloat(v) || undefined;
            }
            let manualEditInfoObj = null;
            if (idx.manualEditInfo >= 0) {
                const raw = cols[idx.manualEditInfo] || '';
                if (raw && raw.trim()) {
                    try { manualEditInfoObj = JSON.parse(raw); } catch(e) { manualEditInfoObj = null; }
                }
            }

            const sig = `${timestamp}|${instId}|${orderType}|${(Number(perf)||0).toFixed(2)}`;
            if (existing.has(sig)) continue;
            existing.add(sig);

            calculationHistory.unshift({
                timestamp: timestamp,
                instrumentId: instId,
                workOrderType: orderType,
                performance: Number(perf) || 0,
                daysFactor: Number.isFinite(days) ? days : 0,
                performanceRatio: String(ratio).replace(/\D/g, '') || '100',
                specialBonus: String(bonus).replace(/\D/g, '') || '0',
                externalDevices: Array.isArray(externalDevicesArr) ? externalDevicesArr : [],
                modificationItems: Array.isArray(modificationItemsArr) ? modificationItemsArr : [],
                maintenance: maintenanceObj,
                calculationProcess: calcProcessObj,
                manualEdit: manualEditFlag,
                originalPerformance: originalPerfParsed,
                manualEditInfo: manualEditInfoObj
            });
            added++;
        }

        if (calculationHistory.length > 50) {
            calculationHistory = calculationHistory.slice(0, 50);
        }
        localStorage.setItem('performanceHistory', JSON.stringify(calculationHistory));
        updateStatsSummary();
        displayHistory();
        return added;
    }

    // 简易CSV解析（支持带引号与逗号）
    function parseCSV(text) {
        const rows = [];
        let row = [];
        let cur = '';
        let inQuotes = false;
        for (let i = 0; i < text.length; i++) {
            const ch = text[i];
            if (inQuotes) {
                if (ch === '"') {
                    if (text[i + 1] === '"') { cur += '"'; i++; }
                    else { inQuotes = false; }
                } else {
                    cur += ch;
                }
            } else {
                if (ch === '"') inQuotes = true;
                else if (ch === ',') { row.push(cur); cur = ''; }
                else if (ch === '\r') { /* skip */ }
                else if (ch === '\n') { row.push(cur); rows.push(row); row = []; cur = ''; }
                else { cur += ch; }
            }
        }
        if (cur.length > 0 || row.length > 0) { row.push(cur); rows.push(row); }
        return rows;
    }

    // 初始化按钮状态（表单默认隐藏）
    if (toggleFormBtn && formContainer && formContainer.classList.contains('hidden')) {
        toggleFormBtn.innerHTML = '<i class="bi bi-eye me-1"></i>显示计算单';
    }

    if (toggleFormBtn && formContainer) {
        toggleFormBtn.addEventListener('click', function() {
            const isHidden = formContainer.classList.toggle('hidden');
            if (isHidden) {
                toggleFormBtn.innerHTML = '<i class="bi bi-eye me-1"></i>显示计算单';
            } else {
                toggleFormBtn.innerHTML = '<i class="bi bi-eye-slash me-1"></i>隐藏计算单';
            }
        });
    }
});