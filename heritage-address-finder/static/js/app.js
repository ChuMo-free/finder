/**
 * 非遗地址坐标查找工具 - 前端JavaScript
 * 功能：Excel上传、字段映射、任务监控、人工审核、地图预览、结果导出
 */

// 全局状态
const state = {
    taskId: null,
    columns: [],
    currentStep: 1,
    statusPollTimer: null,
    amapKey: '',
    currentReviewItem: null,
    map: null,
    marker: null
};

// DOM元素
const elements = {};

// 初始化
document.addEventListener('DOMContentLoaded', function() {
    initElements();
    initEventListeners();
});

// 初始化DOM元素引用
function initElements() {
    // 步骤相关
    elements.step1 = document.getElementById('step-1');
    elements.step2 = document.getElementById('step-2');
    elements.step3 = document.getElementById('step-3');
    elements.step4 = document.getElementById('step-4');
    
    // 面板
    elements.panelUpload = document.getElementById('panel-upload');
    elements.panelMapping = document.getElementById('panel-mapping');
    elements.panelMonitor = document.getElementById('panel-monitor');
    elements.panelResult = document.getElementById('panel-result');
    elements.panelMap = document.getElementById('panel-map');
    
    // 上传相关
    elements.uploadArea = document.getElementById('upload-area');
    elements.fileInput = document.getElementById('file-input');
    elements.btnSelectFile = document.getElementById('btn-select-file');
    elements.fileInfo = document.getElementById('file-info');
    elements.fileName = document.getElementById('file-name');
    elements.fileSize = document.getElementById('file-size');
    elements.fileRows = document.getElementById('file-rows');
    
    // 字段映射
    elements.fieldName = document.getElementById('field-name');
    elements.amapKey = document.getElementById('amap-key');
    elements.btnBackUpload = document.getElementById('btn-back-upload');
    elements.btnStart = document.getElementById('btn-start');
    
    // 监控面板
    elements.taskStatus = document.getElementById('task-status');
    elements.currentPhase = document.getElementById('current-phase');
    elements.progressText = document.getElementById('progress-text');
    elements.searchProgress = document.getElementById('search-progress');
    elements.searchProgressText = document.getElementById('search-progress-text');
    elements.geoProgress = document.getElementById('geo-progress');
    elements.geoProgressText = document.getElementById('geo-progress-text');
    elements.currentItem = document.getElementById('current-item');
    elements.btnPause = document.getElementById('btn-pause');
    elements.btnResume = document.getElementById('btn-resume');
    elements.btnInterrupt = document.getElementById('btn-interrupt');
    elements.btnReturnConfig = document.getElementById('btn-return-config');
    elements.logContainer = document.getElementById('log-container');
    elements.btnClearLog = document.getElementById('btn-clear-log');
    
    // 返回配置弹窗
    elements.modalReturnConfig = document.getElementById('modal-return-config');
    elements.btnCloseReturnModal = document.getElementById('btn-close-return-modal');
    elements.btnCancelReturn = document.getElementById('btn-cancel-return');
    elements.optionSaveProgress = document.getElementById('option-save-progress');
    elements.optionDiscardProgress = document.getElementById('option-discard-progress');
    
    // 结果面板
    elements.resultTotal = document.getElementById('result-total');
    elements.resultSuccess = document.getElementById('result-success');
    elements.resultFail = document.getElementById('result-fail');
    elements.resultTbody = document.getElementById('result-tbody');
    elements.btnExport = document.getElementById('btn-export');
    elements.btnNewTask = document.getElementById('btn-new-task');
    
    // 地图
    elements.mapContainer = document.getElementById('map-container');
    elements.btnCloseMap = document.getElementById('btn-close-map');
    
    // 地址审核弹窗
    elements.modalAddressReview = document.getElementById('modal-address-review');
    elements.reviewItemName = document.getElementById('review-item-name');
    elements.searchResults = document.getElementById('search-results');
    elements.confirmedAddress = document.getElementById('confirmed-address');
    elements.btnCloseAddressModal = document.getElementById('btn-close-address-modal');
    elements.btnRetrySearch = document.getElementById('btn-retry-search');
    elements.btnConfirmAddress = document.getElementById('btn-confirm-address');
    
    // 坐标审核弹窗
    elements.modalCoordReview = document.getElementById('modal-coord-review');
    elements.coordItemName = document.getElementById('coord-item-name');
    elements.coordAddress = document.getElementById('coord-address');
    elements.geoResults = document.getElementById('geo-results');
    elements.confirmedLongitude = document.getElementById('confirmed-longitude');
    elements.confirmedLatitude = document.getElementById('confirmed-latitude');
    elements.btnCloseCoordModal = document.getElementById('btn-close-coord-modal');
    elements.btnRetryGeo = document.getElementById('btn-retry-geo');
    elements.btnConfirmCoord = document.getElementById('btn-confirm-coord');
    elements.btnPreviewMap = document.getElementById('btn-preview-map');
    
    // 任务ID
    elements.taskIdBadge = document.getElementById('task-id-badge');
    elements.taskIdDisplay = document.getElementById('task-id-display');
}

// 初始化事件监听
function initEventListeners() {
    // 文件上传
    elements.btnSelectFile.addEventListener('click', () => elements.fileInput.click());
    elements.uploadArea.addEventListener('click', (e) => {
        if (e.target !== elements.btnSelectFile) {
            elements.fileInput.click();
        }
    });
    elements.fileInput.addEventListener('change', handleFileSelect);
    
    // 拖拽上传
    elements.uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        elements.uploadArea.classList.add('dragover');
    });
    elements.uploadArea.addEventListener('dragleave', () => {
        elements.uploadArea.classList.remove('dragover');
    });
    elements.uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        elements.uploadArea.classList.remove('dragover');
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleFile(files[0]);
        }
    });
    
    // 字段映射
    elements.btnBackUpload.addEventListener('click', () => goToStep(1));
    elements.btnStart.addEventListener('click', startTask);
    
    // 控制按钮
    elements.btnPause.addEventListener('click', pauseTask);
    elements.btnResume.addEventListener('click', resumeTask);
    elements.btnInterrupt.addEventListener('click', interruptTask);
    elements.btnReturnConfig.addEventListener('click', showReturnConfigModal);
    elements.btnClearLog.addEventListener('click', clearLog);
    
    // 返回配置弹窗
    elements.btnCloseReturnModal.addEventListener('click', hideReturnConfigModal);
    elements.btnCancelReturn.addEventListener('click', hideReturnConfigModal);
    elements.optionSaveProgress.addEventListener('click', () => returnToConfig(true));
    elements.optionDiscardProgress.addEventListener('click', () => returnToConfig(false));
    
    // 结果
    elements.btnExport.addEventListener('click', exportExcel);
    elements.btnNewTask.addEventListener('click', resetTask);
    
    // 地图
    elements.btnCloseMap.addEventListener('click', () => {
        elements.panelMap.style.display = 'none';
    });
    
    // 地址审核弹窗
    elements.btnCloseAddressModal.addEventListener('click', () => {
        elements.modalAddressReview.style.display = 'none';
    });
    elements.btnRetrySearch.addEventListener('click', retrySearch);
    elements.btnConfirmAddress.addEventListener('click', confirmAddress);
    
    // 坐标审核弹窗
    elements.btnCloseCoordModal.addEventListener('click', () => {
        elements.modalCoordReview.style.display = 'none';
    });
    elements.btnRetryGeo.addEventListener('click', retryGeo);
    elements.btnConfirmCoord.addEventListener('click', confirmCoord);
    elements.btnPreviewMap.addEventListener('click', previewOnMap);
}

// ==================== 文件上传 ====================

function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file) {
        handleFile(file);
    }
}

function handleFile(file) {
    // 检查文件类型
    const validTypes = ['.xlsx', '.xls'];
    const fileName = file.name.toLowerCase();
    const isValid = validTypes.some(type => fileName.endsWith(type));
    
    if (!isValid) {
        alert('请上传Excel文件（.xlsx 或 .xls 格式）');
        return;
    }
    
    // 上传文件
    const formData = new FormData();
    formData.append('file', file);
    
    // 显示上传中状态
    elements.uploadArea.innerHTML = '<div class="upload-icon">⏳</div><p class="upload-text">正在上传...</p>';
    
    fetch('/api/upload', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            state.taskId = data.task_id;
            state.columns = data.columns;
            
            // 显示文件信息
            elements.fileInfo.style.display = 'flex';
            elements.fileName.textContent = data.filename;
            elements.fileSize.textContent = formatFileSize(file.size);
            elements.fileRows.textContent = `${data.row_count} 行数据`;
            
            // 显示任务ID
            elements.taskIdBadge.style.display = 'block';
            elements.taskIdDisplay.textContent = state.taskId;
            
            // 填充字段下拉框
            populateFieldSelects();
            
            // 进入下一步
            goToStep(2);
        } else {
            alert('上传失败: ' + data.message);
            resetUploadArea();
        }
    })
    .catch(error => {
        console.error('上传错误:', error);
        alert('上传出错，请重试');
        resetUploadArea();
    });
}

function resetUploadArea() {
    elements.uploadArea.innerHTML = `
        <div class="upload-icon">📊</div>
        <p class="upload-text">拖拽Excel文件到此处，或点击选择文件</p>
        <p class="upload-hint">支持 .xlsx / .xls 格式，最大50MB</p>
        <input type="file" id="file-input" accept=".xlsx,.xls" hidden>
        <button class="btn btn-primary" id="btn-select-file">选择文件</button>
    `;
    // 重新绑定事件
    elements.fileInput = document.getElementById('file-input');
    elements.btnSelectFile = document.getElementById('btn-select-file');
    elements.btnSelectFile.addEventListener('click', () => elements.fileInput.click());
    elements.fileInput.addEventListener('change', handleFileSelect);
}

function populateFieldSelects() {
    const select = elements.fieldName;
    select.innerHTML = '<option value="">请选择...</option>';
    
    state.columns.forEach(col => {
        const option = document.createElement('option');
        option.value = col;
        option.textContent = col;
        select.appendChild(option);
    });
    
    // 智能匹配：如果列名包含"名称"、"项目"、"名字"等，自动选中
    for (let col of state.columns) {
        if (/名称|项目|名字|title|name/i.test(col)) {
            select.value = col;
            break;
        }
    }
}

// ==================== 步骤切换 ====================

function goToStep(step) {
    state.currentStep = step;
    
    // 更新步骤指示器
    [1, 2, 3, 4].forEach(n => {
        const stepEl = document.getElementById(`step-${n}`);
        stepEl.classList.remove('active', 'completed');
        if (n < step) {
            stepEl.classList.add('completed');
        } else if (n === step) {
            stepEl.classList.add('active');
        }
    });
    
    // 更新步骤线
    const lines = document.querySelectorAll('.step-line');
    lines.forEach((line, index) => {
        if (index + 1 < step) {
            line.classList.add('completed');
        } else {
            line.classList.remove('completed');
        }
    });
    
    // 显示对应面板
    elements.panelUpload.style.display = step === 1 ? 'block' : 'none';
    elements.panelMapping.style.display = step === 2 ? 'block' : 'none';
    elements.panelMonitor.style.display = step === 3 ? 'block' : 'none';
    elements.panelResult.style.display = step === 4 ? 'block' : 'none';
}

// ==================== 开始任务 ====================

function startTask() {
    const nameField = elements.fieldName.value;
    const amapKey = elements.amapKey.value.trim();
    
    if (!nameField) {
        alert('请选择项目名称字段');
        return;
    }
    
    if (!amapKey) {
        alert('请输入高德地图API Key');
        return;
    }
    
    state.amapKey = amapKey;
    
    // 设置字段映射
    fetch('/api/field-mapping', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            task_id: state.taskId,
            field_mapping: { name: nameField },
            amap_key: amapKey
        })
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            // 启动任务
            return fetch('/api/start', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ task_id: state.taskId })
            });
        } else {
            throw new Error(data.message);
        }
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            goToStep(3);
            startStatusPolling();
        } else {
            throw new Error(data.message);
        }
    })
    .catch(error => {
        alert('启动任务失败: ' + error.message);
    });
}

// ==================== 状态轮询 ====================

function startStatusPolling() {
    if (state.statusPollTimer) {
        clearInterval(state.statusPollTimer);
    }
    
    // 立即获取一次
    fetchStatus();
    
    // 每2秒轮询一次
    state.statusPollTimer = setInterval(fetchStatus, 2000);
}

function stopStatusPolling() {
    if (state.statusPollTimer) {
        clearInterval(state.statusPollTimer);
        state.statusPollTimer = null;
    }
}

function fetchStatus() {
    fetch(`/api/status/${state.taskId}`)
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            updateStatusUI(data.data);
        }
    })
    .catch(error => {
        console.error('获取状态失败:', error);
    });
}

function updateStatusUI(statusData) {
    // 更新状态文本
    elements.taskStatus.textContent = statusData.status;
    elements.currentPhase.textContent = statusData.current_phase || '-';
    elements.progressText.textContent = `${statusData.current_index} / ${statusData.total_items}`;
    elements.currentItem.textContent = statusData.current_item || '-';
    
    // 更新进度条
    elements.searchProgress.style.width = statusData.search_progress + '%';
    elements.searchProgressText.textContent = statusData.search_progress + '%';
    elements.geoProgress.style.width = statusData.geo_progress + '%';
    elements.geoProgressText.textContent = statusData.geo_progress + '%';
    
    // 更新暂停/继续按钮
    if (statusData.paused) {
        elements.btnPause.style.display = 'none';
        elements.btnResume.style.display = 'inline-flex';
    } else {
        elements.btnPause.style.display = 'inline-flex';
        elements.btnResume.style.display = 'none';
    }
    
    // 更新日志
    updateLogs(statusData.logs);
    
    // 检查是否有待审核的项目
    if (statusData.waiting_address_items && statusData.waiting_address_items.length > 0) {
        // 只处理第一个待审核的
        const item = statusData.waiting_address_items[0];
        if (!state.currentReviewItem || state.currentReviewItem.id !== item.id || state.currentReviewItem.type !== 'address') {
            showAddressReviewModal(item);
        }
    }
    
    if (statusData.waiting_coord_items && statusData.waiting_coord_items.length > 0) {
        const item = statusData.waiting_coord_items[0];
        if (!state.currentReviewItem || state.currentReviewItem.id !== item.id || state.currentReviewItem.type !== 'coord') {
            showCoordReviewModal(item);
        }
    }
    
    // 检查任务是否完成
    if (statusData.status === '已完成') {
        stopStatusPolling();
        loadResults();
        goToStep(4);
    }
    
    // 检查任务是否中断
    if (statusData.status === '已中断') {
        stopStatusPolling();
        addLog('任务已中断', 'warning');
    }
}

// ==================== 日志 ====================

function updateLogs(logs) {
    if (!logs || logs.length === 0) return;
    
    const container = elements.logContainer;
    const emptyEl = container.querySelector('.log-empty');
    if (emptyEl) {
        emptyEl.remove();
    }
    
    // 获取当前已有日志数量
    const existingLogs = container.querySelectorAll('.log-entry').length;
    
    // 只添加新日志
    for (let i = existingLogs; i < logs.length; i++) {
        const log = logs[i];
        const entry = document.createElement('div');
        entry.className = 'log-entry';
        entry.innerHTML = `
            <span class="log-time">[${log.time}]</span>
            <span class="log-level-${log.level}">${log.message}</span>
        `;
        container.appendChild(entry);
    }
    
    // 滚动到底部
    container.scrollTop = container.scrollHeight;
}

function addLog(message, level = 'info') {
    const container = elements.logContainer;
    const emptyEl = container.querySelector('.log-empty');
    if (emptyEl) {
        emptyEl.remove();
    }
    
    const time = new Date().toLocaleTimeString('zh-CN', { hour12: false });
    const entry = document.createElement('div');
    entry.className = 'log-entry';
    entry.innerHTML = `
        <span class="log-time">[${time}]</span>
        <span class="log-level-${level}">${message}</span>
    `;
    container.appendChild(entry);
    container.scrollTop = container.scrollHeight;
}

function clearLog() {
    elements.logContainer.innerHTML = '<div class="log-empty">日志已清空...</div>';
}

// ==================== 控制按钮 ====================

function pauseTask() {
    fetch('/api/pause', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ task_id: state.taskId })
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            addLog('已发送暂停指令', 'warning');
        }
    });
}

function resumeTask() {
    fetch('/api/resume', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ task_id: state.taskId })
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            addLog('已发送继续指令', 'info');
        }
    });
}

function interruptTask() {
    if (!confirm('确定要中断任务吗？中断后已处理的数据会保留。')) {
        return;
    }
    
    fetch('/api/interrupt', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ task_id: state.taskId })
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            addLog('已发送中断指令', 'error');
        }
    });
}

// ==================== 返回配置 ====================

function showReturnConfigModal() {
    elements.modalReturnConfig.style.display = 'flex';
}

function hideReturnConfigModal() {
    elements.modalReturnConfig.style.display = 'none';
}

function returnToConfig(saveProgress) {
    hideReturnConfigModal();
    
    const confirmMsg = saveProgress 
        ? '确定要保存进度并返回配置页面吗？' 
        : '确定要丢弃当前进度并重新开始吗？';
    
    if (!confirm(confirmMsg)) {
        return;
    }
    
    fetch('/api/return-to-config', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ 
            task_id: state.taskId,
            save_progress: saveProgress
        })
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            // 停止状态轮询
            stopStatusPolling();
            
            // 关闭可能打开的审核弹窗
            elements.modalAddressReview.style.display = 'none';
            elements.modalCoordReview.style.display = 'none';
            state.currentReviewItem = null;
            
            // 恢复字段映射和API Key
            if (data.field_mapping && data.field_mapping.name) {
                elements.fieldName.value = data.field_mapping.name;
            }
            if (data.amap_key) {
                elements.amapKey.value = data.amap_key;
                state.amapKey = data.amap_key;
            }
            
            // 返回到字段映射页面
            goToStep(2);
            
            addLog(saveProgress ? '已保存进度，返回配置页面' : '已丢弃进度，重新开始', 'info');
        } else {
            alert('返回配置失败: ' + data.message);
        }
    })
    .catch(error => {
        console.error('返回配置错误:', error);
        alert('返回配置失败，请重试');
    });
}

// ==================== 地址审核 ====================

function showAddressReviewModal(item) {
    state.currentReviewItem = { ...item, type: 'address' };
    
    elements.reviewItemName.textContent = item.name;
    elements.confirmedAddress.value = item.recommended_address || '';
    
    // 渲染搜索结果
    const container = elements.searchResults;
    container.innerHTML = '';
    
    if (item.search_results && item.search_results.length > 0) {
        item.search_results.forEach((result, index) => {
            const div = document.createElement('div');
            div.className = 'search-result-item';
            if (result.extracted_address === item.recommended_address) {
                div.classList.add('selected');
            }
            div.innerHTML = `
                <div class="search-result-title">${result.title || '无标题'}</div>
                <div class="search-result-snippet">${result.snippet || ''}</div>
                ${result.extracted_address ? `<div class="search-result-address">📍 ${result.extracted_address}</div>` : ''}
                <div class="search-result-source">来源: ${result.source || '未知'}</div>
            `;
            div.addEventListener('click', () => {
                // 移除其他选中状态
                container.querySelectorAll('.search-result-item').forEach(el => el.classList.remove('selected'));
                div.classList.add('selected');
                // 填充地址
                if (result.extracted_address) {
                    elements.confirmedAddress.value = result.extracted_address;
                }
            });
            container.appendChild(div);
        });
    } else {
        container.innerHTML = '<p style="color: #9ca3af; text-align: center; padding: 20px;">未找到搜索结果，请手动输入地址</p>';
    }
    
    elements.modalAddressReview.style.display = 'flex';
}

function retrySearch() {
    if (!state.currentReviewItem) return;
    
    fetch('/api/retry-item', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            task_id: state.taskId,
            item_id: state.currentReviewItem.id,
            type: 'search'
        })
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            elements.modalAddressReview.style.display = 'none';
            state.currentReviewItem = null;
            addLog('已发送重新搜索指令', 'info');
        }
    });
}

function confirmAddress() {
    const address = elements.confirmedAddress.value.trim();
    
    if (!address) {
        alert('请输入地址');
        return;
    }
    
    if (!state.currentReviewItem) return;
    
    fetch('/api/review-address', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            task_id: state.taskId,
            item_id: state.currentReviewItem.id,
            confirmed_address: address
        })
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            elements.modalAddressReview.style.display = 'none';
            state.currentReviewItem = null;
            addLog(`地址已确认: ${address}`, 'success');
        }
    });
}

// ==================== 坐标审核 ====================

function showCoordReviewModal(item) {
    state.currentReviewItem = { ...item, type: 'coord' };
    
    elements.coordItemName.textContent = item.name;
    elements.coordAddress.textContent = item.confirmed_address;
    
    // 渲染坐标结果
    const container = elements.geoResults;
    container.innerHTML = '';
    
    if (item.geo_results && item.geo_results.length > 0) {
        item.geo_results.forEach((result, index) => {
            const div = document.createElement('div');
            div.className = 'geo-result-item';
            if (item.recommended_coord && 
                result.longitude === item.recommended_coord.longitude && 
                result.latitude === item.recommended_coord.latitude) {
                div.classList.add('selected');
            }
            div.innerHTML = `
                <div class="geo-result-address">${result.formatted_address || '无地址'}</div>
                <div class="geo-result-coord">📍 经度: ${result.longitude}, 纬度: ${result.latitude}</div>
                <div class="geo-result-level">匹配级别: ${result.level || '未知'} | 来源: ${result.source || '未知'}</div>
            `;
            div.addEventListener('click', () => {
                // 移除其他选中状态
                container.querySelectorAll('.geo-result-item').forEach(el => el.classList.remove('selected'));
                div.classList.add('selected');
                // 填充坐标
                elements.confirmedLongitude.value = result.longitude;
                elements.confirmedLatitude.value = result.latitude;
            });
            container.appendChild(div);
        });
        
        // 默认填充第一个结果的坐标
        if (item.recommended_coord) {
            elements.confirmedLongitude.value = item.recommended_coord.longitude;
            elements.confirmedLatitude.value = item.recommended_coord.latitude;
        } else if (item.geo_results[0] && item.geo_results[0].longitude > 0) {
            elements.confirmedLongitude.value = item.geo_results[0].longitude;
            elements.confirmedLatitude.value = item.geo_results[0].latitude;
        }
    } else {
        container.innerHTML = '<p style="color: #9ca3af; text-align: center; padding: 20px;">未找到坐标结果，请手动输入坐标</p>';
    }
    
    elements.modalCoordReview.style.display = 'flex';
}

function retryGeo() {
    if (!state.currentReviewItem) return;
    
    fetch('/api/retry-item', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            task_id: state.taskId,
            item_id: state.currentReviewItem.id,
            type: 'geo'
        })
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            elements.modalCoordReview.style.display = 'none';
            state.currentReviewItem = null;
            addLog('已发送重新转换指令', 'info');
        }
    });
}

function confirmCoord() {
    const longitude = parseFloat(elements.confirmedLongitude.value);
    const latitude = parseFloat(elements.confirmedLatitude.value);
    
    if (isNaN(longitude) || isNaN(latitude)) {
        alert('请输入有效的经纬度');
        return;
    }
    
    if (!state.currentReviewItem) return;
    
    fetch('/api/review-coord', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            task_id: state.taskId,
            item_id: state.currentReviewItem.id,
            longitude: longitude,
            latitude: latitude
        })
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            elements.modalCoordReview.style.display = 'none';
            state.currentReviewItem = null;
            addLog(`坐标已确认: ${longitude}, ${latitude}`, 'success');
        }
    });
}

// ==================== 地图预览 ====================

function previewOnMap() {
    const longitude = parseFloat(elements.confirmedLongitude.value);
    const latitude = parseFloat(elements.confirmedLatitude.value);
    
    if (isNaN(longitude) || isNaN(latitude)) {
        alert('请先输入有效的经纬度');
        return;
    }
    
    // 显示地图面板
    elements.panelMap.style.display = 'block';
    
    // 加载高德地图
    loadAmapScript().then(() => {
        initMap(longitude, latitude);
    }).catch(error => {
        alert('地图加载失败: ' + error.message);
    });
}

function loadAmapScript() {
    return new Promise((resolve, reject) => {
        if (window.AMap) {
            resolve();
            return;
        }
        
        const script = document.createElement('script');
        script.src = `https://webapi.amap.com/maps?v=2.0&key=${state.amapKey}`;
        script.onload = resolve;
        script.onerror = () => reject(new Error('高德地图脚本加载失败'));
        document.head.appendChild(script);
    });
}

function initMap(longitude, latitude) {
    if (!window.AMap) return;
    
    if (state.map) {
        state.map.destroy();
    }
    
    state.map = new AMap.Map('map-container', {
        zoom: 15,
        center: [longitude, latitude]
    });
    
    // 添加标记
    state.marker = new AMap.Marker({
        position: [longitude, latitude],
        title: '坐标位置'
    });
    state.map.add(state.marker);
}

// ==================== 结果展示 ====================

function loadResults() {
    fetch(`/api/items/${state.taskId}`)
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            renderResults(data.items);
        }
    });
}

function renderResults(items) {
    const tbody = elements.resultTbody;
    tbody.innerHTML = '';
    
    let successCount = 0;
    let failCount = 0;
    
    items.forEach((item, index) => {
        const tr = document.createElement('tr');
        
        const isCompleted = item.status === '已完成' || item.longitude;
        if (isCompleted) {
            successCount++;
        } else {
            failCount++;
        }
        
        tr.innerHTML = `
            <td>${index + 1}</td>
            <td>${item.name || '-'}</td>
            <td>${item.confirmed_address || '-'}</td>
            <td>${item.longitude || '-'}</td>
            <td>${item.latitude || '-'}</td>
            <td><span class="status-badge ${isCompleted ? 'completed' : 'pending'}">${isCompleted ? '已完成' : '未完成'}</span></td>
            <td>
                ${item.longitude ? `<button class="btn btn-small btn-secondary" onclick="viewOnMap(${item.longitude}, ${item.latitude})">查看地图</button>` : '-'}
            </td>
        `;
        tbody.appendChild(tr);
    });
    
    elements.resultTotal.textContent = items.length;
    elements.resultSuccess.textContent = successCount;
    elements.resultFail.textContent = failCount;
}

function viewOnMap(longitude, latitude) {
    elements.panelMap.style.display = 'block';
    
    loadAmapScript().then(() => {
        initMap(longitude, latitude);
    }).catch(error => {
        alert('地图加载失败: ' + error.message);
    });
}

// ==================== 导出 ====================

function exportExcel() {
    window.location.href = `/api/export/${state.taskId}`;
}

// ==================== 重置任务 ====================

function resetTask() {
    if (!confirm('确定要新建任务吗？当前任务数据将会丢失。')) {
        return;
    }
    
    stopStatusPolling();
    state.taskId = null;
    state.columns = [];
    state.currentStep = 1;
    state.currentReviewItem = null;
    
    // 重置UI
    elements.fileInfo.style.display = 'none';
    elements.taskIdBadge.style.display = 'none';
    elements.logContainer.innerHTML = '<div class="log-empty">等待任务开始...</div>';
    elements.amapKey.value = '';
    
    resetUploadArea();
    goToStep(1);
}

// ==================== 工具函数 ====================

function formatFileSize(bytes) {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// 暴露给全局的函数（供HTML内联事件使用）
window.viewOnMap = viewOnMap;
