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
    doubaoKey: '',  // 豆包API Key
    doubaoModelId: '',  // 推理接入点ID
    autoMode: false,  // 处理模式：false=半自动，true=全自动
    mode: 'semi',      // 处理模式：semi=半自动，auto=全自动，two_pass=两遍模式
    currentReviewItem: null,
    map: null,
    marker: null,
    resultFilter: 'all',  // 结果筛选：all/success/fail/pending
    selectedProvinces: []  // 已选择的省份列表（空数组=搜索全国）
};

// ==================== 省份数据 ====================
const PROVINCE_LIST = [
    '北京市', '天津市', '上海市', '重庆市',
    '河北省', '山西省', '辽宁省', '吉林省', '黑龙江省',
    '江苏省', '浙江省', '安徽省', '福建省', '江西省', '山东省',
    '河南省', '湖北省', '湖南省', '广东省', '海南省',
    '四川省', '贵州省', '云南省', '陕西省', '甘肃省',
    '青海省', '台湾省',
    '内蒙古自治区', '广西壮族自治区', '西藏自治区',
    '宁夏回族自治区', '新疆维吾尔自治区',
    '香港特别行政区', '澳门特别行政区'
];

// DOM元素
const elements = {};

// 初始化
document.addEventListener('DOMContentLoaded', function() {
    try {
        console.log('[DEBUG] 开始初始化 initElements...');
        initElements();
        console.log('[DEBUG] initElements 完成，开始 initEventListeners...');
        initEventListeners();
        console.log('[DEBUG] initEventListeners 完成，所有初始化成功！');
    } catch(err) {
        console.error('[FATAL] 初始化失败:', err.message, err.stack);
        alert('页面初始化出错: ' + err.message + '\n请按F12查看控制台获取详细信息');
    }
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
    elements.doubaoModelId = document.getElementById('doubao-model-id');
    elements.btnBackUpload = document.getElementById('btn-back-upload');
    elements.btnStart = document.getElementById('btn-start');
    // 模式选择
    elements.modeSemi = document.getElementById('mode-semi');
    elements.modeAuto = document.getElementById('mode-auto');
    elements.modeTwoPass = document.getElementById('mode-two-pass');
    // 两遍模式相关元素
    elements.twoPassNotice = document.getElementById('two-pass-notice');
    elements.pass1Summary = document.getElementById('pass1-summary');
    elements.btnStartPass2 = document.getElementById('btn-start-pass2');
    
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
    elements.btnCancelAddressReview = document.getElementById('btn-cancel-address-review');
    elements.btnConfirmAddress = document.getElementById('btn-confirm-address');
    
    // 坐标审核弹窗
    elements.modalCoordReview = document.getElementById('modal-coord-review');
    elements.coordItemName = document.getElementById('coord-item-name');
    elements.coordAddress = document.getElementById('coord-address');
    elements.geoResults = document.getElementById('geo-results');
    elements.confirmedLongitude = document.getElementById('confirmed-longitude');
    elements.confirmedLatitude = document.getElementById('confirmed-latitude');
    elements.btnCloseCoordModal = document.getElementById('btn-close-coord-modal');
    elements.btnCancelCoordReview = document.getElementById('btn-cancel-coord-review');
    elements.btnRetryGeo = document.getElementById('btn-retry-geo');
    elements.btnConfirmCoord = document.getElementById('btn-confirm-coord');
    elements.btnPreviewMap = document.getElementById('btn-preview-map');
    
    // 任务ID
    elements.taskIdBadge = document.getElementById('task-id-badge');
    elements.taskIdDisplay = document.getElementById('task-id-display');
    
    // 自定义搜索
    elements.customSearchKeyword = document.getElementById('custom-search-keyword');
    elements.btnCustomSearch = document.getElementById('btn-custom-search');
    
    // 加载状态
    elements.searchLoading = document.getElementById('search-loading');
    elements.searchLoadingHint = document.getElementById('search-loading-hint');
    
    // 模式选择
    elements.modeSemi = document.getElementById('mode-semi');
    elements.modeAuto = document.getElementById('mode-auto');
    elements.modeTwoPass = document.getElementById('mode-two-pass');
    
    // 结果筛选
    elements.filterBtns = document.querySelectorAll('.filter-btn');

    // 省份选择器
    elements.provinceSearch = document.getElementById('province-search');
    elements.provinceCheckboxes = document.getElementById('province-checkboxes');
    elements.provinceTags = document.getElementById('province-tags');
    elements.btnClearProvinces = document.getElementById('btn-clear-provinces');
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
    
    // 模式选择
    elements.modeSemi.addEventListener('click', () => selectMode('semi'));
    elements.modeAuto.addEventListener('click', () => selectMode('auto'));
    if (elements.modeTwoPass) {
        elements.modeTwoPass.addEventListener('click', () => selectMode('two_pass'));
    }
    
    // 两遍模式：开始第二遍审核按钮
    if (elements.btnStartPass2) {
        elements.btnStartPass2.addEventListener('click', startPass2);
    }
    
    // 结果筛选
    elements.filterBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const filter = btn.dataset.filter;
            filterResults(filter);
        });
    });
    
    // 字段映射
    elements.btnBackUpload.addEventListener('click', () => goToStep(1));
    elements.btnStart.addEventListener('click', startTask);

    // API Key 眼睛按钮
    const btnToggleKey = document.getElementById('btn-toggle-key');
    if (btnToggleKey) {
        btnToggleKey.addEventListener('click', () => {
            const input = elements.amapKey;
            if (input.type === 'password') {
                input.type = 'text';
                btnToggleKey.textContent = '🙈';
                btnToggleKey.title = '隐藏API Key';
            } else {
                input.type = 'password';
                btnToggleKey.textContent = '👁️';
                btnToggleKey.title = '显示API Key';
            }
        });
    }

    // 记住 API Key
    const btnSaveKey = document.getElementById('btn-save-key');
    if (btnSaveKey) {
        btnSaveKey.addEventListener('click', (e) => {
            e.preventDefault();
            const key = elements.amapKey.value.trim();
            if (!key) {
                showToast('请先输入API Key再保存', 'warning');
                return;
            }
            localStorage.setItem('amap_key_saved', key);
            showToast('API Key 已保存到本地', 'success', 2000);
        });
    }

    // 豆包API Key 眼睛按钮
    const btnToggleDoubaoKey = document.getElementById('btn-toggle-doubao-key');
    if (btnToggleDoubaoKey) {
        btnToggleDoubaoKey.addEventListener('click', () => {
            const input = document.getElementById('doubao-key');
            if (input.type === 'password') {
                input.type = 'text';
                btnToggleDoubaoKey.textContent = '🙈';
            } else {
                input.type = 'password';
                btnToggleDoubaoKey.textContent = '👁️';
            }
        });
    }

    // 记住豆包 API Key
    const btnSaveDoubaoKey = document.getElementById('btn-save-doubao-key');
    if (btnSaveDoubaoKey) {
        btnSaveDoubaoKey.addEventListener('click', (e) => {
            e.preventDefault();
            const key = document.getElementById('doubao-key').value.trim();
            if (!key) {
                showToast('请先输入豆包API Key再保存', 'warning');
                return;
            }
            localStorage.setItem('doubao_key_saved', key);
            showToast('豆包API Key 已保存到本地', 'success', 2000);
        });
    }

    // 清除豆包 API Key
    const btnClearDoubaoKey = document.getElementById('btn-clear-doubao-key');
    if (btnClearDoubaoKey) {
        btnClearDoubaoKey.addEventListener('click', (e) => {
            e.preventDefault();
            localStorage.removeItem('doubao_key_saved');
            document.getElementById('doubao-key').value = '';
            showToast('豆包API Key 已清除', 'info');
        });
    }

    // 记住推理接入点ID
    const btnSaveDoubaoModel = document.getElementById('btn-save-doubao-model');
    if (btnSaveDoubaoModel) {
        btnSaveDoubaoModel.addEventListener('click', (e) => {
            e.preventDefault();
            const modelId = document.getElementById('doubao-model-id').value.trim();
            if (!modelId) {
                showToast('请先输入推理接入点ID再保存', 'warning');
                return;
            }
            localStorage.setItem('doubao_model_id_saved', modelId);
            showToast('推理接入点ID 已保存到本地', 'success', 2000);
        });
    }

    // 清除推理接入点ID
    const btnClearDoubaoModel = document.getElementById('btn-clear-doubao-model');
    if (btnClearDoubaoModel) {
        btnClearDoubaoModel.addEventListener('click', (e) => {
            e.preventDefault();
            localStorage.removeItem('doubao_model_id_saved');
            document.getElementById('doubao-model-id').value = '';
            showToast('推理接入点ID 已清除', 'info');
        });
    }

    // 清除记忆的 API Key
    const btnClearKey = document.getElementById('btn-clear-key');
    if (btnClearKey) {
        btnClearKey.addEventListener('click', (e) => {
            e.preventDefault();
            localStorage.removeItem('amap_key_saved');
            showToast('已清除记忆的 API Key', 'info', 1800);
        });
    }

    // 移除文件按钮
    const btnRemoveFile = document.getElementById('btn-remove-file');
    if (btnRemoveFile) {
        btnRemoveFile.addEventListener('click', (e) => {
            e.stopPropagation();
            elements.fileInfo.style.display = 'none';
            state.taskId = null;
            state.columns = [];
            resetUploadArea();
        });
    }

    // 加载记忆的 API Key
    const savedKey = localStorage.getItem('amap_key_saved');
    if (savedKey && elements.amapKey) {
        elements.amapKey.value = savedKey;
    }

    // 加载记忆的推理接入点ID
    const savedModelId = localStorage.getItem('doubao_model_id_saved');
    if (savedModelId && elements.doubaoModelId) {
        elements.doubaoModelId.value = savedModelId;
    }
    
    // 控制按钮
    elements.btnPause.addEventListener('click', pauseTask);
    elements.btnResume.addEventListener('click', resumeTask);
    elements.btnInterrupt.addEventListener('click', interruptTask);
    elements.btnReturnConfig.addEventListener('click', showReturnConfigModal);
    elements.btnClearLog.addEventListener('click', clearLog);

    // 日志滚动到最新
    const btnScrollLog = document.getElementById('btn-scroll-log');
    if (btnScrollLog) {
        btnScrollLog.addEventListener('click', () => {
            elements.logContainer.scrollTop = elements.logContainer.scrollHeight;
        });
    }
    
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
    
    // 地址审核弹窗 - 移除右上角关闭按钮的直接关闭功能
    elements.btnCloseAddressModal.addEventListener('click', async () => {
        const ok = await showConfirm('确定要取消审核吗？取消后可以稍后继续处理。', '取消审核', '继续');
        if (ok) closeAddressReviewModal();
    });
    // 取消按钮
    if (elements.btnCancelAddressReview) {
        elements.btnCancelAddressReview.addEventListener('click', async () => {
            const ok = await showConfirm('确定要取消审核吗？取消后可以稍后继续处理。', '取消审核', '继续');
            if (ok) closeAddressReviewModal();
        });
    }
    elements.btnRetrySearch.addEventListener('click', retrySearch);
    elements.btnCustomSearch.addEventListener('click', customSearch);
    elements.customSearchKeyword.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            customSearch();
        }
    });
    elements.btnConfirmAddress.addEventListener('click', confirmAddress);
    
    // 阻止点击弹窗外部区域关闭弹窗
    const addressOverlay = elements.modalAddressReview.querySelector('.modal-overlay');
    if (addressOverlay) {
        addressOverlay.addEventListener('click', (e) => {
            e.stopPropagation();
            e.preventDefault();
            showToast('请使用弹窗底部的按钮进行操作', 'info', 1800);
        });
    }
    
    // 阻止ESC键关闭弹窗
    document.addEventListener('keydown', async (e) => {
        if (e.key === 'Escape' && elements.modalAddressReview.style.display === 'flex') {
            e.preventDefault();
            e.stopPropagation();
            const ok = await showConfirm('确定要取消审核吗？取消后可以稍后继续处理。', '取消审核', '继续');
            if (ok) closeAddressReviewModal();
        }
    });
    
    // 坐标审核弹窗 - 移除右上角关闭按钮的直接关闭功能
    elements.btnCloseCoordModal.addEventListener('click', async () => {
        const ok = await showConfirm('确定要取消审核吗？取消后可以稍后继续处理。', '取消审核', '继续');
        if (ok) closeCoordReviewModal();
    });
    // 取消按钮
    if (elements.btnCancelCoordReview) {
        elements.btnCancelCoordReview.addEventListener('click', async () => {
            const ok = await showConfirm('确定要取消审核吗？取消后可以稍后继续处理。', '取消审核', '继续');
            if (ok) closeCoordReviewModal();
        });
    }
    elements.btnRetryGeo.addEventListener('click', retryGeo);
    elements.btnConfirmCoord.addEventListener('click', confirmCoord);
    elements.btnPreviewMap.addEventListener('click', previewOnMap);
    
    // 阻止点击坐标审核弹窗外部区域关闭弹窗
    const coordOverlay = elements.modalCoordReview.querySelector('.modal-overlay');
    if (coordOverlay) {
        coordOverlay.addEventListener('click', (e) => {
            e.stopPropagation();
            e.preventDefault();
            showToast('请使用弹窗底部的按钮进行操作', 'info', 1800);
        });
    }
    
    // 阻止ESC键关闭坐标审核弹窗
    document.addEventListener('keydown', async (e) => {
        if (e.key === 'Escape' && elements.modalCoordReview.style.display === 'flex') {
            e.preventDefault();
            e.stopPropagation();
            const ok = await showConfirm('确定要取消审核吗？取消后可以稍后继续处理。', '取消审核', '继续');
            if (ok) closeCoordReviewModal();
        }
    });

    // 初始化省份选择器
    initProvinceSelector();
    // 绑定AI搜索按钮事件
    bindAISearchButton();
}

// ==================== 省份选择器 ====================

function initProvinceSelector() {
    if (!elements.provinceCheckboxes) return;

    // 渲染所有省份复选框
    elements.provinceCheckboxes.innerHTML = PROVINCE_LIST.map(p => `
        <label class="province-checkbox-item" data-province="${p}">
            <input type="checkbox" value="${p}" aria-label="${p}">
            <span>${p.replace('省','').replace('市','').replace('自治区','').replace('壮族','').replace('回族','').replace('维吾尔','').replace('特别行政区','')}</span>
        </label>
    `).join('');

    // 勾选事件：委托到容器
    elements.provinceCheckboxes.addEventListener('change', (e) => {
        if (e.target.type === 'checkbox') {
            const prov = e.target.value;
            if (e.target.checked) {
                if (!state.selectedProvinces.includes(prov)) {
                    state.selectedProvinces.push(prov);
                }
            } else {
                state.selectedProvinces = state.selectedProvinces.filter(p => p !== prov);
            }
            updateProvinceUI();
        }
    });

    // 搜索过滤
    if (elements.provinceSearch) {
        elements.provinceSearch.addEventListener('input', (e) => {
            const q = e.target.value.trim();
            elements.provinceCheckboxes.querySelectorAll('.province-checkbox-item').forEach(item => {
                const name = item.dataset.province;
                item.classList.toggle('hidden', q && !name.includes(q));
            });
        });
    }

    // 清空按钮
    if (elements.btnClearProvinces) {
        elements.btnClearProvinces.addEventListener('click', () => {
            state.selectedProvinces = [];
            // 取消所有勾选
            elements.provinceCheckboxes.querySelectorAll('input[type="checkbox"]').forEach(cb => {
                cb.checked = false;
            });
            elements.provinceCheckboxes.querySelectorAll('.province-checkbox-item').forEach(item => {
                item.classList.remove('checked');
            });
            updateProvinceUI();
        });
    }

    updateProvinceUI();
}

function updateProvinceUI() {
    if (!elements.provinceTags) return;

    // 同步复选框勾选状态
    if (elements.provinceCheckboxes) {
        elements.provinceCheckboxes.querySelectorAll('.province-checkbox-item').forEach(item => {
            const prov = item.dataset.province;
            const cb = item.querySelector('input[type="checkbox"]');
            const selected = state.selectedProvinces.includes(prov);
            if (cb) cb.checked = selected;
            item.classList.toggle('checked', selected);
        });
    }

    // 渲染标签
    if (state.selectedProvinces.length === 0) {
        elements.provinceTags.innerHTML = '<span class="province-empty-hint">未限定（搜索全国）</span>';
        if (elements.btnClearProvinces) elements.btnClearProvinces.style.display = 'none';
    } else {
        elements.provinceTags.innerHTML = state.selectedProvinces.map(p => {
            const shortName = p.replace('省','').replace('市','').replace('自治区','')
                              .replace('壮族','').replace('回族','').replace('维吾尔','')
                              .replace('特别行政区','');
            return `<span class="province-tag" data-province="${p}">
                ${shortName}
                <button class="province-tag-remove" data-province="${p}" title="移除 ${p}" type="button">×</button>
            </span>`;
        }).join('');
        if (elements.btnClearProvinces) elements.btnClearProvinces.style.display = '';

        // 标签上的移除按钮
        elements.provinceTags.querySelectorAll('.province-tag-remove').forEach(btn => {
            btn.addEventListener('click', () => {
                const prov = btn.dataset.province;
                state.selectedProvinces = state.selectedProvinces.filter(p => p !== prov);
                updateProvinceUI();
            });
        });
    }
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
        showToast('请上传Excel文件（.xlsx 或 .xls 格式）', 'error');
        return;
    }

    // 检查文件大小（50MB限制）
    if (file.size > 50 * 1024 * 1024) {
        showToast('文件超过50MB限制，请分批处理', 'error');
        return;
    }
    
    // ★ 先保存进度条元素引用（在替换innerHTML之前！）
    const progressEl = document.getElementById('upload-progress');
    const progressFill = document.getElementById('upload-progress-fill');
    const progressText = document.getElementById('upload-progress-text');
    
    // 上传文件
    const formData = new FormData();
    formData.append('file', file);
    
    // 显示上传中状态（带进度条）
    elements.uploadArea.innerHTML = `
        <div class="upload-icon">⏳</div>
        <p class="upload-text">正在上传文件...</p>
    `;
    
    // 使用XHR上传以支持进度（进度条元素已在上方保存引用）
    if (progressEl) {
        progressEl.style.display = 'block';
        uploadWithProgress(formData, progressFill, progressText);
    } else {
        //  fallback：无进度条时使用普通fetch
        fetch('/api/upload', {
            method: 'POST',
            body: formData
        })
        .then(response => response.json())
        .then(data => handleUploadResponse(data, file))
        .catch(error => {
            console.error('上传错误:', error);
            showToast('上传出错，请重试', 'error');
            resetUploadArea();
        });
    }
}

function uploadWithProgress(formData, progressFill, progressText) {
    const xhr = new XMLHttpRequest();
    // 保存 xhr 引用，便于用户取消
    window._uploadXHR = xhr;

    // 延长超时到 5 分钟（300秒），给服务器足够时间读取大 Excel
    xhr.timeout = 300000;

    xhr.upload.addEventListener('progress', (e) => {
        if (e.lengthComputable) {
            const pct = Math.round((e.loaded / e.total) * 100);
            if (progressFill) progressFill.style.width = pct + '%';
            if (progressText) progressText.textContent = `正在上传... ${pct}%`;
        }
    });

    // 上传完成后（100%），切换为"等待服务器"状态
    xhr.upload.addEventListener('load', () => {
        if (progressFill) progressFill.style.width = '100%';
        // 显示更明确的提示：文件已传到服务器，正在解析 Excel
        if (progressText) {
            progressText.textContent = '✅ 文件已上传，正在解析 Excel（约 10-30 秒），请勿关闭页面...';
            progressText.style.fontWeight = '600';
        }
        // 添加取消按钮（让用户有退路）
        const cancelBtn = document.createElement('button');
        cancelBtn.className = 'btn btn-text btn-small';
        cancelBtn.textContent = '取消上传';
        cancelBtn.style.marginTop = '8px';
        cancelBtn.type = 'button';
        cancelBtn.addEventListener('click', () => {
            try { xhr.abort(); } catch (e) {}
            showToast('已取消上传', 'info');
            resetUploadArea();
        });
        // 插入到 progress-text 附近（如果 progress-text 有父元素）
        if (progressText && progressText.parentNode) {
            progressText.parentNode.appendChild(cancelBtn);
        }
    });

    xhr.addEventListener('load', () => {
        // 清理取消按钮引用
        window._uploadXHR = null;
        const progressEl = document.getElementById('upload-progress');
        if (progressEl) progressEl.style.display = 'none';

        if (xhr.status === 200) {
            try {
                console.log('[上传] 服务器响应:', xhr.responseText.substring(0, 200));
                const data = JSON.parse(xhr.responseText);
                const file = formData.get('file');
                handleUploadResponse(data, file);
            } catch (e) {
                // 响应不是合法 JSON —— 可能返回了 Flask 的 HTML 错误页
                console.error('[上传] 解析响应失败:', e, '原始响应:', xhr.responseText);
                const firstLine = (xhr.responseText || '').substring(0, 200).replace(/<[^>]*>/g, '');
                showToast(`服务器返回异常，请重试（${firstLine || '未知错误'}）`, 'error');
                resetUploadArea();
            }
        } else {
            console.error('[上传] HTTP错误:', xhr.status, xhr.responseText);
            let msg = `上传失败（HTTP ${xhr.status}）`;
            // 尝试解析后端的 JSON 错误信息
            try {
                const data = JSON.parse(xhr.responseText);
                if (data && data.message) msg = data.message;
            } catch (e) {}
            showToast(msg, 'error');
            resetUploadArea();
        }
    });

    xhr.addEventListener('error', () => {
        console.error('[上传] 网络错误');
        window._uploadXHR = null;
        const progressEl = document.getElementById('upload-progress');
        if (progressEl) progressEl.style.display = 'none';
        showToast('网络错误，请检查连接后重试', 'error');
        resetUploadArea();
    });

    // ★ 超时处理
    xhr.addEventListener('timeout', () => {
        console.error('[上传] 请求超时(300s)');
        window._uploadXHR = null;
        const progressEl = document.getElementById('upload-progress');
        if (progressEl) progressEl.style.display = 'none';
        showToast('请求超时，文件可能过大，请分批处理或重试', 'error');
        resetUploadArea();
    });

    xhr.open('POST', '/api/upload');
    xhr.send(formData);
}

function handleUploadResponse(data, file) {
    if (data.success) {
        state.taskId = data.task_id;
        state.columns = data.columns;
        
        // 显示文件信息
        elements.fileInfo.style.display = 'flex';
        elements.fileName.textContent = data.filename;
        elements.fileSize.textContent = formatFileSize(typeof file === 'object' && file.size ? file.size : 0);
        elements.fileRows.textContent = `${data.row_count} 行数据`;
        
        // 显示任务ID
        elements.taskIdBadge.style.display = 'block';
        elements.taskIdDisplay.textContent = state.taskId;
        
        // 填充字段下拉框
        populateFieldSelects();
        
        // 进入下一步
        goToStep(2);

        showToast(`文件上传成功，共 ${data.row_count} 行数据`, 'success');
    } else {
        showToast('上传失败: ' + (data.message || '未知错误'), 'error');
        resetUploadArea();
    }
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

// ==================== 模式选择 ====================
function selectMode(mode) {
    state.mode = mode;
    state.autoMode = (mode === 'auto' || mode === 'two_pass');
    
    // 更新UI
    if (elements.modeSemi && elements.modeAuto && elements.modeTwoPass) {
        elements.modeSemi.classList.toggle('selected', mode === 'semi');
        elements.modeAuto.classList.toggle('selected', mode === 'auto');
        elements.modeTwoPass.classList.toggle('selected', mode === 'two_pass');
    }
}

// ==================== 两遍模式：开始第二遍审核 ====================
function startPass2() {
    if (!state.taskId) return;
    fetch('/api/start_pass2', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ task_id: state.taskId })
    })
    .then(r => r.json())
    .then(d => {
        if (d.success) {
            showToast('第二遍审核已启动', 'success');
            if (elements.twoPassNotice) {
                elements.twoPassNotice.style.display = 'none';
            }
        } else {
            showToast('启动第二遍审核失败：' + (d.message || '未知错误'), 'error');
        }
    })
    .catch(err => {
        showToast('启动第二遍审核出错：' + err.message, 'error');
    });
}

// ==================== 开始任务 ====================


function startTask() {
    const nameField = elements.fieldName.value;
    const amapKey = elements.amapKey.value.trim();
    const doubaoKey = document.getElementById('doubao-key').value.trim();
    const doubaoModelId = elements.doubaoModelId ? elements.doubaoModelId.value.trim() : '';
    
    if (!nameField) {
        showToast('请选择项目名称字段', 'warning');
        elements.fieldName.focus();
        return;
    }
    
    if (!amapKey) {
        showToast('请输入高德地图API Key', 'warning');
        elements.amapKey.focus();
        return;
    }
    
    state.amapKey = amapKey;

    // 添加"正在加载中"的提示，避免用户以为卡死
    const originalBtn = document.getElementById('btn-start-task');
    let btnOriginalText = '';
    if (originalBtn) {
        btnOriginalText = originalBtn.textContent;
        originalBtn.disabled = true;
        originalBtn.textContent = '⏳ 正在解析数据，请稍候...';
    }

    // 超时控制器：5 分钟超时（给大 Excel 足够时间）
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 300000);

    // 设置字段映射
    fetch('/api/field-mapping', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        signal: controller.signal,
        body: JSON.stringify({
            task_id: state.taskId,
            field_mapping: { name: nameField },
            amap_key: amapKey,
            doubao_api_key: doubaoKey || undefined,
            doubao_model_id: doubaoModelId || undefined,
            auto_mode: state.autoMode,
            mode: state.mode,
            provinces: state.selectedProvinces
        })
    })
    .then(response => {
        clearTimeout(timeoutId);
        if (!response.ok) throw new Error(`服务器响应 ${response.status}`);
        return response.json();
    })
    .then(data => {
        if (data.success) {
            // 启动任务
            return fetch('/api/start', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ task_id: state.taskId })
            })
            .then(r => r.json())
            .then(d => {
                if (d.success) {
                    if (originalBtn) {
                        originalBtn.disabled = false;
                        originalBtn.textContent = btnOriginalText;
                    }
                    goToStep(3);
                    startStatusPolling();
                    showToast('任务已启动', 'success');
                } else {
                    throw new Error(d.message || '启动失败');
                }
            });
        } else {
            throw new Error(data.message || '字段映射失败');
        }
    })
    .catch(err => {
        clearTimeout(timeoutId);
        if (originalBtn) {
            originalBtn.disabled = false;
            originalBtn.textContent = btnOriginalText;
        }
        if (err.name === 'AbortError') {
            showToast('请求超时，文件可能过大，请分批处理', 'error');
        } else {
            showToast('操作失败：' + (err.message || '未知错误'), 'error');
        }
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
    
    // 更新模式状态
    if (statusData.auto_mode !== undefined) {
        state.autoMode = statusData.auto_mode;
    }
    
    // 更新按钮可用状态（确保所有控制按钮都可用）
    // 处理中：暂停、中断、返回配置都可用
    // 暂停时：继续、中断、返回配置都可用
    // 等待审核时：所有按钮都可用
    const isRunning = statusData.status !== '已完成' && statusData.status !== '出错';
    elements.btnInterrupt.disabled = !isRunning;
    elements.btnReturnConfig.disabled = !isRunning;
    elements.btnPause.disabled = !isRunning || statusData.paused;
    elements.btnResume.disabled = !isRunning || !statusData.paused;
    
    // 更新按钮样式（禁用状态）
    [elements.btnPause, elements.btnResume, elements.btnInterrupt, elements.btnReturnConfig].forEach(btn => {
        if (btn.disabled) {
            btn.classList.add('btn-disabled');
            btn.style.opacity = '0.5';
            btn.style.cursor = 'not-allowed';
        } else {
            btn.classList.remove('btn-disabled');
            btn.style.opacity = '1';
            btn.style.cursor = 'pointer';
        }
    });
    
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
    
    // 全自动模式下实时更新结果列表
    if (statusData.auto_mode && state.currentStep === 3) {
        loadResults();
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
        // 中断后也加载结果
        loadResults();
        goToStep(4);
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
    showConfirm('确定要中断任务吗？中断后已处理的数据会保留。', '中断任务', '取消').then(ok => {
        if (!ok) return;

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
    
    showConfirm(confirmMsg, '确定', '取消').then(ok => {
        if (!ok) return;

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
                stopStatusPolling();
                elements.modalAddressReview.style.display = 'none';
                elements.modalCoordReview.style.display = 'none';
                state.currentReviewItem = null;
                
                if (data.field_mapping && data.field_mapping.name) {
                    elements.fieldName.value = data.field_mapping.name;
                }
                if (data.amap_key) {
                    elements.amapKey.value = data.amap_key;
                    if (data.doubao_model_id && elements.doubaoModelId) {
                        elements.doubaoModelId.value = data.doubao_model_id;
                    }
                    state.amapKey = data.amap_key;
                }
                if (data.auto_mode !== undefined) {
                    state.autoMode = data.auto_mode;
                    selectMode(data.auto_mode ? 'auto' : 'semi');
                }
                // 恢复省份选择
                if (data.provinces && Array.isArray(data.provinces)) {
                    state.selectedProvinces = [...data.provinces];
                    updateProvinceUI();
                }
                
                goToStep(2);
                addLog(saveProgress ? '已保存进度，返回配置页面' : '已丢弃进度，重新开始', 'info');
            } else {
                showToast('返回配置失败: ' + (data.message || '未知错误'), 'error');
            }
        })
        .catch(error => {
            console.error('返回配置错误:', error);
            showToast('返回配置失败，请重试', 'error');
        });
    });
}

// ==================== 地址审核 ====================


// 关键词高亮
function highlightKeyword(text, keyword) {
    if (!text || !keyword) return text || '';
    try {
        const regex = new RegExp(`(${keyword.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})`, 'gi');
        return text.replace(regex, '<mark class="highlight">$1</mark>');
    } catch (e) {
        return text;
    }
}

function showAddressReviewModal(item) {
    state.currentReviewItem = { ...item, type: 'address' };
    
    elements.reviewItemName.textContent = item.name;
    elements.confirmedAddress.value = item.recommended_address || '';
    
    // 清除搜索超时定时器
    clearSearchTimeout();
    
    // 隐藏加载状态
    hideSearchLoading();
    
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
            
            // 来源类型CSS类
            const sourceType = result.source_type || 'other';
            const sourceClass = `source-${sourceType}`;
            
            // 来源URL
            const sourceUrl = result.source_url || result.link || '';
            const sourceLink = sourceUrl 
                ? `<a href="${sourceUrl}" target="_blank" class="source-link" onclick="event.stopPropagation()">🔗 查看来源</a>` 
                : '';
            
            // 关键词高亮
            const highlightedTitle = highlightKeyword(result.title || '无标题', item.name);
            const highlightedSnippet = highlightKeyword(result.snippet || '', item.name);
            
            // 显示来源URL（域名部分）
            let sourceUrlDisplay = '';
            if (sourceUrl) {
                try {
                    const url = new URL(sourceUrl);
                    sourceUrlDisplay = `<span class="source-url">${url.hostname}</span>`;
                } catch (e) {
                    sourceUrlDisplay = '';
                }
            }
            
            div.innerHTML = `
                <div class="search-result-title">${highlightedTitle}</div>
                <div class="search-result-snippet">${highlightedSnippet}</div>
                ${result.extracted_address ? `<div class="search-result-address">📍 ${result.extracted_address}</div>` : ''}
                <div class="search-result-source">
                    <span class="source-name ${sourceClass}">${result.source_name || result.source || '未知'}</span>
                    ${sourceUrlDisplay}
                    ${sourceLink}
                </div>
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


// 显示搜索加载状态
function showSearchLoading(text) {
    if (elements.searchLoading) {
        elements.searchLoading.style.display = 'flex';
        // 更新加载文本
        const loadingText = elements.searchLoading.querySelector('.loading-text');
        if (loadingText && text) {
            loadingText.textContent = text;
        }
    }
    if (elements.searchResults) {
        elements.searchResults.style.opacity = '0.3';
    }
    // 禁用按钮
    if (elements.btnRetrySearch) {
        elements.btnRetrySearch.disabled = true;
        elements.btnRetrySearch.style.opacity = '0.5';
    }
    if (elements.btnCustomSearch) {
        elements.btnCustomSearch.disabled = true;
        elements.btnCustomSearch.style.opacity = '0.5';
    }
    
    // 10秒后显示超时提示
    if (elements.searchLoadingHint) {
        setTimeout(() => {
            if (elements.searchLoading && elements.searchLoading.style.display !== 'none') {
                elements.searchLoadingHint.style.display = 'block';
            }
        }, 10000);
    }
}

// 隐藏搜索加载状态
function hideSearchLoading() {
    if (elements.searchLoading) {
        elements.searchLoading.style.display = 'none';
    }
    if (elements.searchResults) {
        elements.searchResults.style.opacity = '1';
    }
    // 启用按钮
    if (elements.btnRetrySearch) {
        elements.btnRetrySearch.disabled = false;
        elements.btnRetrySearch.style.opacity = '1';
    }
    if (elements.btnCustomSearch) {
        elements.btnCustomSearch.disabled = false;
        elements.btnCustomSearch.style.opacity = '1';
    }
    // 重置超时提示
    if (elements.searchLoadingHint) {
        elements.searchLoadingHint.style.display = 'none';
    }
}

// 搜索超时定时器
let searchTimeoutTimer = null;

function retrySearch() {
    if (!state.currentReviewItem) return;
    
    // 显示加载状态
    showSearchLoading('正在重新搜索...');
    
    // 设置15秒超时
    setupSearchTimeout();
    
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
            addLog('已发送重新搜索指令', 'info');
            // 不关闭弹窗，等待新结果回来
            // 状态轮询会自动更新弹窗内容
        } else {
            throw new Error(data.message || '重新搜索失败');
        }
    })
    .catch(error => {
        console.error('重新搜索失败:', error);
        hideSearchLoading();
        clearSearchTimeout();
        // 区分网络错误和其他错误
        let errorMsg = error.message || '请求失败，请重试';
        if (error.name === 'TypeError' || error.message.includes('Failed to fetch') || error.message.includes('NetworkError')) {
            errorMsg = '网络连接失败，请检查网络后重试';
        }
        showSearchError(errorMsg);
    });
}

function customSearch() {
    if (!state.currentReviewItem) return;
    
    const keyword = elements.customSearchKeyword.value.trim();
    if (!keyword) {
        showToast('请输入搜索关键词', 'warning');
        elements.customSearchKeyword.focus();
        return;
    }
    
    // 显示加载状态
    showSearchLoading(`正在使用关键词"${keyword}"搜索...`);
    
    // 设置15秒超时
    setupSearchTimeout();
    
    fetch('/api/retry-item', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            task_id: state.taskId,
            item_id: state.currentReviewItem.id,
            type: 'search',
            custom_keyword: keyword
        })
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            addLog(`已使用关键词"${keyword}"重新搜索`, 'info');
            // 不关闭弹窗，等待新结果回来
        } else {
            throw new Error(data.message || '搜索失败');
        }
    })
    .catch(error => {
        console.error('自定义搜索失败:', error);
        hideSearchLoading();
        clearSearchTimeout();
        // 区分网络错误和其他错误
        let errorMsg = error.message || '请求失败，请重试';
        if (error.name === 'TypeError' || error.message.includes('Failed to fetch') || error.message.includes('NetworkError')) {
            errorMsg = '网络连接失败，请检查网络后重试';
        }
        showSearchError(errorMsg);
    });
}

// 设置搜索超时
function setupSearchTimeout() {
    clearSearchTimeout();
    searchTimeoutTimer = setTimeout(() => {
        if (elements.searchLoading && elements.searchLoading.style.display !== 'none') {
            hideSearchLoading();
            showSearchError('搜索超时，请检查网络连接后重试');
        }
    }, 15000); // 15秒超时
}

// 清除搜索超时
function clearSearchTimeout() {
    if (searchTimeoutTimer) {
        clearTimeout(searchTimeoutTimer);
        searchTimeoutTimer = null;
    }
}

// 显示搜索错误
function showSearchError(errorMessage) {
    const container = elements.searchResults;
    container.innerHTML = `
        <div style="text-align: center; padding: 40px 20px;">
            <div style="font-size: 48px; margin-bottom: 16px;">⚠️</div>
            <div style="color: #ef4444; font-size: 16px; margin-bottom: 12px;">搜索失败</div>
            <div style="color: #6b7280; font-size: 14px; margin-bottom: 20px;">${escapeHtml(errorMessage)}</div>
            <button class="btn btn-primary" onclick="retrySearch()">🔄 重试</button>
        </div>
    `;
    container.style.opacity = '1';
}

// 关闭地址审核弹窗（统一入口）
function closeAddressReviewModal() {
    elements.modalAddressReview.style.display = 'none';
    state.currentReviewItem = null;
}

function confirmAddress() {
    const address = elements.confirmedAddress.value.trim();
    
    if (!address) {
        showToast('请输入确认地址', 'warning');
        elements.confirmedAddress.focus();
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
            closeAddressReviewModal();
            addLog(`地址已确认: ${address}`, 'success');
            showToast('地址已确认', 'success', 1500);
        }
    });
}

// ==================== Toast 提示（使用CSS动画版）====================
/**
 * 显示 Toast 提示
 * @param {string} message - 提示文字
 * @param {'info'|'success'|'warning'|'error'} type - 类型
 * @param {number} duration - 显示时长(ms)
 */
function showToast(message, type = 'info', duration = 2500) {
    let container = document.getElementById('toast-container');
    if (!container) {
        container = document.createElement('div');
        container.id = 'toast-container';
        document.body.appendChild(container);
    }

    const toast = document.createElement('div');
    toast.className = `toast-item${type !== 'info' ? ' ' + type : ''}`;
    toast.textContent = message;
    container.appendChild(toast);

    setTimeout(() => {
        toast.classList.add('hiding');
        setTimeout(() => toast.remove(), 320);
    }, duration);
}

// ==================== 自定义确认弹窗（替代原生 confirm）====================
/**
 * 显示自定义确认对话框，返回 Promise<boolean>
 */
function showConfirm(message, okText = '确认', cancelText = '取消') {
    return new Promise((resolve) => {
        const modal = document.getElementById('modal-confirm');
        const msgEl = document.getElementById('confirm-modal-message');
        const okBtn = document.getElementById('btn-confirm-ok');
        const cancelBtn = document.getElementById('btn-confirm-cancel');

        if (!modal || !msgEl) {
            // 降级到原生
            resolve(confirm(message));
            return;
        }

        msgEl.textContent = message;
        okBtn.textContent = okText;
        cancelBtn.textContent = cancelText;
        modal.style.display = 'flex';

        const cleanup = (result) => {
            modal.style.display = 'none';
            okBtn.removeEventListener('click', onOk);
            cancelBtn.removeEventListener('click', onCancel);
            resolve(result);
        };

        const onOk = () => cleanup(true);
        const onCancel = () => cleanup(false);
        okBtn.addEventListener('click', onOk);
        cancelBtn.addEventListener('click', onCancel);
    });
}

// ==================== HTML 转义（XSS防护）====================
function escapeHtml(str) {
    if (!str) return '';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}
function closeCoordReviewModal() {
    elements.modalCoordReview.style.display = 'none';
    state.currentReviewItem = null;
}

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
        showToast('请输入有效的经纬度', 'warning');
        return;
    }

    // 简单范围校验（中国大陆范围）
    if (longitude < 73 || longitude > 136 || latitude < 3 || latitude > 54) {
        showToast('经纬度超出合理范围，请检查（中国范围：经度73-136，纬度3-54）', 'warning', 4000);
        // 不阻止提交，仅提示
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
            showToast('坐标已确认', 'success', 1500);
        }
    });
}

// ==================== 地图预览 ====================

function previewOnMap() {
    const longitude = parseFloat(elements.confirmedLongitude.value);
    const latitude = parseFloat(elements.confirmedLatitude.value);
    
    if (isNaN(longitude) || isNaN(latitude)) {
        showToast('请先输入有效的经纬度', 'warning');
        return;
    }
    
    // 显示地图面板
    elements.panelMap.style.display = 'block';
    
    // 加载高德地图
    loadAmapScript().then(() => {
        initMap(longitude, latitude);
    }).catch(error => {
        showToast('地图加载失败: ' + error.message, 'error');
    });
}

function viewOnMap(longitude, latitude) {
    elements.panelMap.style.display = 'block';
    
    loadAmapScript().then(() => {
        initMap(longitude, latitude);
    }).catch(error => {
        showToast('地图加载失败: ' + error.message, 'error');
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


// 筛选结果
function filterResults(filter) {
    state.resultFilter = filter;
    
    // 更新按钮状态
    elements.filterBtns.forEach(btn => {
        btn.classList.toggle('active', btn.dataset.filter === filter);
    });
    
    // 重新渲染结果
    loadResults();
}

function renderResults(items) {
    const tbody = elements.resultTbody;
    tbody.innerHTML = '';
    
    let successCount = 0;
    let failCount = 0;
    let pendingCount = 0;

    // 先计算全量统计（不受筛选影响）
    items.forEach(item => {
        if (item.search_failed || item.needs_manual) {
            failCount++;
        } else if (item.status === '已完成' || item.longitude) {
            successCount++;
        } else if (item.status === '出错' || item.status === '错误') {
            failCount++;
        } else {
            pendingCount++;
        }
    });

    // 始终更新统计数字
    elements.resultTotal.textContent = items.length;
    elements.resultSuccess.textContent = successCount;
    elements.resultFail.textContent = failCount;
    const pendingEl = document.getElementById('result-pending');
    if (pendingEl) pendingEl.textContent = pendingCount;
    
    // 根据筛选条件过滤
    let filteredItems = items;
    if (state.resultFilter === 'success') {
        filteredItems = items.filter(item => item.status === '已完成' || item.longitude);
    } else if (state.resultFilter === 'fail') {
        filteredItems = items.filter(item => item.status === '出错' || item.status === '错误' || item.search_failed || item.needs_manual);
    } else if (state.resultFilter === 'pending') {
        filteredItems = items.filter(item => 
            item.status !== '已完成' && 
            item.status !== '出错' && 
            item.status !== '错误' &&
            !item.search_failed &&
            !item.needs_manual &&
            !item.longitude
        );
    }
    
    filteredItems.forEach((item, index) => {
        const tr = document.createElement('tr');
        
        // 判断状态
        let statusText = '处理中';
        let statusClass = 'pending';
        let statusIcon = '⏳';
        
        // 检查是否是搜索失败或需人工处理
        if (item.search_failed || item.needs_manual) {
            statusText = '需人工处理';
            statusClass = 'error';
            statusIcon = '❌';
            failCount++;
        } else if (item.status === '已完成' || item.longitude) {
            // 已完成的，区分自动确认和人工确认
            if (item.auto_approved) {
                statusText = '自动确认';
                statusClass = 'success';
                statusIcon = '✅';
            } else {
                statusText = '成功';
                statusClass = 'success';
                statusIcon = '✅';
            }
            successCount++;
        } else if (item.status === '出错' || item.status === '错误') {
            statusText = '失败';
            statusClass = 'error';
            statusIcon = '❌';
            failCount++;
        } else if (item.status === '等待地址审核' || item.status === '等待坐标审核') {
            // 待审核的，区分置信度
            if (item.address_quality === 'medium') {
                statusText = '建议复核';
                statusClass = 'warning';
                statusIcon = '⚠️';
            } else {
                statusText = '待审核';
                statusClass = 'warning';
                statusIcon = '⚠️';
            }
            pendingCount++;
        } else {
            pendingCount++;
        }
        
        // 地址质量标签
        let qualityBadge = '';
        if (item.address_confidence && item.address_confidence > 0) {
            const confidencePercent = (item.address_confidence * 100).toFixed(0);
            let qualityClass = '';
            let qualityText = '';
            
            if (item.address_quality === 'high' || item.address_confidence >= 0.7) {
                qualityClass = 'quality-high';
                qualityText = '高置信度';
            } else if (item.address_quality === 'medium' || item.address_confidence >= 0.5) {
                qualityClass = 'quality-medium';
                qualityText = '中置信度';
            } else if (item.address_confidence > 0) {
                qualityClass = 'quality-low';
                qualityText = '低置信度';
            }
            
            if (qualityText) {
                qualityBadge = `<span class="quality-badge ${qualityClass}">${qualityText} ${confidencePercent}%</span>`;
            }
        }
        
        // 操作按钮
        let actionHtml = '';
        if (item.longitude) {
            actionHtml += `<button class="btn btn-small btn-secondary" onclick="viewOnMap(${item.longitude}, ${item.latitude})">🗺️ 地图</button>`;
        }
        if (item.status !== '已完成' && item.status !== '处理中') {
            actionHtml += `<button class="btn btn-small btn-warning" onclick="retryItem(${item.id})">🔄 重试</button>`;
        }
        if (!actionHtml) {
            actionHtml = '-';
        }
        
        tr.innerHTML = `
            <td>${item.id + 1}</td>
            <td>${item.name || '-'}</td>
            <td>
                ${item.confirmed_address || '-'}
                ${qualityBadge ? '<br>' + qualityBadge : ''}
            </td>
            <td>${item.longitude ? item.longitude.toFixed(6) : '-'}</td>
            <td>${item.latitude ? item.latitude.toFixed(6) : '-'}</td>
            <td><span class="status-badge ${statusClass}">${statusIcon} ${statusText}</span></td>
            <td class="action-cell">${actionHtml}</td>
        `;
        
        tbody.appendChild(tr);
    });
}

function viewOnMap(longitude, latitude) {
    elements.panelMap.style.display = 'block';
    
    loadAmapScript().then(() => {
        initMap(longitude, latitude);
    }).catch(error => {
        showToast('地图加载失败: ' + error.message, 'error');
    });
}

function retryItem(itemId) {
    showConfirm('确定要重试这条数据吗？', '重试', '取消').then(ok => {
        if (!ok) return;

        fetch('/api/retry-item', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                task_id: state.taskId,
                item_id: itemId,
                type: 'search'
            })
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                addLog(`已发送第 ${itemId + 1} 条重试指令`, 'info');
                setTimeout(loadResults, 1000);
            }
        });
    });
}



// ==================== 导出 ====================

function exportExcel() {
    window.location.href = `/api/export/${state.taskId}`;
}

// ==================== 重置任务 ====================

function resetTask() {
    showConfirm('确定要新建任务吗？当前任务数据将会丢失。', '新建任务', '取消').then(ok => {
        if (!ok) return;

        stopStatusPolling();
        state.taskId = null;
        state.columns = [];
        state.currentStep = 1;
        state.currentReviewItem = null;
        
        elements.fileInfo.style.display = 'none';
        elements.taskIdBadge.style.display = 'none';
        elements.logContainer.innerHTML = '<div class="log-empty">等待任务开始...</div>';
        elements.amapKey.value = '';
        
        resetUploadArea();
        goToStep(1);
    });
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
window.interruptTask = interruptTask;
window.showReturnConfigModal = showReturnConfigModal;
window.retryItem = retryItem;
window.retrySearch = retrySearch;
window.customSearch = customSearch;


// ==================== AI搜索（豆包） ====================
// 页面加载时恢复保存的豆包Key
document.addEventListener('DOMContentLoaded', function() {
    const savedDoubaoKey = localStorage.getItem('doubao_key_saved');
    if (savedDoubaoKey) {
        const input = document.getElementById('doubao-key');
        if (input) input.value = savedDoubaoKey;
    }
});

// 调用豆包AI搜索
async function searchWithAI(projectName) {
    const doubaoKey = document.getElementById('doubao-key').value.trim();
    const doubaoModelId = document.getElementById('doubao-model-id') ? document.getElementById('doubao-model-id').value.trim() : '';
    if (!doubaoKey) {
        showToast('请先在配置页面填写豆包API Key', 'warning');
        return null;
    }

    const amapKey = elements.amapKey.value.trim();

    showSearchLoading('🤖 AI搜索中...');
    try {
        const resp = await fetch('/api/search_with_ai', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({
                project_name: projectName,
                doubao_api_key: doubaoKey,
                amap_key: amapKey || undefined,
                doubao_model_id: doubaoModelId || undefined
            })
        });
        const data = await resp.json();
        hideSearchLoading();

        if (data.success) {
            const aiResult = {
                title: '🤖 豆包AI回答',
                snippet: data.raw_response || data.address,
                address: data.address,
                lng: data.lng,
                lat: data.lat,
                confidence: data.confidence,
                source: data.source,
                is_ai: true
            };
            displayAISearchResult(aiResult);
            // 自动填入地址 + 经纬度
            const addrEl = document.getElementById('confirmed-address');
            if (addrEl) addrEl.value = aiResult.address;
            const lngEl = document.getElementById('confirmed-longitude');
            const latEl = document.getElementById('confirmed-latitude');
            if (lngEl && aiResult.lng) lngEl.value = aiResult.lng;
            if (latEl && aiResult.lat) latEl.value = aiResult.lat;
            showToast('AI搜索完成！已自动填入地址和坐标', 'success');
            return aiResult;
        } else {
            showToast('AI搜索失败：' + (data.message || '未知错误'), 'error');
            return null;
        }
    } catch (err) {
        hideSearchLoading();
        showToast('AI搜索请求失败：' + err.message, 'error');
        return null;
    }
}

// 显示AI搜索结果
function displayAISearchResult(aiResult) {
    const container = elements.searchResults;
    container.innerHTML = '';

    const item = document.createElement('div');
    item.className = 'search-result-item ai-result';
    item.style.borderLeft = '4px solid #667eea';
    item.innerHTML = `
        <div class="result-title" style="color:#667eea;font-weight:600;">${aiResult.title}</div>
        <div class="result-snippet">${aiResult.snippet}</div>
        <div class="result-address" style="margin-top:6px;color:#059669;font-weight:500;">
            📍 推荐地址：<span class="ai-address-text">${aiResult.address}</span>
        </div>
        ${aiResult.lng ? `<div class="result-coord" style="font-size:12px;color:#6b7280;">坐标：${aiResult.lng}, ${aiResult.lat}</div>` : ''}
        <div class="result-confidence" style="font-size:12px;color:#6b7280;">置信度：${(aiResult.confidence * 100).toFixed(0)}% | 来源：${aiResult.source}</div>
    `;
    item.addEventListener('click', () => {
        selectAISearchResult(aiResult);
    });
    container.appendChild(item);
}

// 选择AI搜索结果
function selectAISearchResult(aiResult) {
    // 高亮选中项
    document.querySelectorAll('.search-result-item').forEach(el => el.classList.remove('selected'));
    const aiItem = document.querySelector('.ai-result');
    if (aiItem) aiItem.classList.add('selected');

    // 填入确认地址 + 经纬度输入框
    const addrEl = document.getElementById('confirmed-address');
    if (addrEl) addrEl.value = aiResult.address;
    const lngEl = document.getElementById('confirmed-longitude');
    const latEl = document.getElementById('confirmed-latitude');
    if (lngEl && aiResult.lng) lngEl.value = aiResult.lng;
    if (latEl && aiResult.lat) latEl.value = aiResult.lat;

    showToast('已选择AI搜索结果，请确认地址后点击确认', 'info');
}

// 绑定AI搜索按钮事件
function bindAISearchButton() {
    const btn = document.getElementById('btn-ai-search');
    if (btn && !btn._aiBound) {
        btn._aiBound = true;
        btn.addEventListener('click', async () => {
            if (!state.currentReviewItem) {
                showToast('当前没有正在审核的项目', 'warning');
                return;
            }
            const projectName = state.currentReviewItem.name || state.currentReviewItem.project_name;
            if (!projectName) {
                showToast('无法获取项目名称', 'error');
                return;
            }
            // 基本的豆包 Key / model_id 校验
            const doubaoKey = document.getElementById('doubao-key');
            if (!doubaoKey || !doubaoKey.value.trim()) {
                showToast('请先在配置页面填写豆包API Key', 'warning');
                return;
            }
            const doubaoModelId = document.getElementById('doubao-model-id');
            const modelIdVal = doubaoModelId ? doubaoModelId.value.trim() : '';
            if (modelIdVal && !/^(ep|pp|ark|doubao)/i.test(modelIdVal)) {
                // 只在有内容且明显不合规时提示一下，不强制 block
                showToast('提示：接入点 ID 看起来不合规（通常以 ep- 开头）', 'warning');
            }
            await searchWithAI(projectName);
        });
    }
}

// 在 initEventListeners 末尾调用
const _origInit = initEventListeners;
