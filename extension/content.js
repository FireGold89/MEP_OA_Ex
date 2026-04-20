/**
 * Version: 4.12.4 - Browser Field UI Fix
 * 採購立項自動化工具 - 核心邏輯
 */

(function () {
    const isTop = (window.self === window.top);
    if (!isTop) return;

    let panel, ball;
    let projects = [];
    const fields = ['label', 'propertyName', 'reportNo', 'manager', 'projectContent', 'budget', 'total', 'quoteType', 'buyType', 'currency', 'inviteCount', 'replyCount', 'reason', 'winnerName', 'contractCurrency', 'contractAmount'];
    
    // 映射表：將顯示文字/舊索引轉為 OA 系統真實的 Value (根據 F12 截圖校準)
    const quoteMap = { "0": "0", "1": "1", "2": "2", "3": "3", "報價邀請": "0", "招標": "1", "特殊情況 (緊急)": "2", "續約": "3" };
    const buyMap = { 
        "0": "9", "1": "10", "2": "11", "3": "12", "4": "14", "5": "4", "6": "13",
        "保養維修材料": "9", "保養合約分判": "10", "工程材料": "11", 
        "分判工程": "12", "後加工程": "14", "固定資產": "4", "其他": "13" 
    };

    function init() {
        if (document.getElementById('oa-side-panel')) return;

        // 注入橋接腳本至頁面環境
        const script = document.createElement('script');
        script.src = chrome.runtime.getURL('main_bridge.js');
        (document.head || document.documentElement).appendChild(script);

        injectPanelHTML();
        injectBall();
        updateVisibility();
        syncData();
        startPickerWatcher();
        listenMessages();

        // 🌟 核心同步功能：監聽儲存庫變動
        chrome.storage.onChanged.addListener((changes, area) => {
            if (area === 'local' && changes.oa_projects_v13) {
                console.log("OA Content: Storage change detected! Syncing UI...");
                projects = changes.oa_projects_v13.newValue || [];
                renderList();
            }
            if (area === 'local' && changes.oa_show_float_ball !== undefined) {
                updateVisibility();
            }
        });
    }

    async function updateVisibility() {
        chrome.storage.local.get(['oa_show_float_ball'], (res) => {
            const visible = (res.oa_show_float_ball !== false);
            const display = visible ? 'flex' : 'none';
            if (ball) ball.style.display = display;
            if (panel && !panel.classList.contains('open')) panel.style.display = display;
        });
    }

    function listenMessages() {
        chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
            if (msg.type === "OA_LINK_FILL" || msg.type === "OA_LINK_CLEAR") {
                console.log("OA Content: Proxying Message to Bridge ->", msg.type);
                if (msg.project) msg.project = normalizeProjectData(msg.project);
                window.postMessage(msg, "*"); 
                sendResponse({ success: true });
            } else if (msg.type === "TOGGLE_FLOAT_BALL") {
                updateVisibility();
                // 同步更新設置頁面的開關狀態
                const checkBall = document.getElementById('oa-check-ball-side');
                if (checkBall) checkBall.checked = msg.visible;
                sendResponse({ success: true });
            } else if (msg.type === "UPDATE_SETTING") {
                console.log("OA Content: Received Setting Update ->", msg.key, msg.value);
                // 實時更新 UI
                if (msg.key === "title_size") {
                    const slider = document.getElementById('side-title-size');
                    const valLab = document.getElementById('side-title-size-val');
                    if (slider) slider.value = msg.value;
                    if (valLab) valLab.textContent = msg.value + 'px';
                    renderList();
                } else if (msg.key === "subtitle_mode") {
                    const ctrl = document.getElementById('side-subtitle-mode');
                    if (ctrl) {
                        ctrl.querySelectorAll('.oa-seg-btn').forEach(b => b.classList.toggle('active', b.dataset.value === msg.value));
                    }
                    renderList();
                } else if (msg.key === "compact") {
                    const cb = document.getElementById('side-compact-mode');
                    if (cb) cb.checked = msg.value;
                    renderList();
                }
                sendResponse({ success: true });
            }
            return true;
        });
    }

    function normalizeProjectData(p) {
        const np = JSON.parse(JSON.stringify(p));
        // 確保發送給 Bridge 的是數值 ID
        if (np.quoteType && isNaN(np.quoteType)) np.quoteType = quoteMap[np.quoteType] || np.quoteType;
        if (np.buyType && isNaN(np.buyType)) np.buyType = buyMap[np.buyType] || np.buyType;
        return np;
    }

    function injectPanelHTML() {
        panel = document.createElement('div');
        panel.id = 'oa-side-panel';
        panel.innerHTML = `
            <div class="ym-header">
                <div class="ym-header-left" style="display:flex; align-items:center; gap:10px;">
                    <div class="ym-logo"><div class="oa-icon-o"><span class="oa-icon-a">A</span></div></div>
                    <div>
                        <div class="ym-header-title">採購一鍵通</div>
                        <div class="ym-header-subtitle">OA 表單自動化工具</div>
                    </div>
                </div>
                <div class="ym-header-right">
                    <span id="oa-btn-refresh" title="同步數據">↻</span>
                    <span id="oa-panel-close" title="隱藏視窗">✕</span>
                </div>
            </div>

            <div class="oa-action-toolbar">
                <button class="oa-tool-btn outline" id="oa-btn-add">➕ 新增</button>
                <button class="oa-tool-btn" id="oa-btn-import-top">↑ 匯入</button>
                <button class="oa-tool-btn" id="oa-btn-export-top">↓ 匯出</button>
                <button class="oa-tool-btn primary" id="oa-btn-save-top">✔ 保存</button>
            </div>

            <div class="oa-p-tabs">
                <div class="oa-p-tab active" id="tab-v-fill">快速填充</div>
                <div class="oa-p-tab" id="tab-v-manage">項目管理</div>
                <div class="oa-p-tab" id="tab-v-settings">更多設置</div>
            </div>

            <div id="view-v-fill" class="oa-panel-body">
                <div style="display:flex; gap:8px; margin-bottom:12px; align-items:center;">
                    <input type="text" id="oa-search-input" class="oa-search-input" placeholder="🔍 搜尋項目或經理...">
                </div>
                <div class="oa-fill-toolbar">
                    <button id="oa-btn-undo">↩ 恢復</button>
                    <button id="oa-btn-clear">🧹 清空</button>
                </div>
                <div id="oa-disp-list"></div>
            </div>

            <div id="view-v-manage" class="oa-panel-body hidden">
                <div style="font-weight:700; font-size:16px; margin-bottom:20px; color:#1d1d1f;" id="oa-p-title">新增項目</div>
                <input type="hidden" id="in-f-id">
                <input type="hidden" id="in-f-managerId">

                <div class="oa-p-group"><label>項目標題 (顯示用)</label><input type="text" id="in-f-label" placeholder="例如：信成 - 電器材料"></div>

                <div class="oa-grid-row">
                    <div class="oa-p-group"><label>物業名稱</label><input type="text" id="in-f-propertyName"></div>
                    <div class="oa-p-group"><label>項目經理</label>
                        <div class="oa-picker-container">
                            <input type="text" id="in-f-manager" placeholder="輸入或選取">
                            <button id="btn-f-pick" class="oa-picker-btn">🔍</button>
                        </div>
                    </div>
                </div>

                <div class="oa-p-group"><label>美博報價編號</label><input type="text" id="in-f-reportNo"></div>
                <div class="oa-p-group"><label>項目內容</label><textarea id="in-f-projectContent" rows="3"></textarea></div>

                <div class="oa-grid-row">
                    <div class="oa-p-group">
                        <label>合約總價</label><input type="text" id="in-f-total"></div>
                    <div class="oa-p-group">
                        <label>報價形式</label>
                        <select id="in-f-quoteType">
                            <option value="">請選擇</option>
                            <option value="0">報價邀請</option>
                            <option value="1">招標</option>
                            <option value="2">特殊情況 (緊急)</option>
                            <option value="3">續約</option>
                        </select>
                    </div>
                </div>

                <div class="oa-grid-row">
                    <div class="oa-p-group">
                        <label>採購類別</label>
                        <select id="in-f-buyType">
                            <option value="">請選擇</option>
                            <option value="9">保養維修材料</option>
                            <option value="10">保養合約分判</option>
                            <option value="11">工程材料</option>
                            <option value="12">分判工程</option>
                            <option value="14">後加工程</option>
                            <option value="4">固定資產</option>
                            <option value="13">其他</option>
                        </select>
                    </div>
                    <div class="oa-p-group"><label>立項預算金額</label><input type="text" id="in-f-budget"></div>
                </div>

                <div class="oa-p-group"><label>幣種</label>
                    <select id="in-f-currency">
                        <option value="0">HKD</option>
                        <option value="1">RMB</option>
                        <option value="2">USD</option>
                        <option value="3">MOP</option>
                    </select>
                </div>

                <div style="font-weight:700; font-size:13px; margin:20px 0 12px; color:var(--ym-primary);">📦 報價明細</div>
                <div id="oa-detail-list"></div>
                <button id="btn-add-detail" style="width:100%; padding:12px; border:1px dashed var(--ym-primary); color:var(--ym-primary); background:none; border-radius:14px; cursor:pointer; margin-bottom:20px; font-weight:700; font-size:13px;">+ 添加明細</button>

                <div class="oa-grid-row">
                    <div class="oa-p-group"><label>邀請公司 (間)</label><input type="text" id="in-f-inviteCount"></div>
                    <div class="oa-p-group"><label>有效報價/回復 (份)</label><input type="text" id="in-f-replyCount"></div>
                </div>

                <div class="oa-p-group"><label>理由</label><textarea id="in-f-reason" rows="2"></textarea></div>
                <div class="oa-p-group"><label>中標公司</label><input type="text" id="in-f-winnerName"></div>

                <div class="oa-grid-row">
                    <div class="oa-p-group">
                        <label>合約幣種</label>
                        <select id="in-f-contractCurrency">
                            <option value="0">HKD</option>
                            <option value="1">RMB</option>
                            <option value="2">USD</option>
                            <option value="3">MOP</option>
                        </select>
                    </div>
                    <div class="oa-p-group"><label>合約金額</label><input type="text" id="in-f-contractAmount"></div>
                </div>

                <button id="btn-f-save-main" class="oa-p-btn-action">儲存項目</button>
                <button id="btn-f-save-as-new" class="oa-p-btn-action hidden" style="background:#fff; border:1px solid #1d1d1f; color:#1d1d1f; margin-top:12px;">另存為新項目</button>
                <div id="btn-f-cancel" style="text-align:center; margin-top:15px; font-size:13px; color:#999; cursor:pointer;" class="hidden">取消編輯</div>

                <div style="margin-top:25px; border-top:1px solid var(--ym-border); padding-top:20px; display:flex; gap:12px;">
                    <button id="btn-f-export" style="flex:1; font-size:12px; padding:10px; border-radius:12px; border:1px solid #ddd; background:#fff; cursor:pointer; font-weight:600;">匯出數據</button>
                    <button id="btn-f-import" style="flex:1; font-size:12px; padding:10px; border-radius:12px; border:1px solid #ddd; background:#fff; cursor:pointer; font-weight:600;">匯入數據</button>
                    <input type="file" id="in-f-file-side" class="hidden" accept=".csv">
                </div>
            </div>

            <div id="view-v-settings" class="oa-panel-body hidden">
                <div class="s-section">🎨 顯示偏好</div>

                <div class="s-row">
                    <div><div class="s-label">項目標題字體大小</div><div class="s-desc">調整列表中標題的大小</div></div>
                    <div style="display:flex; align-items:center; gap:8px;">
                        <input type="range" class="oa-slider" id="side-title-size" min="12" max="18" value="15">
                        <span class="oa-slider-val" id="side-title-size-val">15px</span>
                    </div>
                </div>

                <div class="s-row">
                    <div><div class="s-label">副標題資訊</div><div class="s-desc">報價號旁的顯示資訊</div></div>
                    <div class="oa-seg-ctrl" id="side-subtitle-mode">
                        <button class="oa-seg-btn active" data-value="both">全部</button>
                        <button class="oa-seg-btn" data-value="reportNo">報價號</button>
                        <button class="oa-seg-btn" data-value="budget">預算</button>
                    </div>
                </div>

                <div class="s-row">
                    <div><div class="s-label">緊湊模式</div><div class="s-desc">縮小卡片間距</div></div>
                    <label class="oa-toggle">
                        <input type="checkbox" id="side-compact-mode">
                        <div class="oa-toggle-track"></div>
                        <div class="oa-toggle-thumb"></div>
                    </label>
                </div>

                <div class="s-section">⚙️ 功能設置</div>

                <div class="s-row">
                    <div><div class="s-label">顯示懸浮輔助球</div><div class="s-desc">頁面右下角快速工具球</div></div>
                    <label class="oa-toggle">
                        <input type="checkbox" id="oa-check-ball-side" checked>
                        <div class="oa-toggle-track"></div>
                        <div class="oa-toggle-thumb"></div>
                    </label>
                </div>

                <div class="s-row">
                    <div><div class="s-label">匯出確認提示</div><div class="s-desc">匯出前詢問是否覆蓋舊檔案</div></div>
                    <label class="oa-toggle">
                        <input type="checkbox" id="side-export-confirm" checked>
                        <div class="oa-toggle-track"></div>
                        <div class="oa-toggle-thumb"></div>
                    </label>
                </div>
            </div>

            <div style="padding:8px; text-align:center; font-size:10px; color:#c0c0cc; background:#fff; border-top:1px solid var(--ym-border); letter-spacing:0.5px; flex-shrink:0;">v4.12.4</div>
        `;
        document.body.appendChild(panel);
        bindEvents();
    }

    function injectBall() {
        if (document.getElementById('oa-float-ball')) return;
        ball = document.createElement('div');
        ball.id = 'oa-float-ball';
        ball.innerHTML = `<div class="oa-ball-main"><div class="oa-icon-o"><span class="oa-icon-a">A</span></div></div>`;
        document.body.appendChild(ball);

        const main = ball.querySelector('.oa-ball-main');
        let isDragging = false;
        let startX = 0, startY = 0, initialX = 0, initialY = 0;

        main.addEventListener('pointerdown', (e) => {
            // 只響應主鍵（左鍵）
            if (e.button !== 0) return;
            e.preventDefault();

            isDragging = false;
            startX = e.clientX;
            startY = e.clientY;
            initialX = ball.offsetLeft;
            initialY = ball.offsetTop;
            ball.style.transition = 'none';
            main.style.cursor = 'grabbing';

            // 關鍵：Pointer Capture 確保所有事件在鼠標移出窗口後仍由此元素接收
            main.setPointerCapture(e.pointerId);
        });

        main.addEventListener('pointermove', (e) => {
            // 沒有主指鉢錢則跳過
            if (!main.hasPointerCapture(e.pointerId)) return;
            const dx = e.clientX - startX;
            const dy = e.clientY - startY;
            if (Math.abs(dx) > 5 || Math.abs(dy) > 5) {
                isDragging = true;
                let nextX = initialX + dx;
                let nextY = initialY + dy;
                const maxW = window.innerWidth - ball.offsetWidth;
                const maxH = window.innerHeight - ball.offsetHeight;
                if (nextX < 0) nextX = 0; if (nextX > maxW) nextX = maxW;
                if (nextY < 0) nextY = 0; if (nextY > maxH) nextY = maxH;
                ball.style.left = nextX + 'px';
                ball.style.top = nextY + 'px';
                ball.style.right = 'auto';
                ball.style.bottom = 'auto';
            }
        });

        // pointerup / pointercancel 都触發此清理函數
        const onPointerEnd = (e) => {
            if (!main.hasPointerCapture(e.pointerId)) return;
            main.releasePointerCapture(e.pointerId);
            main.style.cursor = 'grab';

            if (!isDragging) {
                panel.classList.toggle('open');
            } else {
                const screenW = window.innerWidth;
                ball.style.transition = 'all 0.5s cubic-bezier(0.19, 1, 0.22, 1)';
                const finalX = screenW - ball.offsetWidth - 24;
                ball.style.left = finalX + 'px';
                main.classList.add('ball-snapping');
                setTimeout(() => {
                    main.classList.remove('ball-snapping');
                    ball.style.transition = 'none';
                    ball.style.left = 'auto';
                    ball.style.right = '24px';
                    const currentTop = parseInt(ball.style.top) || 0;
                    const maxTop = window.innerHeight - ball.offsetHeight - 24;
                    if (currentTop < 24) ball.style.top = '24px';
                    else if (currentTop > maxTop) ball.style.top = maxTop + 'px';
                }, 500);
            }
            isDragging = false;
        };

        main.addEventListener('pointerup', onPointerEnd);
        main.addEventListener('pointercancel', onPointerEnd);
    }

    function addDetailRow(data = {}) {
        const container = document.getElementById('oa-detail-list');
        const div = document.createElement('div');
        div.className = 'oa-detail-row';
        div.innerHTML = `
            <div style="position:absolute; right:10px; top:10px; cursor:pointer; color:#ccc; font-size:18px; font-weight:700;" class="dt-remove">✕</div>
            <div class="oa-p-group"><label>承判商</label><input type="text" class="dt-vendor" value="${data.vendorName || ''}"></div>
            <div class="oa-p-group"><label>報價內容</label><input type="text" class="dt-content" value="${data.content || ''}"></div>
            <div class="oa-grid-row">
                <div class="oa-p-group">
                    <label>幣種</label>
                    <select class="dt-currency">
                        <option value="0" ${data.detailCurrency === '0' || !data.detailCurrency ? 'selected' : ''}>HKD</option>
                        <option value="1" ${data.detailCurrency === '1' ? 'selected' : ''}>RMB</option>
                        <option value="2" ${data.detailCurrency === '2' ? 'selected' : ''}>USD</option>
                        <option value="3" ${data.detailCurrency === '3' ? 'selected' : ''}>MOP</option>
                    </select>
                </div>
                <div class="oa-p-group"><label>金額</label><input type="text" class="dt-amount" value="${data.amount || ''}"></div>
            </div>
        `;
        container.appendChild(div);

        // 自動更新邀請公司數量
        const updateInviteCount = () => {
            const count = container.querySelectorAll('.oa-detail-row').length;
            const inviteInput = document.getElementById('in-f-inviteCount');
            if (inviteInput) inviteInput.value = count;
        };

        div.querySelector('.dt-remove').onclick = () => {
            div.remove();
            updateInviteCount();
        };

        updateInviteCount();

        // 🌟 新增：第一行同步邏輯 (立項預算/合約金額/中標公司)
        const rows = container.querySelectorAll('.oa-detail-row');
        if (rows.length === 1) {
            const vInput = div.querySelector('.dt-vendor');
            const aInput = div.querySelector('.dt-amount');

            const syncFirstRow = () => {
                const firstRow = container.querySelector('.oa-detail-row');
                if (!firstRow) return;
                
                const vendor = firstRow.querySelector('.dt-vendor').value.trim();
                const amount = firstRow.querySelector('.dt-amount').value.trim();
                
                const winnerInput = document.getElementById('in-f-winnerName');
                const budgetInput = document.getElementById('in-f-budget');
                const contractAmountInput = document.getElementById('in-f-contractAmount');

                if (winnerInput) winnerInput.value = vendor;
                if (budgetInput) budgetInput.value = amount;
                if (contractAmountInput) contractAmountInput.value = amount;
            };

            vInput.addEventListener('input', syncFirstRow);
            aInput.addEventListener('input', syncFirstRow);
            
            // 初始同步一次 (如果是從編輯加載)
            syncFirstRow();
        }
    }

    function resetForm() {
        document.getElementById('in-f-id').value = "";
        document.getElementById('in-f-managerId').value = "";
        fields.forEach(f => {
            const el = document.getElementById('in-f-'+f);
            if(el) {
                if (el.tagName === 'SELECT') {
                    // 下拉選單默認值處理
                    if (f === 'currency' || f==='contractCurrency') el.value = "0";
                    else el.value = "";
                } else {
                    el.value = "";
                }
            }
        });
        document.getElementById('oa-detail-list').innerHTML = "";
        addDetailRow();
        document.getElementById('oa-p-title').innerText = '新增項目';
        document.getElementById('btn-f-save-main').innerText = '儲存項目';
        document.getElementById('btn-f-cancel').classList.add('hidden');
        document.getElementById('btn-f-save-as-new').classList.add('hidden');
    }

    async function syncData() {
        const res = await chrome.storage.local.get(['oa_projects_v13']);
        projects = res.oa_projects_v13 || [];
        renderList();
    }

    function buildCsvDataUrl() {
        if(projects.length === 0) { alert("尚無數據可匯出"); return null; }
        const h = ["項目標題", "物業名稱", "報價編號", "項目經理", "項目內容", "立項預算", "合約總價", "報價形式", "採購類別", "幣種", "承判商", "報價內容", "明細幣種", "金額", "邀請公司", "有效報價", "推荐理由", "中標公司", "合約幣種", "合約金額"];
        const rows = [];
        projects.forEach(p => {
            (p.details || [{}]).forEach(dt => {
                const row = [p.label, p.propertyName, p.reportNo, p.manager, p.projectContent, p.budget, p.total, p.quoteType, p.buyType, p.currency, dt.vendorName, dt.content, dt.detailCurrency, dt.amount, p.inviteCount, p.replyCount, p.reason, p.winnerName, p.contractCurrency, p.contractAmount];
                rows.push(row.map(v => `"${(String(v || '')).replace(/"/g, '""')}"`).join(","));
            });
        });
        const csvContent = "\uFEFF" + h.join(",") + "\n" + rows.join("\n");
        return 'data:text/csv;charset=utf-8,' + encodeURIComponent(csvContent);
    }

    function triggerDownload(filename, conflictAction) {
        const dataUrl = buildCsvDataUrl();
        if (!dataUrl) return;
        chrome.runtime.sendMessage({ type: 'DOWNLOAD_CSV', dataUrl, filename, conflictAction }, (res) => {
            if (chrome.runtime.lastError) console.warn('OA Export error:', chrome.runtime.lastError.message);
        });
    }

    async function exportToCSV() {
        if(projects.length === 0) return alert("尚無數據可匯出");

        const setting = await chrome.storage.local.get(['oa_setting_export_confirm']);
        const needsConfirm = (setting.oa_setting_export_confirm !== false);

        if (!needsConfirm) {
            triggerDownload('OA_Projects.csv', 'overwrite');
            return;
        }

        // 如果對話框已存在則移除
        const existingModal = document.getElementById('oa-export-modal');
        if (existingModal) existingModal.remove();

        const modal = document.createElement('div');
        modal.id = 'oa-export-modal';
        modal.style.cssText = `
            position:fixed; inset:0; z-index:2147483647;
            display:flex; align-items:center; justify-content:center;
            background:rgba(0,0,0,0.4); backdrop-filter:blur(4px);
            animation:oa-fadein 0.15s ease;
        `;
        modal.innerHTML = `
            <div style="background:#fff; border-radius:20px; padding:28px 32px; width:320px;
                        box-shadow:0 20px 60px rgba(0,0,0,0.25); font-family:-apple-system,sans-serif;">
                <div style="font-size:22px; text-align:center; margin-bottom:8px">📤</div>
                <div style="font-weight:700; font-size:16px; color:#1d1d1f; text-align:center; margin-bottom:8px">匯出確認</div>
                <div style="font-size:13px; color:#666; text-align:center; margin-bottom:24px; line-height:1.6">
                    下載資料夾中可能已存在<br>
                    <strong style="color:#1d1d1f">OA_Projects.csv</strong><br>
                    請選擇處理方式：
                </div>
                <button id="oa-exp-overwrite" style="width:100%; padding:14px; margin-bottom:10px;
                    border-radius:12px; border:none; background:#1d1d1f; color:#fff;
                    font-size:14px; font-weight:600; cursor:pointer;">
                    ✅ 覆蓋舊檔（OA_Projects.csv）
                </button>
                <button id="oa-exp-new" style="width:100%; padding:14px; margin-bottom:10px;
                    border-radius:12px; border:1px solid #ddd; background:#f9f9f9; color:#1d1d1f;
                    font-size:14px; font-weight:600; cursor:pointer;">
                    📅 另存新檔（加日期時間）
                </button>
                <button id="oa-exp-cancel" style="width:100%; padding:10px;
                    border-radius:12px; border:none; background:none; color:#999;
                    font-size:13px; cursor:pointer;">取消</button>
            </div>
        `;
        document.body.appendChild(modal);

        const close = () => modal.remove();
        modal.addEventListener('click', (e) => { if (e.target === modal) close(); });
        document.getElementById('oa-exp-cancel').onclick = close;

        document.getElementById('oa-exp-overwrite').onclick = () => {
            close();
            triggerDownload('OA_Projects.csv', 'overwrite');
        };
        document.getElementById('oa-exp-new').onclick = () => {
            close();
            const ts = new Date().toISOString().slice(0,16).replace('T','_').replace(':','-');
            triggerDownload(`OA_Projects_${ts}.csv`, 'uniquify');
        };
    }

    function bindEvents() {
        document.getElementById('oa-btn-refresh').onclick = syncData;
        document.getElementById('oa-panel-close').onclick = () => panel.classList.remove('open');
        
        document.getElementById('oa-btn-add').onclick = () => { resetForm(); document.getElementById('tab-v-manage').click(); };
        document.getElementById('oa-btn-save-top').onclick = () => document.getElementById('btn-f-save-main').click();
        document.getElementById('oa-btn-export-top').onclick = exportToCSV;
        const fileIn = document.getElementById('in-f-file-side');
        document.getElementById('oa-btn-import-top').onclick = () => fileIn.click();

        const tFill = document.getElementById('tab-v-fill'), tManage = document.getElementById('tab-v-manage'), tSettings = document.getElementById('tab-v-settings');
        const vFill = document.getElementById('view-v-fill'), vManage = document.getElementById('view-v-manage'), vSettings = document.getElementById('view-v-settings');

        tFill.onclick = () => { tFill.classList.add('active'); tManage.classList.remove('active'); tSettings.classList.remove('active'); vFill.classList.remove('hidden'); vManage.classList.add('hidden'); vSettings.classList.add('hidden'); };
        tManage.onclick = () => { tManage.classList.add('active'); tFill.classList.remove('active'); tSettings.classList.remove('active'); vManage.classList.remove('hidden'); vFill.classList.add('hidden'); vSettings.classList.add('hidden'); };
        tSettings.onclick = () => { tSettings.classList.add('active'); tFill.classList.remove('active'); tManage.classList.remove('active'); vSettings.classList.remove('hidden'); vFill.classList.add('hidden'); vManage.classList.add('hidden'); };

        // ===== 輔助球開關 =====
        const checkBall = document.getElementById('oa-check-ball-side');
        if (checkBall) {
            chrome.storage.local.get(['oa_show_float_ball'], (res) => {
                checkBall.checked = (res.oa_show_float_ball !== false);
            });
            checkBall.onchange = () => {
                chrome.storage.local.set({ 'oa_show_float_ball': checkBall.checked });
                updateVisibility();
            };
        }

        // ===== 標題字體大小滑桿 =====
        const titleSizeSlider = document.getElementById('side-title-size');
        const titleSizeVal = document.getElementById('side-title-size-val');
        if (titleSizeSlider) {
            chrome.storage.local.get(['oa_setting_title_size'], (res) => {
                const v = res.oa_setting_title_size || 15;
                titleSizeSlider.value = v;
                titleSizeVal.textContent = v + 'px';
            });
            titleSizeSlider.oninput = () => {
                const v = titleSizeSlider.value;
                titleSizeVal.textContent = v + 'px';
                chrome.storage.local.set({ 'oa_setting_title_size': parseInt(v) });
                renderList();
            };
        }

        // ===== 副標題顯示模式 =====
        const subtitleCtrl = document.getElementById('side-subtitle-mode');
        if (subtitleCtrl) {
            chrome.storage.local.get(['oa_setting_subtitle_mode'], (res) => {
                const mode = res.oa_setting_subtitle_mode || 'both';
                subtitleCtrl.querySelectorAll('.oa-seg-btn').forEach(btn => {
                    btn.classList.toggle('active', btn.dataset.value === mode);
                });
            });
            subtitleCtrl.querySelectorAll('.oa-seg-btn').forEach(btn => {
                btn.onclick = () => {
                    subtitleCtrl.querySelectorAll('.oa-seg-btn').forEach(b => b.classList.remove('active'));
                    btn.classList.add('active');
                    chrome.storage.local.set({ 'oa_setting_subtitle_mode': btn.dataset.value });
                    renderList();
                };
            });
        }

        // ===== 緊湊模式 =====
        const compactCheck = document.getElementById('side-compact-mode');
        if (compactCheck) {
            chrome.storage.local.get(['oa_setting_compact'], (res) => {
                compactCheck.checked = !!res.oa_setting_compact;
            });
            compactCheck.onchange = () => {
                chrome.storage.local.set({ 'oa_setting_compact': compactCheck.checked });
                renderList();
            };
        }

        // ===== 匯出確認開關 =====
        const exportConfirmCheck = document.getElementById('side-export-confirm');
        if (exportConfirmCheck) {
            chrome.storage.local.get(['oa_setting_export_confirm'], (res) => {
                exportConfirmCheck.checked = (res.oa_setting_export_confirm !== false);
            });
            exportConfirmCheck.onchange = () => {
                chrome.storage.local.set({ 'oa_setting_export_confirm': exportConfirmCheck.checked });
            };
        }

        document.getElementById('oa-search-input').oninput = (e) => renderList(e.target.value.trim());
        document.getElementById('oa-btn-undo').onclick = () => window.postMessage({type: "OA_LINK_FILL", project: {label:"恢復備份", details:[]}}, "*");
        document.getElementById('oa-btn-clear').onclick = () => window.postMessage({type: "OA_LINK_CLEAR"}, "*");

        document.getElementById('btn-add-detail').onclick = () => addDetailRow();
        document.getElementById('btn-f-pick').onclick = () => {
            const findAndClick = (win) => {
                try {
                    const btn = win.document.getElementById('field1366309_browserbtn');
                    if (btn) { btn.click(); return true; }
                    for (let i = 0; i<win.frames.length; i++) if(findAndClick(win.frames[i])) return true;
                } catch(e){} return false;
            };
            findAndClick(window);
        };

        const saveFn = async (isNew = false) => {
            const id = isNew ? "p_"+Date.now() : (document.getElementById('in-f-id').value || "p_"+Date.now());
            const p = { id: id, managerId: document.getElementById('in-f-managerId').value, details: [] };
            fields.forEach(f => { const el = document.getElementById('in-f-'+f); if(el) p[f] = el.value.trim(); });
            document.querySelectorAll('.oa-detail-row').forEach(row => {
                p.details.push({
                    vendorName: row.querySelector('.dt-vendor').value.trim(),
                    content: row.querySelector('.dt-content').value.trim(),
                    detailCurrency: row.querySelector('.dt-currency').value,
                    amount: row.querySelector('.dt-amount').value.trim()
                });
            });
            if (!p.label || p.details.length === 0) return alert("請輸入標題並添加至少一項明細");
            
            // 🌟 獲取最新狀態防止覆寫
            const currentRes = await chrome.storage.local.get(['oa_projects_v13']);
            let currentProjects = currentRes.oa_projects_v13 || [];
            if (!isNew) {
                const globalIdx = currentProjects.findIndex(x => x.id === id);
                if(globalIdx > -1) currentProjects[globalIdx] = p; else currentProjects.push(p);
            } else {
                currentProjects.push(p);
            }
            
            await chrome.storage.local.set({ 'oa_projects_v13': currentProjects });
            alert("✅ 保存成功"); syncData(); tFill.click(); resetForm();
        };

        document.getElementById('btn-f-save-main').onclick = () => saveFn(false);
        document.getElementById('btn-f-save-as-new').onclick = () => saveFn(true);
        document.getElementById('btn-f-cancel').onclick = () => { resetForm(); tFill.click(); };
        document.getElementById('btn-f-export').onclick = exportToCSV;
        document.getElementById('btn-f-import').onclick = () => fileIn.click();

        fileIn.onchange = (e) => {
            if (!e.target.files.length) return;
            const reader = new FileReader();
            reader.onload = async (ev) => {
                const contentText = ev.target.result;
                const lines = contentText.split(/\r?\n/).filter(l => l.trim());
                if (lines.length < 2) return;
                const grouped = {};
                lines.slice(1).forEach(l => {
                    const cols = []; let c = "", q = false;
                    for (let i = 0; i < l.length; i++) {
                        if (l[i] === '"') q = !q; else if (l[i] === ',' && !q) { cols.push(c); c = ""; } else c += l[i];
                    }
                    cols.push(c);
                    const label = (cols[0] || "").trim(), reportNo = (cols[2] || "").trim();
                    if (!label && !reportNo) return;
                    const k = label + "_" + reportNo;
                    if (!grouped[k]) {
                        grouped[k] = {
                            id: "p_" + Date.now() + "_" + Math.floor(Math.random()*1000),
                            label: label, propertyName: (cols[1] || "").trim(), reportNo: reportNo,
                            manager: (cols[3] || "").trim(), projectContent: (cols[4] || "").trim(),
                            budget: (cols[5] || "").trim(), total: (cols[6] || "").trim(),
                            quoteType: (cols[7] || "").trim(), buyType: (cols[8] || "").trim(),
                            currency: (cols[9] || "0").trim(), details: [],
                            inviteCount: (cols[14] || "").trim(), replyCount: (cols[15] || "").trim(),
                            reason: (cols[16] || "").trim(), winnerName: (cols[17] || "").trim(),
                            contractCurrency: (cols[18] || "0").trim(), contractAmount: (cols[19] || "").trim()
                        };
                    }
                    if (cols[10] || cols[11] || cols[13]) {
                        grouped[k].details.push({ vendorName: (cols[10]||"").trim(), content: (cols[11]||"").trim(), detailCurrency: (cols[12]||"0").trim(), amount: (cols[13]||"").trim() });
                    }
                });
                const imported = Object.values(grouped);
                if (imported.length > 0) {
                    await chrome.storage.local.set({ 'oa_projects_v13': imported });
                    alert(`✅ 成功匯入 ${imported.length} 個項目`); syncData();
                }
            };
            reader.readAsText(e.target.files[0]);
            fileIn.value = "";
        };
    }

    async function renderList(query = "") {
        const listContainer = document.getElementById('oa-disp-list'); if (!listContainer) return;

        // 讀取設置
        const settings = await chrome.storage.local.get([
            'oa_setting_title_size',
            'oa_setting_subtitle_mode',
            'oa_setting_compact'
        ]);

        const titleSize = settings.oa_setting_title_size || 15;
        const subMode = settings.oa_setting_subtitle_mode || 'both';
        const isCompact = !!settings.oa_setting_compact;

        let filtered = projects;
        if (query) {
            const q = query.toLowerCase();
            filtered = projects.filter(p => (p.label && p.label.toLowerCase().includes(q)));
        }

        listContainer.innerHTML = filtered.length === 0 ? '<p style="text-align:center; color:#ccd; padding:40px; font-size:12px;">尚無記錄</p>' : '';
        filtered.forEach(p => {
            const div = document.createElement('div');
            div.className = 'oa-p-item';

            // 緊湊模式樣式
            if (isCompact) {
                div.style.padding = '10px 14px';
                div.style.marginBottom = '6px';
            }

            // 構建副標題內容
            let subContent = "";
            if (subMode === 'both') {
                subContent = `${p.reportNo || '--'}${p.budget ? '　$ ' + p.budget : ''}`;
            } else if (subMode === 'reportNo') {
                subContent = p.reportNo || '--';
            } else if (subMode === 'budget') {
                subContent = p.budget ? '$ ' + p.budget : '--';
            }

            div.innerHTML = `
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <div style="flex:1; overflow:hidden;">
                        <div style="font-weight:700; font-size:${titleSize}px; color:#1d1d1f; margin-bottom:4px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${p.label}</div>
                        <div style="font-size:11px; color:#86868b; letter-spacing:0.5px;">${subContent}</div>
                    </div>
                    <div style="display:flex; gap:14px; margin-left:12px; flex-shrink:0;">
                        <span class="edit-btn" style="cursor:pointer; font-size:16px; opacity:0.3; transition:opacity 0.2s;">✎</span>
                        <span class="del-btn" style="cursor:pointer; font-size:16px; color:#ff4d4f; opacity:0.3; transition:opacity 0.2s;">🗑</span>
                    </div>
                </div>
            `;
            // ... (其餘邏輯不變)
            div.onclick = (e) => {
                if (e.target.classList.contains('edit-btn') || e.target.classList.contains('del-btn')) return;
                const normP = normalizeProjectData(p);
                window.postMessage({type: "OA_LINK_FILL", project: normP}, "*");
                const originalBg = div.style.background;
                div.style.background = '#f2f2f7';
                setTimeout(() => div.style.background = originalBg || '#fff', 200);
            };
            div.querySelector('.edit-btn').onclick=(e)=>{
                e.stopPropagation(); document.getElementById('in-f-id').value=p.id;
                document.getElementById('in-f-managerId').value = p.managerId || "";
                
                fields.forEach(f=>{
                    const el=document.getElementById('in-f-'+f);
                    if(el) {
                        const val = p[f] || '';
                        if (el.tagName === 'SELECT') {
                            let found = false;
                            for(let opt of el.options) {
                                if(opt.value === String(val) || opt.text.trim() === String(val).trim()) {
                                    el.value = opt.value; found = true; break;
                                }
                            }
                            if(!found) el.value = "";
                        } else {
                            el.value = val;
                        }
                    }
                });

                document.getElementById('oa-detail-list').innerHTML='';
                if(p.details && p.details.length > 0) p.details.forEach(dt=>addDetailRow(dt)); else addDetailRow();
                document.getElementById('oa-p-title').innerText='編輯項目';
                document.getElementById('btn-f-save-main').innerText='更新項目';
                document.getElementById('btn-f-save-as-new').classList.remove('hidden');
                document.getElementById('btn-f-cancel').classList.remove('hidden');
                document.getElementById('tab-v-manage').click();
            };
            div.querySelector('.del-btn').onclick=(e)=>{
                e.stopPropagation(); if(confirm(`確定刪除「${p.label}」？`)){ 
                    const newProjects = projects.filter(x=>x.id!==p.id); 
                    chrome.storage.local.set({'oa_projects_v13': newProjects});
                }
            };
            listContainer.appendChild(div);
        });
    }

    function startPickerWatcher() {
        setInterval(() => {
            const span = document.getElementById('field1366309span'), val = document.getElementById('field1366309');
            if (span && val && val.value) {
                const nameIn = document.getElementById('in-f-manager');
                if (nameIn && nameIn.value !== span.innerText.trim()) {
                    nameIn.value = span.innerText.trim();
                    document.getElementById('in-f-managerId').value = val.value;
                }
            }
        }, 1200);
    }

    init();
})();
