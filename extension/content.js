/**
 * Version: 4.13.1 - 與 manifest / popup / bridge 版號對齊
 * 採購立項自動化工具 - 核心邏輯
 */

(function () {
    const isTop = (window.self === window.top);
    if (!isTop) return;

    let panel, ball;
    let projects = [];
    const fields = ['label', 'propertyName', 'reportNo', 'manager', 'projectContent', 'budget', 'total', 'quoteType', 'buyType', 'currency', 'inviteCount', 'replyCount', 'reason', 'winnerName', 'contractCurrency', 'contractAmount'];
    const FEISHU_KEY_TO_INPUT_ID = {
        oa_feishu_app_id: 'side-feishu-app-id',
        oa_feishu_app_secret: 'side-feishu-app-secret',
        oa_feishu_app_token: 'side-feishu-app-token',
        oa_feishu_table_id: 'side-feishu-table-id',
        oa_feishu_detail_table_id: 'side-feishu-detail-table-id'
    };

    // 映射表：將顯示文字/舊索引轉為 OA 系統真實的 Value (根據 F12 截圖校準)
    const quoteMap = { "0": "0", "1": "1", "2": "2", "3": "3", "報價邀請": "0", "招標": "1", "特殊情況 (緊急)": "2", "續約": "3" };
    const buyMap = {
        "0": "9", "1": "10", "2": "11", "3": "12", "4": "14", "5": "4", "6": "13",
        "保養維修材料": "9", "保養合約分判": "10", "工程材料": "11",
        "分判工程": "12", "後加工程": "14", "固定資產": "4", "其他": "13"
    };

    function init() {
        if (document.getElementById('oa-side-panel')) return;

        // main_bridge.js 已由 manifest「MAIN world」content_scripts 注入，避免重複載入造成雙重監聽

        injectPanelHTML();
        injectBall();
        updateVisibility();
        syncData();
        startPickerWatcher();
        listenMessages();

        // 🌟 核心同步功能：監聽儲存庫變動
        chrome.storage.onChanged.addListener((changes, area) => {
            // 主流程設定在 local，飛書設定會 local/sync 雙寫，所以兩個區域都監聽

            if (area === 'local' && changes.oa_projects_v13) {
                console.log("OA Content: Storage change detected! Syncing UI...");
                projects = changes.oa_projects_v13.newValue || [];
                renderList();
            }
            if (area === 'local' && changes.oa_show_float_ball !== undefined) {
                updateVisibility();
            }

            // popup 與懸浮面板設定同步：以 storage 變更作為單一事實來源
            const settingMap = {
                oa_setting_title_size: 'title_size',
                oa_setting_subtitle_mode: 'subtitle_mode',
                oa_setting_compact: 'compact',
                oa_setting_export_confirm: 'export_confirm',
                oa_setting_export_dir: 'export_dir',
                oa_setting_sort_mode: 'sort_mode'
            };
            Object.keys(settingMap).forEach(storageKey => {
                if (changes[storageKey]) {
                    applySettingUpdate(settingMap[storageKey], changes[storageKey].newValue);
                }
            });

            // 飛書設定：來自 popup 的更新需即時反映到懸浮版輸入框
            Object.keys(FEISHU_KEY_TO_INPUT_ID).forEach((storageKey) => {
                if (!changes[storageKey]) return;
                const input = document.getElementById(FEISHU_KEY_TO_INPUT_ID[storageKey]);
                if (input) input.value = changes[storageKey].newValue || '';
            });
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

    function applySettingUpdate(key, value) {
        if (key === "title_size") {
            const slider = document.getElementById('side-title-size');
            const valLab = document.getElementById('side-title-size-val');
            if (slider) slider.value = value;
            if (valLab) valLab.textContent = value + 'px';
            renderList();
        } else if (key === "subtitle_mode") {
            const ctrl = document.getElementById('side-subtitle-mode');
            if (ctrl) {
                ctrl.querySelectorAll('.oa-seg-btn').forEach(b => b.classList.toggle('active', b.dataset.value === value));
            }
            renderList();
        } else if (key === "compact") {
            const cb = document.getElementById('side-compact-mode');
            if (cb) cb.checked = !!value;
            renderList();
        } else if (key === "export_confirm") {
            const cb = document.getElementById('side-export-confirm');
            if (cb) cb.checked = (value !== false);
        } else if (key === "export_dir") {
            const de = document.getElementById('side-export-dir');
            if (de) de.value = value || "";
        } else if (key === "sort_mode") {
            const sortSel = document.getElementById('oa-sort-select');
            if (sortSel) sortSel.value = value || 'default';
            renderList(document.getElementById('oa-search-input')?.value.trim() || '');
        }
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
                applySettingUpdate(msg.key, msg.value);
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
                <button class="oa-tool-btn primary" id="oa-btn-feishu-top">🐦 飛書</button>
            </div>

            <div class="oa-p-tabs">
                <div class="oa-p-tab active" id="tab-v-fill">快速填充</div>
                <div class="oa-p-tab" id="tab-v-manage">項目管理</div>
                <div class="oa-p-tab" id="tab-v-settings">更多設置</div>
            </div>

            <div id="view-v-fill" class="oa-panel-body">
                <div style="display:flex; gap:6px; margin-bottom:12px; align-items:center;">
                    <input type="text" id="oa-search-input" class="oa-search-input" placeholder="🔍 搜尋項目或經理..." style="flex:1; min-width:0;">
                    <select id="oa-sort-select" title="排序方式" style="flex-shrink:0; height:34px; padding:0 6px; border-radius:10px; border:1px solid var(--ym-border); background:#fff; font-size:11px; color:var(--ym-text-secondary); cursor:pointer; outline:none;">
                        <option value="default">📋 預設</option>
                        <option value="created_desc">🕐 新增↓</option>
                        <option value="created_asc">🕐 新增↑</option>
                        <option value="updated_desc">✏️ 修改↓</option>
                        <option value="updated_asc">✏️ 修改↑</option>
                    </select>
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

                <!-- 項目ID 顯示 -->
                <div id="in-f-projectId-wrap" style="display:flex; align-items:center; gap:8px; margin-bottom:12px; padding:7px 12px; background:var(--ym-primary-soft); border-radius:10px; border:1px solid rgba(0,122,255,0.15);">
                    <span style="font-size:11px; color:var(--ym-primary); font-weight:600; flex-shrink:0;">🔖 項目ID</span>
                    <span id="in-f-projectId-display" style="font-size:12px; color:var(--ym-text-secondary); font-family:monospace; letter-spacing:0.5px;">（儲存後自動生成）</span>
                    <input type="hidden" id="in-f-projectId">
                </div>

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

                <div class="s-row">
                    <div><div class="s-label">匯出子資料夾 (選填)</div><div class="s-desc">在下載資料夾中建立專屬目錄</div></div>
                    <div style="display:flex; align-items:center;">
                        <input type="text" id="side-export-dir" class="oa-search-input" style="width:120px; padding:6px 10px; font-size:12px; text-align:right;" placeholder="例如: OA_Backup">
                    </div>
                </div>

                <div class="s-section" id="side-feishu-section">🐦 飛書 BITABLE 同步</div>
                <div style="display:grid; grid-template-columns:1fr 1fr; gap:8px; margin-bottom:8px;">
                    <input type="text" id="side-feishu-app-id" class="oa-search-input" placeholder="App ID (cli_xxx)">
                    <input type="password" id="side-feishu-app-secret" class="oa-search-input" placeholder="App Secret">
                </div>
                <input type="text" id="side-feishu-app-token" class="oa-search-input" style="margin-bottom:8px;" placeholder="Bitable App Token (sqa...)">
                <input type="text" id="side-feishu-table-id" class="oa-search-input" style="margin-bottom:8px;" placeholder="Table ID (tbl...)">
                <input type="text" id="side-feishu-detail-table-id" class="oa-search-input" style="margin-bottom:8px;" placeholder="明細表 Table ID (選填)">
                <button id="side-feishu-sync-btn" class="oa-p-btn-action" style="margin-top:4px;">🐦 立即同步到飛書</button>
                <div id="side-feishu-sync-status" style="display:none; margin-top:8px; padding:8px 10px; border-radius:10px; font-size:12px; line-height:1.6;"></div>
            </div>

            <div style="padding:8px; text-align:center; font-size:10px; color:#c0c0cc; background:#fff; border-top:1px solid var(--ym-border); letter-spacing:0.5px; flex-shrink:0;">v4.13.1</div>
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
            // 若目前元素未持有此 pointer 的 capture 則跳過
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
                const isOpen = panel.classList.toggle('open');
                if (isOpen) {
                    main.classList.add('ball-hidden');
                } else {
                    main.classList.remove('ball-hidden');
                }
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
            const el = document.getElementById('in-f-' + f);
            if (el) {
                if (el.tagName === 'SELECT') {
                    // 下拉選單默認值處理
                    if (f === 'currency' || f === 'contractCurrency') el.value = "0";
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

        // 自動遷移：為沒有時間戳的舊項目補上時間
        let needsSave = false;
        projects.forEach(p => {
            if (!p.createdAt) {
                const tsMatch = p.id && p.id.match(/p_(\d+)/);
                const ts = tsMatch ? parseInt(tsMatch[1]) : null;
                const guessedTime = (ts && ts > 1000000000000) ? new Date(ts).toISOString() : null;
                p.createdAt = guessedTime || new Date().toISOString();
                p.updatedAt = p.createdAt;
                needsSave = true;
            }
        });

        // 自動遷移：為沒有 projectId 的舊項目自動補填 ID
        // 順序處理，使序號能正確逐步累加
        projects.forEach(p => {
            if (!p.projectId) {
                p.projectId = generateProjectId(p.buyType, p.reportNo, p.createdAt, projects);
                needsSave = true;
            }
        });

        if (needsSave) {
            await chrome.storage.local.set({ 'oa_projects_v13': projects });
        }

        renderList();
    }

    // ===== 項目 ID 生成（方案 B：Q1322-24JC-MAT-001）=====
    const BUY_TYPE_CODE = { "9": "MNT", "10": "SVC", "11": "MAT", "12": "SUB", "14": "ADD", "4": "FAR", "13": "OTH" };
    function generateProjectId(buyType, reportNo, createdAt, allProjects) {
        const typeCode = BUY_TYPE_CODE[buyType] || 'OTH';

        // 解析報價編號，例：MS/Q1322/24/jc → Q1322-24JC
        let reportPart;
        const rn = (reportNo || '').trim();
        if (rn) {
            const parts = rn.split('/');
            if (parts.length >= 3) {
                const qNum = parts[1] || '';
                const yr = parts[2] || '';
                const init = (parts[3] || '').toUpperCase();
                reportPart = `${qNum}-${yr}${init}`;
            } else {
                reportPart = rn.replace(/[\/\\:*?"<>| ]/g, '').toUpperCase().slice(0, 12);
            }
        } else {
            const d = createdAt ? new Date(createdAt) : new Date();
            const pad = n => String(n).padStart(2, '0');
            reportPart = `NA-${String(d.getFullYear()).slice(2)}${pad(d.getMonth() + 1)}`;
        }

        const fullPfx = `${reportPart}-${typeCode}-`;
        let maxSeq = 0;
        (allProjects || []).forEach(p => {
            if (p.projectId && p.projectId.startsWith(fullPfx)) {
                const seq = parseInt(p.projectId.slice(fullPfx.length)) || 0;
                if (seq > maxSeq) maxSeq = seq;
            }
        });
        return fullPfx + String(maxSeq + 1).padStart(3, '0');
    }

    function buildCsvDataUrl() {
        if (projects.length === 0) { alert("尚無數據可匯出"); return null; }
        const h = ["項目標題", "物業名稱", "報價編號", "項目經理", "項目內容", "立項預算", "合約總價", "報價形式", "採購類別", "幣種", "承判商", "報價內容", "明細幣種", "金額", "邀請公司", "有效報價", "推荐理由", "中標公司", "合約幣種", "合約金額", "項目ID", "新增日期時間", "修改日期時間"];
        const rows = [];
        projects.forEach(p => {
            (p.details || [{}]).forEach(dt => {
                const row = [p.label, p.propertyName, p.reportNo, p.manager, p.projectContent, p.budget, p.total, p.quoteType, p.buyType, p.currency, dt.vendorName, dt.content, dt.detailCurrency, dt.amount, p.inviteCount, p.replyCount, p.reason, p.winnerName, p.contractCurrency, p.contractAmount,
                p.projectId || '', p.createdAt || '', p.updatedAt || ''];
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
        if (projects.length === 0) return alert("尚無數據可匯出");

        const setting = await chrome.storage.local.get(['oa_setting_export_confirm', 'oa_setting_export_dir']);
        const needsConfirm = (setting.oa_setting_export_confirm !== false);
        const exportDir = (setting.oa_setting_export_dir || "").trim();
        const basePath = exportDir ? (exportDir + '/') : '';

        if (!needsConfirm) {
            triggerDownload(basePath + 'OA_Projects.csv', 'overwrite');
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
            triggerDownload(basePath + 'OA_Projects.csv', 'overwrite');
        };
        document.getElementById('oa-exp-new').onclick = () => {
            close();
            const ts = new Date().toISOString().slice(0, 16).replace('T', '_').replace(':', '-');
            triggerDownload(basePath + `OA_Projects_${ts}.csv`, 'uniquify');
        };
    }

    function bindEvents() {
        document.getElementById('oa-btn-refresh').onclick = syncData;
        document.getElementById('oa-panel-close').onclick = () => {
            panel.classList.remove('open');
            const mainBall = document.querySelector('.oa-ball-main');
            if (mainBall) mainBall.classList.remove('ball-hidden');
        };

        document.getElementById('oa-btn-add').onclick = () => { resetForm(); document.getElementById('tab-v-manage').click(); };
        document.getElementById('oa-btn-save-top').onclick = () => document.getElementById('btn-f-save-main').click();
        document.getElementById('oa-btn-export-top').onclick = exportToCSV;
        document.getElementById('oa-btn-feishu-top').onclick = () => {
            // 與 popup 一致：快速跳轉到「更多設置」中的飛書區塊
            tSettings.click();
            setTimeout(() => {
                const el = document.getElementById('side-feishu-section');
                if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }, 100);
        };
        const fileIn = document.getElementById('in-f-file-side');
        document.getElementById('oa-btn-import-top').onclick = () => fileIn.click();

        // 注意：上方快捷列（新增/匯入/匯出/保存/飛書）需與 popup 保持一致
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

        // ===== 匯出子目錄 =====
        const exportDirSide = document.getElementById('side-export-dir');
        if (exportDirSide) {
            chrome.storage.local.get(['oa_setting_export_dir'], (res) => {
                exportDirSide.value = res.oa_setting_export_dir || '';
            });
            exportDirSide.addEventListener('blur', () => {
                const val = exportDirSide.value.trim().replace(/[\/:*?"<>|]/g, '');
                exportDirSide.value = val;
                chrome.storage.local.set({ 'oa_setting_export_dir': val });
            });
        }

        // ===== 飛書設定（懸浮版與彈窗版一致）=====
        const feishuInputIds = ['side-feishu-app-id', 'side-feishu-app-secret', 'side-feishu-app-token', 'side-feishu-table-id', 'side-feishu-detail-table-id'];
        const feishuKeys = ['oa_feishu_app_id', 'oa_feishu_app_secret', 'oa_feishu_app_token', 'oa_feishu_table_id', 'oa_feishu_detail_table_id'];
        const feishuInputs = feishuInputIds.map(id => document.getElementById(id));
        const feishuSyncBtn = document.getElementById('side-feishu-sync-btn');
        const feishuStatus = document.getElementById('side-feishu-sync-status');
        let feishuSaveTimer = null;

        const collectFeishuValues = () => {
            const obj = {};
            feishuKeys.forEach((key, i) => { obj[key] = (feishuInputs[i]?.value || '').trim(); });
            return obj;
        };

        const persistFeishuValues = (values) => {
            chrome.storage.local.set(values);
            chrome.storage.sync.set(values);
        };

        const schedulePersistFeishuValues = () => {
            if (feishuSaveTimer) clearTimeout(feishuSaveTimer);
            feishuSaveTimer = setTimeout(() => {
                persistFeishuValues(collectFeishuValues());
                feishuSaveTimer = null;
            }, 300);
        };

        const showFeishuStatus = (type, html) => {
            if (!feishuStatus) {
                console.warn('OA: feishuStatus element not found');
                return;
            }
            feishuStatus.style.display = 'block';
            feishuStatus.style.visibility = 'visible';
            feishuStatus.style.opacity = '1';
            feishuStatus.style.borderRadius = '10px';
            feishuStatus.style.fontSize = '12px';
            feishuStatus.style.lineHeight = '1.6';
            feishuStatus.style.boxSizing = 'border-box';
            if (type === 'loading') {
                feishuStatus.style.background = 'rgba(51,112,255,0.08)';
                feishuStatus.style.color = '#3370FF';
                feishuStatus.style.border = '1px solid rgba(51,112,255,0.2)';
            } else if (type === 'success') {
                feishuStatus.style.background = 'rgba(52,199,89,0.08)';
                feishuStatus.style.color = '#1a8a35';
                feishuStatus.style.border = '1px solid rgba(52,199,89,0.3)';
            } else {
                feishuStatus.style.background = 'rgba(255,59,48,0.08)';
                feishuStatus.style.color = '#c0392b';
                feishuStatus.style.border = '1px solid rgba(255,59,48,0.2)';
            }
            feishuStatus.innerHTML = html;
        };

        (async () => {
            const [localRes, syncRes] = await Promise.all([
                chrome.storage.local.get(feishuKeys),
                chrome.storage.sync.get(feishuKeys)
            ]);
            const merged = {};
            feishuKeys.forEach((key) => {
                const localVal = (localRes[key] || '').trim();
                const syncVal = (syncRes[key] || '').trim();
                merged[key] = localVal || syncVal || '';
            });
            feishuKeys.forEach((key, i) => {
                if (feishuInputs[i]) feishuInputs[i].value = merged[key] || '';
            });
            if (feishuKeys.some(k => merged[k])) persistFeishuValues(merged);
        })();

        feishuInputs.forEach((el) => {
            if (!el) return;
            el.addEventListener('input', schedulePersistFeishuValues);
            el.addEventListener('change', schedulePersistFeishuValues);
            el.addEventListener('blur', () => persistFeishuValues(collectFeishuValues()));
        });

        if (feishuSyncBtn) {
            feishuSyncBtn.onclick = async () => {
                const cfg = collectFeishuValues();
                persistFeishuValues(cfg);
                const config = {
                    appId: cfg.oa_feishu_app_id || '',
                    appSecret: cfg.oa_feishu_app_secret || '',
                    appToken: cfg.oa_feishu_app_token || '',
                    tableId: cfg.oa_feishu_table_id || '',
                    detailTableId: cfg.oa_feishu_detail_table_id || ''
                };

                if (!config.appId || !config.appSecret || !config.appToken || !config.tableId) {
                    showFeishuStatus('error', '❌ 請先填寫完整的飛書設定（App ID、App Secret、App Token、Table ID）');
                    return;
                }
                if (projects.length === 0) {
                    showFeishuStatus('error', '❌ 尚無項目資料可同步');
                    return;
                }

                feishuSyncBtn.disabled = true;
                feishuSyncBtn.textContent = '⏳ 同步中...';
                showFeishuStatus('loading', `🐦 正在連線飛書並同步 ${projects.length} 個項目，請稍候...`);
                try {
                    const result = await new Promise((resolve, reject) => {
                        chrome.runtime.sendMessage(
                            { type: 'FEISHU_SYNC', config: config, projects: projects },
                            (res) => {
                                if (chrome.runtime.lastError) reject(new Error(chrome.runtime.lastError.message));
                                else if (!res || !res.success) reject(new Error(res?.error || '未知錯誤'));
                                else resolve(res);
                            }
                        );
                    });

                    let successHtml =
                        `✅ <strong>同步成功</strong><br>` +
                        `📥 新增：${result.created} 條` +
                        `&emsp;✏️ 更新：${result.updated} 條` +
                        `&emsp;📄 共計：${result.total} 條`;
                    if (result.downloaded > 0) {
                        successHtml += `<br>📥 從飛書下載：${result.downloaded} 條（新項目已合併到本地）`;
                    }
                    if (result.detailCreated > 0) {
                        successHtml += `<br>🔸 明細新增：${result.detailCreated} 條（共 ${result.detailCount} 條明細）`;
                    } else if (result.detailCount > 0) {
                        successHtml += `<br>⚠️ 有 ${result.detailCount} 條明細但未寫入（請確認明細表 Table ID 設定）`;
                    }
                    if (result.detailDeleted > 0) {
                        successHtml += `<br>🧹 已刪除舊明細：${result.detailDeleted} 條（覆蓋式同步）`;
                    }
                    if (result.missingFields && result.missingFields.length > 0) {
                        successHtml +=
                            `<div style="margin-top:10px; padding:8px 10px; background:rgba(255,204,0,0.12);` +
                            `border:1px solid rgba(255,170,0,0.4); border-radius:8px; color:#7a5c00; font-size:11px; line-height:1.8;">` +
                            `⚠️ <strong>以下 ${result.missingFields.length} 個欄位在飛書表格中找不到</strong>，相關數據未被寫入：<br>` +
                            result.missingFields.map(f => `&nbsp;• ${f}`).join('<br>') +
                            `<br><span style="opacity:0.75">▶ 請在多維表格中新增對應欄位後重新同步</span></div>`;
                    }
                    showFeishuStatus('success', successHtml);

                    // 如果有從飛書下載的新項目，自動合併到本地存儲
                    if (result.feishuProjects && result.feishuProjects.length > 0) {
                        const mergedProjects = [...projects, ...result.feishuProjects];
                        await chrome.storage.local.set({ 'oa_projects_v13': mergedProjects });
                        syncData(); // 重新讀取並渲染列表
                    }
                } catch (err) {
                    showFeishuStatus('error', `❌ <strong>同步失敗</strong><br>${err.message}`);
                } finally {
                    feishuSyncBtn.disabled = false;
                    feishuSyncBtn.textContent = '🐦 立即同步到飛書';
                }
            };
        }

        document.getElementById('oa-search-input').oninput = (e) => renderList(e.target.value.trim());

        // ===== 排序選單 =====
        const sortSel = document.getElementById('oa-sort-select');
        if (sortSel) {
            chrome.storage.local.get(['oa_setting_sort_mode'], (res) => {
                sortSel.value = res.oa_setting_sort_mode || 'default';
            });
            sortSel.onchange = () => {
                chrome.storage.local.set({ 'oa_setting_sort_mode': sortSel.value });
                renderList(document.getElementById('oa-search-input')?.value.trim() || '');
            };
        }
        document.getElementById('oa-btn-undo').onclick = () => window.postMessage({ type: "OA_LINK_FILL", project: { label: "恢復備份", details: [] } }, "*");
        document.getElementById('oa-btn-clear').onclick = () => window.postMessage({ type: "OA_LINK_CLEAR" }, "*");

        document.getElementById('btn-add-detail').onclick = () => addDetailRow();
        document.getElementById('btn-f-pick').onclick = () => {
            const findAndClick = (win) => {
                try {
                    const btn = win.document.getElementById('field1366309_browserbtn');
                    if (btn) { btn.click(); return true; }
                    for (let i = 0; i < win.frames.length; i++) if (findAndClick(win.frames[i])) return true;
                } catch (e) { } return false;
            };
            findAndClick(window);
        };

        const saveFn = async (isNew = false) => {
            const id = isNew ? "p_" + Date.now() : (document.getElementById('in-f-id').value || "p_" + Date.now());
            const p = { id: id, managerId: document.getElementById('in-f-managerId').value, details: [] };
            fields.forEach(f => { const el = document.getElementById('in-f-' + f); if (el) p[f] = el.value.trim(); });
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
            const now = new Date().toISOString();
            if (!isNew) {
                const globalIdx = currentProjects.findIndex(x => x.id === id);
                if (globalIdx > -1) {
                    p.createdAt = currentProjects[globalIdx].createdAt || now;
                    p.updatedAt = now;
                    p.projectId = currentProjects[globalIdx].projectId || generateProjectId(p.buyType, p.reportNo, p.createdAt, currentProjects);
                    currentProjects[globalIdx] = p;
                } else {
                    p.createdAt = now;
                    p.updatedAt = now;
                    currentProjects.push(p);
                    p.projectId = generateProjectId(p.buyType, p.reportNo, p.createdAt, currentProjects);
                }
            } else {
                p.createdAt = now;
                p.updatedAt = now;
                currentProjects.push(p);
                p.projectId = generateProjectId(p.buyType, p.reportNo, p.createdAt, currentProjects);
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
                        const importTime = new Date().toISOString();
                        const csvProjectId = (cols[20] || '').trim();
                        const csvCreatedAt = (cols[21] || '').trim();
                        const csvUpdatedAt = (cols[22] || '').trim();
                        grouped[k] = {
                            id: "p_" + Date.now() + "_" + Math.floor(Math.random() * 1000),
                            label: label, propertyName: (cols[1] || "").trim(), reportNo: reportNo,
                            manager: (cols[3] || "").trim(), projectContent: (cols[4] || "").trim(),
                            budget: (cols[5] || "").trim(), total: (cols[6] || "").trim(),
                            quoteType: (cols[7] || "").trim(), buyType: (cols[8] || "").trim(),
                            currency: (cols[9] || "0").trim(), details: [],
                            inviteCount: (cols[14] || "").trim(), replyCount: (cols[15] || "").trim(),
                            reason: (cols[16] || "").trim(), winnerName: (cols[17] || "").trim(),
                            contractCurrency: (cols[18] || "0").trim(), contractAmount: (cols[19] || "").trim(),
                            projectId: csvProjectId || '',
                            createdAt: csvCreatedAt || importTime,
                            updatedAt: csvUpdatedAt || importTime
                        };
                    }
                    if (cols[10] || cols[11] || cols[13]) {
                        grouped[k].details.push({ vendorName: (cols[10] || "").trim(), content: (cols[11] || "").trim(), detailCurrency: (cols[12] || "0").trim(), amount: (cols[13] || "").trim() });
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
            'oa_setting_compact',
            'oa_setting_sort_mode'
        ]);

        const titleSize = settings.oa_setting_title_size || 15;
        const subMode = settings.oa_setting_subtitle_mode || 'both';
        const isCompact = !!settings.oa_setting_compact;
        const sortMode = settings.oa_setting_sort_mode || 'default';

        // 同步排序選單顯示
        const sortSelEl = document.getElementById('oa-sort-select');
        if (sortSelEl && sortSelEl.value !== sortMode) sortSelEl.value = sortMode;

        let filtered = projects;
        if (query) {
            const q = query.toLowerCase();
            filtered = projects.filter(p =>
                (p.label && p.label.toLowerCase().includes(q)) ||
                (p.manager && p.manager.toLowerCase().includes(q)) ||
                (p.details && p.details.some(dt => dt.vendorName && dt.vendorName.toLowerCase().includes(q)))
            );
        }

        // ===== 排序邏輯 =====
        const getTime = (p, field) => p[field] || p.createdAt || '';
        if (sortMode !== 'default') {
            filtered = [...filtered];
            if (sortMode === 'created_desc') filtered.sort((a, b) => getTime(b, 'createdAt') > getTime(a, 'createdAt') ? 1 : -1);
            if (sortMode === 'created_asc') filtered.sort((a, b) => getTime(a, 'createdAt') > getTime(b, 'createdAt') ? 1 : -1);
            if (sortMode === 'updated_desc') filtered.sort((a, b) => getTime(b, 'updatedAt') > getTime(a, 'updatedAt') ? 1 : -1);
            if (sortMode === 'updated_asc') filtered.sort((a, b) => getTime(a, 'updatedAt') > getTime(b, 'updatedAt') ? 1 : -1);
        }

        listContainer.innerHTML = filtered.length === 0 ? '<p style="text-align:center; color:#ccd; padding:40px; font-size:12px;">尚無記錄</p>' : '';

        const grouped = {};
        const unGrouped = [];
        filtered.forEach(p => {
            const rNo = (p.reportNo || '').trim();
            if (rNo) {
                if (!grouped[rNo]) grouped[rNo] = [];
                grouped[rNo].push(p);
            } else {
                unGrouped.push(p);
            }
        });

        const createItemEl = (p) => {
            const div = document.createElement('div');
            div.className = 'oa-p-item';

            // 緊湊模式樣式
            if (isCompact) {
                div.style.padding = '10px 14px';
                div.style.marginBottom = '6px';
            }

            // 格式化日期時間（短格式：MM-DD HH:mm）
            function formatDateTime(isoStr) {
                if (!isoStr) return '';
                try {
                    const d = new Date(isoStr);
                    const pad = n => String(n).padStart(2, '0');
                    return `${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
                } catch (e) { return ''; }
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

            // 構建日期時間行
            const createdStr = p.createdAt ? formatDateTime(p.createdAt) : '';
            const updatedStr = p.updatedAt ? formatDateTime(p.updatedAt) : '';
            const isModified = p.createdAt && p.updatedAt && p.createdAt !== p.updatedAt;
            let dateHtml = '';
            if (createdStr) {
                dateHtml = `<div style="font-size:10px; color:#b0b0ba; margin-top:2px;">🕐 ${createdStr}`;
                if (isModified) dateHtml += `　✏️ ${updatedStr}`;
                dateHtml += `</div>`;
            }

            div.innerHTML = `
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <div style="flex:1; overflow:hidden;">
                        <div style="font-weight:700; font-size:${titleSize}px; color:#1d1d1f; margin-bottom:4px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${p.label}</div>
                        <div style="font-size:11px; color:#86868b; letter-spacing:0.5px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${subContent}</div>
                        ${dateHtml}
                    </div>
                    <div style="display:flex; gap:14px; margin-left:12px; flex-shrink:0;">
                        <span class="edit-btn" style="cursor:pointer; font-size:16px; opacity:0.3; transition:opacity 0.2s;">✎</span>
                        <span class="del-btn" style="cursor:pointer; font-size:16px; color:#ff4d4f; opacity:0.3; transition:opacity 0.2s;">🗑</span>
                    </div>
                </div>
            `;

            div.onclick = (e) => {
                if (e.target.classList.contains('edit-btn') || e.target.classList.contains('del-btn')) return;
                const normP = normalizeProjectData(p);
                window.postMessage({ type: "OA_LINK_FILL", project: normP }, "*");
                const originalBg = div.style.background;
                div.style.background = '#f2f2f7';
                setTimeout(() => div.style.background = originalBg || '#fff', 200);
            };
            div.querySelector('.edit-btn').onclick = (e) => {
                e.stopPropagation(); document.getElementById('in-f-id').value = p.id;
                document.getElementById('in-f-managerId').value = p.managerId || "";

                fields.forEach(f => {
                    const el = document.getElementById('in-f-' + f);
                    if (el) {
                        const val = p[f] || '';
                        if (el.tagName === 'SELECT') {
                            let found = false;
                            for (let opt of el.options) {
                                if (opt.value === String(val) || opt.text.trim() === String(val).trim()) {
                                    el.value = opt.value; found = true; break;
                                }
                            }
                            if (!found) el.value = "";
                        } else {
                            el.value = val;
                        }
                    }
                });

                document.getElementById('oa-detail-list').innerHTML = '';
                if (p.details && p.details.length > 0) p.details.forEach(dt => addDetailRow(dt)); else addDetailRow();
                // 更新項目ID顯示
                const pidDisp = document.getElementById('in-f-projectId-display');
                const pidHid = document.getElementById('in-f-projectId');
                if (pidDisp) pidDisp.textContent = p.projectId || '（儲存後生成）';
                if (pidHid) pidHid.value = p.projectId || '';
                document.getElementById('oa-p-title').innerText = '編輯項目';
                document.getElementById('btn-f-save-main').innerText = '更新項目';
                document.getElementById('btn-f-save-as-new').classList.remove('hidden');
                document.getElementById('btn-f-cancel').classList.remove('hidden');
                document.getElementById('tab-v-manage').click();
            };
            div.querySelector('.del-btn').onclick = (e) => {
                e.stopPropagation(); if (confirm(`確定刪除「${p.label}」？`)) {
                    const newProjects = projects.filter(x => x.id !== p.id);
                    chrome.storage.local.set({ 'oa_projects_v13': newProjects });
                }
            };
            return div;
        };

        // ===== 資料夾排序 =====
        const getFolderTime = (rNo, field) =>
            grouped[rNo].reduce((mx, p) => { const t = p[field] || p.createdAt || ''; return t > mx ? t : mx; }, '');

        let folderKeys = Object.keys(grouped);
        if (sortMode === 'created_desc') folderKeys.sort((a, b) => getFolderTime(b, 'createdAt') > getFolderTime(a, 'createdAt') ? 1 : -1);
        if (sortMode === 'created_asc') folderKeys.sort((a, b) => getFolderTime(a, 'createdAt') > getFolderTime(b, 'createdAt') ? 1 : -1);
        if (sortMode === 'updated_desc') folderKeys.sort((a, b) => getFolderTime(b, 'updatedAt') > getFolderTime(a, 'updatedAt') ? 1 : -1);
        if (sortMode === 'updated_asc') folderKeys.sort((a, b) => getFolderTime(a, 'updatedAt') > getFolderTime(b, 'updatedAt') ? 1 : -1);

        folderKeys.forEach(rNo => {
            const groupItems = grouped[rNo];
            const folder = document.createElement('div');
            folder.className = 'oa-folder';
            folder.innerHTML = `
                <div class="oa-folder-header">
                    <div class="oa-folder-title">
                        <span>📁</span>
                        <span>${rNo}</span>
                        <span class="oa-folder-count">${groupItems.length}</span>
                    </div>
                    <span class="oa-folder-arrow">▼</span>
                </div>
                <div class="oa-folder-content"></div>
            `;
            folder.querySelector('.oa-folder-header').onclick = () => folder.classList.toggle('open');
            const content = folder.querySelector('.oa-folder-content');
            groupItems.forEach(p => content.appendChild(createItemEl(p)));
            listContainer.appendChild(folder);
        });

        unGrouped.forEach(p => listContainer.appendChild(createItemEl(p)));
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
