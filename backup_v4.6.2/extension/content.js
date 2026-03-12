/**
 * Weaver OA Persistent Floating Bubble (YouMind Style)
 * Version: 4.6.2
 * 核心：實現帶有視覺回饋的人員 ID 快取與自動匹配
 */

(function () {
    const isTop = (window.self === window.top);
    if (!isTop) return;

    let panel, ball;
    let projects = [];
    let fields = ['label', 'propertyName', 'reportNo', 'manager', 'projectContent', 'budget', 'total', 'quoteType', 'buyType', 'currency', 'inviteCount', 'replyCount', 'reason', 'winnerName', 'contractCurrency', 'contractAmount'];
    let inputs = {};
    let lastPickedId = "";

    function init() {
        if (document.getElementById('oa-side-panel')) return;

        const script = document.createElement('script');
        script.src = chrome.runtime.getURL('main_bridge.js');
        (document.head || document.documentElement).appendChild(script);

        injectPanelHTML();
        injectBall();
        updateVisibility(); // 根據設定顯示/隱藏
        syncData();
        startPickerWatcher();
        listenMessages(); // 監聽來自 popup 的指令
    }

    async function updateVisibility() {
        const res = await chrome.storage.local.get(['oa_show_float_ball']);
        const visible = (res.oa_show_float_ball !== false);
        const display = visible ? 'flex' : 'none';
        if (ball) ball.style.display = display;
        if (panel && !panel.classList.contains('open')) panel.style.display = display;
    }

    function listenMessages() {
        chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
            if (msg.type === "TOGGLE_FLOAT_BALL") {
                const display = msg.visible ? 'flex' : 'none';
                if (ball) ball.style.display = display;
                if (panel && !panel.classList.contains('open')) panel.style.display = display;
                sendResponse({ success: true });
            } else if (msg.type === "OA_LINK_FILL" || msg.type === "OA_LINK_CLEAR") {
                window.postMessage(msg, "*");
                const bc = (win) => { for (let i = 0; i < win.frames.length; i++) { try { win.frames[i].postMessage(msg, "*"); bc(win.frames[i]); } catch (e) { } } };
                bc(window);
                sendResponse({ success: true });
            }
            return true; // Keep message channel open for async response
        });
    }

    function injectPanelHTML() {
        panel = document.createElement('div');
        panel.id = 'oa-side-panel';
        panel.innerHTML = `
            <div id="oa-panel-header">
                <div class="oa-header-left">
                    <div class="oa-logo">un</div>
                    <div class="oa-title-group">
                        <span style="font-size:16px;">💾</span>
                        <span>採購一鍵通</span>
                        <span style="font-size:10px; opacity:0.5;">▼</span>
                    </div>
                </div>
                <div class="oa-header-right">
                    <span class="oa-header-icon" title="更多設置">···</span>
                    <span class="oa-header-icon" id="oa-btn-refresh" title="同步數據">↻</span>
                    <span class="oa-header-icon close" id="oa-panel-close" title="隱藏視窗">✕</span>
                </div>
            </div>

            <div class="oa-action-toolbar" style="padding:10px 15px; display:flex; gap:6px; background:#fff; border-bottom:1px solid #eee;">
                <button class="oa-tool-btn" id="oa-btn-add" title="新增項目" style="flex:1; padding:6px 0; border:1px solid #764ba2; color:#764ba2; background:#fff; border-radius:6px; cursor:pointer; font-size:11px;">新增</button>
                <button class="oa-tool-btn" id="oa-btn-import" title="匯入 CSV" style="flex:1; padding:6px 0; border:1px solid #eee; color:#666; background:#fff; border-radius:6px; cursor:pointer; font-size:11px;">匯入</button>
                <button class="oa-tool-btn" id="oa-btn-export" title="匯出 CSV" style="flex:1; padding:6px 0; border:1px solid #eee; color:#666; background:#fff; border-radius:6px; cursor:pointer; font-size:11px;">匯出</button>
                <button class="oa-tool-btn" id="oa-btn-save" title="保存當前" style="flex:1.2; padding:6px 0; border:none; color:#fff; background:linear-gradient(135deg, #764ba2 0%, #667eea 100%); border-radius:6px; cursor:pointer; font-size:11px; font-weight:bold;">保存</button>
            </div>

            <div class="oa-p-tabs">
                <div class="oa-p-tab active" id="tab-v-fill">快速填充</div>
                <div class="oa-p-tab" id="tab-v-manage">項目管理</div>
            </div>

            <div id="view-v-fill" class="oa-panel-body">
                <div class="oa-section-header" style="display:flex; justify-content:space-between; align-items:center;">
                    <div class="oa-section-title">📖 填充列表</div>
                    <div class="oa-search-box" style="margin-left:auto;">
                        <input type="text" id="oa-search-input" placeholder="搜尋項目或經理..." style="width:120px; padding:4px 8px; border-radius:12px; border:1px solid #ddd; font-size:12px; outline:none;">
                    </div>
                </div>
                <!-- 新增：快速填充視圖按鈕工具欄 -->
                <div class="oa-fill-toolbar" style="padding:0 10px 10px 10px; display:flex; gap:8px;">
                    <button id="oa-btn-undo" class="oa-tool-btn" style="flex:1; padding:5px; border-radius:6px; border:1px solid #eee; background:#fff; cursor:pointer; font-size:11px;">↩️ 恢復</button>
                    <button id="oa-btn-clear" class="oa-tool-btn" style="flex:1; padding:5px; border-radius:6px; border:1px solid #eee; background:#fff; cursor:pointer; font-size:11px;">🧹 清空</button>
                </div>
                <div id="oa-disp-list"></div>
            </div>

            <div id="view-v-manage" class="oa-panel-body hidden">
                <h4 id="oa-p-title" style="margin-top:0;">新增項目</h4>
                <input type="hidden" id="in-f-id">
                <input type="hidden" id="in-f-managerId">
                
                <div class="oa-p-group"><label>項目標題 (顯示用)</label><input type="text" id="in-f-label" placeholder="例如：信成 - 電器材料"></div>
                
                <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
                    <div class="oa-p-group"><label>物業名稱</label><input type="text" id="in-f-propertyName" placeholder="例如: 九龍灣車廠"></div>
                    <div class="oa-p-group">
                        <label>項目經理</label>
                        <div style="display:flex;gap:4px;">
                            <input type="text" id="in-f-manager" style="flex:1;" placeholder="輸入姓名或選取">
                            <button id="btn-f-pick" style="padding:4px 10px; border-radius:8px; border:1px solid #ddd; cursor:pointer; background:white;">🔍</button>
                        </div>
                    </div>
                </div>

                <div class="oa-p-group"><label>美博報價編號</label><input type="text" id="in-f-reportNo" placeholder="例如: MS/Q1241/24/kp"></div>

                <div class="oa-p-group"><label>項目內容</label><textarea id="in-f-projectContent" rows="3" placeholder="項目內容"></textarea></div>

                <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
                    <div class="oa-p-group"><label>合約總價</label><input type="text" id="in-f-total"></div>
                    <div class="oa-p-group">
                        <label>報價形式</label>
                        <select id="in-f-quoteType">
                            <option value="">請選擇</option>
                            <option value="報價邀請">報價邀請</option>
                            <option value="招標">招標</option>
                            <option value="特殊情況 (緊急)">特殊情況 (緊急)</option>
                            <option value="續約">續約</option>
                        </select>
                    </div>
                </div>

                <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
                    <div class="oa-p-group">
                        <label>採購類別</label>
                        <select id="in-f-buyType">
                            <option value="">請選擇</option>
                            <option value="保養維修材料">保養維修材料</option>
                            <option value="保養合約分判">保養合約分判</option>
                            <option value="工程材料">工程材料</option>
                            <option value="分判工程">分判工程</option>
                            <option value="後加工程">後加工程</option>
                            <option value="固定資產">固定資產</option>
                            <option value="其他">其他</option>
                        </select>
                    </div>
                    <div class="oa-p-group"><label>立項預算金額</label><input type="text" id="in-f-budget"></div>
                </div>

                <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
                    <div class="oa-p-group">
                        <label>幣種</label>
                        <select id="in-f-currency">
                            <option value="0">HKD</option>
                            <option value="1">RMB</option>
                            <option value="2">USD</option>
                            <option value="3">MOP</option>
                        </select>
                    </div>
                </div>

                <div class="oa-section-title">📦 報價明細</div>
                <div id="oa-detail-list" style="margin-bottom:10px;"></div>
                <button id="btn-add-detail" style="width:100%; padding:8px; border:1px dashed #764ba2; color:#764ba2; background:none; border-radius:10px; cursor:pointer; margin-bottom:15px; font-weight:600;">+ 添加明細</button>

                <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
                    <div class="oa-p-group"><label>邀請公司 (間)</label><input type="text" id="in-f-inviteCount"></div>
                    <div class="oa-p-group"><label>有效報價/回復 (份)</label><input type="text" id="in-f-replyCount"></div>
                </div>

                <div class="oa-p-group"><label>理由</label><textarea id="in-f-reason" rows="2"></textarea></div>
                <div class="oa-p-group"><label>中標公司</label><input type="text" id="in-f-winnerName"></div>

                <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
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

                <button id="btn-f-save" class="oa-p-btn-action">儲存項目</button>
                <button id="btn-f-save-as-new" class="oa-p-btn-action hidden" style="background:none; border:1px solid #764ba2; color:#764ba2; margin-top:10px;">另存為新項目</button>
                <div id="btn-f-cancel" style="text-align:center; margin-top:10px; font-size:12px; color:#999; cursor:pointer;" class="hidden">取消編輯</div>

                <div style="margin-top:20px; border-top:1px solid #eee; padding-top:10px; display:flex; gap:8px;">
                    <button id="btn-f-export" style="flex:1; font-size:11px; padding:6px; cursor:pointer;">匯出</button>
                    <button id="btn-f-import" style="flex:1; font-size:11px; padding:6px; cursor:pointer;">匯入</button>
                    <input type="file" id="in-f-file" class="hidden" accept=".csv">
                </div>
            </div>
            
            <div style="padding:15px; font-size:10px; color:#ddd; text-align:center; background:#fafafa;">v4.6.2</div>
        `;
        document.body.appendChild(panel);

        fields.forEach(f => {
            const el = document.getElementById('in-f-' + f);
            if (el) {
                inputs[f] = el;
                if (['budget', 'total', 'amount', 'contractAmount'].includes(f)) {
                    el.addEventListener('blur', function () {
                        if (this.value && !isNaN(this.value)) this.value = parseFloat(this.value).toFixed(2);
                    });
                }
            }
        });
        bindEvents();
        initSearch();
    }

    function initSearch() {
        const input = document.getElementById('oa-search-input');
        if (input) {
            input.oninput = () => renderList(input.value.trim());
        }
    }

    function addDetailRow(data = {}) {
        const container = document.getElementById('oa-detail-list');
        if (!container) return;
        const div = document.createElement('div');
        div.className = 'oa-detail-row';
        div.style = "background:#f9f9f9; padding:8px; border-radius:10px; margin-bottom:10px; position:relative; border:1px solid #eee;";
        div.innerHTML = `
            <div style="position:absolute; right:5px; top:5px; cursor:pointer; color:#ccc;" onclick="this.parentElement.remove()">✕</div>
            <div class="oa-p-group"><label>承判商</label><input type="text" class="dt-vendor" value="${data.vendorName || ''}"></div>
            <div class="oa-p-group"><label>報價內容</label><input type="text" class="dt-content" value="${data.content || ''}"></div>
            <div style="display:grid; grid-template-columns:1fr 1fr; gap:8px;">
                <div class="oa-p-group">
                    <label>幣種</label>
                    <select class="dt-currency">
                        <option value="0" ${data.detailCurrency === '0' ? 'selected' : ''}>HKD</option>
                        <option value="1" ${data.detailCurrency === '1' ? 'selected' : ''}>RMB</option>
                        <option value="2" ${data.detailCurrency === '2' ? 'selected' : ''}>USD</option>
                        <option value="3" ${data.detailCurrency === '3' ? 'selected' : ''}>MOP</option>
                    </select>
                </div>
                <div class="oa-p-group"><label>金額</label><input type="text" class="dt-amount" value="${data.amount || ''}"></div>
            </div>
        `;
        container.appendChild(div);
        div.querySelector('.dt-amount').addEventListener('blur', function () {
            if (this.value && !isNaN(this.value)) this.value = parseFloat(this.value).toFixed(2);
        });
    }

    function injectBall() {
        if (document.getElementById('oa-float-ball')) return;

        // 全螢幕遮罩，防止滑鼠進入 iframe 丟失事件
        const dragOverlay = document.createElement('div');
        dragOverlay.id = 'oa-drag-overlay';
        dragOverlay.style.cssText = 'position:fixed; top:0; left:0; width:100vw; height:100vh; z-index:9999998; display:none; cursor:grabbing;';
        document.body.appendChild(dragOverlay);

        ball = document.createElement('div');
        ball.id = 'oa-float-ball';
        ball.innerHTML = `
            <div class="oa-ball-menu">
                <div class="oa-ball-action pos-top" title="打開/隱藏面板" id="oa-ba-panel">
                    <svg viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                </div>
                <div class="oa-ball-action pos-left" title="恢復上一步" id="oa-ba-undo">
                    <svg viewBox="0 0 24 24"><path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5"/></svg>
                </div>
                <div class="oa-ball-action pos-right" title="清空表單" id="oa-ba-clear">
                    <svg viewBox="0 0 24 24"><path d="M21 4H8l-7 8 7 8h13a2 2 0 0 0 2-2V6a2 2 0 0 0-2-2z"/><line x1="18" y1="9" x2="12" y2="15"/><line x1="12" y1="9" x2="18" y2="15"/></svg>
                </div>
                <div class="oa-ball-action pos-bottom" title="寫入最新項目" id="oa-ba-fill">
                    <svg viewBox="0 0 24 24"><path d="M21 16.039c-.532-.054-1.071-.168-1.574-.344a6 6 0 0 1-3.082-3.082c-.176-.503-.29-1.042-.344-1.574-.054.532-.168 1.071-.344 1.574a6 6 0 0 1-3.082 3.082c-.503.176-1.042.29-1.574.344.532.054 1.071.168 1.574.344a6 6 0 0 1 3.082 3.082c.176.503.29 1.042.344 1.574.054-.532.168-1.071.344-1.574a6 6 0 0 1 3.082-3.082c.503-.176 1.042-.29 1.574-.344Z"/><path d="M9 7.02c-.355-.036-.714-.112-1.05-.23a4 4 0 0 1-2.054-2.055C5.776 4.399 5.7 4.04 5.665 3.685c-.035.355-.111.714-.23 1.05a4 4 0 0 1-2.054 2.054c-.336.118-.695.194-1.05.23.355.036.714.112 1.05.23a4 4 0 0 1 2.054 2.055c.118.336.195.695.23 1.05.036-.355.112-.714.23-1.05a4 4 0 0 1 2.055-2.054c.335-.118.694-.194 1.05-.23Z"/></svg>
                </div>
            </div>
            <div class="oa-ball-main"><span>un</span></div>
        `;
        document.body.appendChild(ball);

        let isDragging = false;
        let startX, startY, initialX, initialY;
        const mainBtn = ball.querySelector('.oa-ball-main');

        mainBtn.addEventListener('mousedown', dragStart);

        function dragStart(e) {
            initialX = ball.offsetLeft;
            initialY = ball.offsetTop;
            startX = e.clientX;
            startY = e.clientY;
            isDragging = false;

            ball.style.right = 'auto';
            ball.style.bottom = 'auto';
            ball.style.left = initialX + 'px';
            ball.style.top = initialY + 'px';
            ball.style.transition = 'none';

            dragOverlay.style.display = 'block';
            ball.classList.add('dragging'); // 避免 css hover 觸發懸停選單

            document.addEventListener('mousemove', drag);
            document.addEventListener('mouseup', dragEnd);
        }

        function drag(e) {
            let dx = e.clientX - startX;
            let dy = e.clientY - startY;
            if (Math.abs(dx) > 3 || Math.abs(dy) > 3) {
                isDragging = true;
            }
            if (!isDragging) return;

            ball.style.left = (initialX + dx) + 'px';
            ball.style.top = (initialY + dy) + 'px';
        }

        function dragEnd(e) {
            dragOverlay.style.display = 'none';
            document.removeEventListener('mousemove', drag);
            document.removeEventListener('mouseup', dragEnd);
            ball.classList.remove('dragging');

            if (isDragging) {
                // 加入彈性過渡動畫 (YouMind Style spring effect)
                ball.style.transition = 'left 0.4s cubic-bezier(0.34, 1.56, 0.64, 1), top 0.4s cubic-bezier(0.34, 1.56, 0.64, 1)';

                // 動態吸附到右側邊緣，並確保上下不超出視窗
                let padding = 20;
                let ballWidth = 48;
                let rightLimit = window.innerWidth - ballWidth - padding;

                let currentTop = parseInt(ball.style.top) || initialY;
                if (currentTop < padding) currentTop = padding;
                if (currentTop > window.innerHeight - ballWidth - padding) currentTop = window.innerHeight - ballWidth - padding;

                // 强制靠右
                ball.style.left = rightLimit + 'px';
                ball.style.top = currentTop + 'px';

                setTimeout(() => {
                    ball.style.transition = 'none';
                    ball.style.left = 'auto';
                    ball.style.right = padding + 'px';
                }, 400);

                setTimeout(() => { isDragging = false; }, 50);
            } else {
                ball.style.transition = 'all 0.3s';
                // 未拖曳視為點擊，直接開關面板
                panel.classList.toggle('open');
                if (panel.classList.contains('open')) syncData();
            }
        }

        // --- Action Buttons ---
        document.getElementById('oa-ba-panel').onclick = (e) => {
            e.stopPropagation();
            panel.classList.toggle('open');
            if (panel.classList.contains('open')) syncData();
        };

        document.getElementById('oa-ba-undo').onclick = (e) => {
            e.stopPropagation();
            if (document.getElementById('oa-btn-undo')) document.getElementById('oa-btn-undo').click();
        };

        document.getElementById('oa-ba-clear').onclick = (e) => {
            e.stopPropagation();
            if (document.getElementById('oa-btn-clear')) document.getElementById('oa-btn-clear').click();
        };

        document.getElementById('oa-ba-fill').onclick = (e) => {
            e.stopPropagation();
            if (projects && projects.length > 0) {
                let latest = projects[projects.length - 1];
                let msg = { type: "OA_LINK_FILL", project: JSON.parse(JSON.stringify(latest)) };
                window.postMessage(msg, "*");
                const bc = (win) => { for (let i = 0; i < win.frames.length; i++) { try { win.frames[i].postMessage(msg, "*"); bc(win.frames[i]); } catch (e) { } } };
                bc(window);
                alert("✨ 已為您自動填寫列表中最新的項目：" + latest.label);
            } else {
                alert("⚠️ 尚未保存任何項目圖紙");
            }
        };
    }

    function bindEvents() {
        document.getElementById('oa-btn-refresh').onclick = syncData;
        document.getElementById('oa-panel-close').onclick = () => panel.classList.remove('open');

        // 快速填充工具列
        document.getElementById('oa-btn-undo').onclick = () => {
            if (lastBackup) {
                const msg = { type: "OA_LINK_FILL", project: { label: "恢復備份", details: [], ...lastBackup } };
                window.postMessage(msg, "*");
                const bc = (win) => { for (let i = 0; i < win.frames.length; i++) { try { win.frames[i].postMessage(msg, "*"); bc(win.frames[i]); } catch (e) { } } };
                bc(window);
            } else {
                alert("⚠️ 尚無可恢復的備份");
            }
        };
        document.getElementById('oa-btn-clear').onclick = () => {
            window.postMessage({ type: "OA_LINK_CLEAR" }, "*");
            const bc = (win) => { for (let i = 0; i < win.frames.length; i++) { try { win.frames[i].postMessage({ type: "OA_LINK_CLEAR" }, "*"); bc(win.frames[i]); } catch (e) { } } };
            bc(window);
        };

        // 快捷工具列事件
        document.getElementById('oa-btn-add').onclick = () => { resetForm(); tabManage.click(); };
        document.getElementById('oa-btn-save').onclick = () => document.getElementById('btn-f-save').click();
        document.getElementById('oa-btn-export').onclick = () => document.getElementById('btn-f-export').click();
        document.getElementById('oa-btn-import').onclick = () => document.getElementById('btn-f-import').click();
        const tabFill = document.getElementById('tab-v-fill'), tabManage = document.getElementById('tab-v-manage');
        const viewFill = document.getElementById('view-v-fill'), viewManage = document.getElementById('view-v-manage');
        tabFill.onclick = () => { tabFill.classList.add('active'); tabManage.classList.remove('active'); viewFill.classList.remove('hidden'); viewManage.classList.add('hidden'); };
        tabManage.onclick = () => { tabManage.classList.add('active'); tabFill.classList.remove('active'); viewManage.classList.remove('hidden'); viewFill.classList.add('hidden'); };

        document.getElementById('btn-f-pick').onclick = () => {
            function findAndClickInFrames(win) {
                try {
                    const btn = win.document.getElementById('field1366309_browserbtn');
                    if (btn) { btn.click(); return true; }
                    for (let i = 0; i < win.frames.length; i++) { if (findAndClickInFrames(win.frames[i])) return true; }
                } catch (e) { } return false;
            }
            if (!findAndClickInFrames(window)) alert("⚠️ 找不到人員選擇按鈕");
        };

        document.getElementById('btn-add-detail').onclick = () => addDetailRow();

        function exportToCSV() {
            const h = ["項目標題", "物業名稱", "報價編號", "項目經理", "項目內容", "立項預算", "合約總價", "報價形式", "採購類別", "幣種", "承判商", "報價內容", "明細幣種", "金額", "邀請公司", "有效報價", "推荐理由", "中標公司", "合約幣種", "合約金額"];
            const rows = [];
            projects.forEach(p => {
                (p.details || [{}]).forEach(dt => {
                    const row = [p.label, p.propertyName, p.reportNo, p.manager, p.projectContent, p.budget, p.total, p.quoteType, p.buyType, p.currency, dt.vendorName, dt.content, dt.detailCurrency, dt.amount, p.inviteCount, p.replyCount, p.reason, p.winnerName, p.contractCurrency, p.contractAmount];
                    rows.push(row.map(v => `"${(String(v || '')).replace(/"/g, '""')}"`).join(","));
                });
            });
            const b = new Blob(["\uFEFF" + h.join(",") + "\n" + rows.join("\n")], { type: 'text/csv;charset=utf-8' });
            const a = document.createElement('a'); a.href = URL.createObjectURL(b); a.download = `OA_Projects_${Date.now()}.csv`; a.click();
        }

        document.getElementById('btn-f-save').onclick = async () => {
            const id = document.getElementById('in-f-id').value || "p_" + Date.now();
            const p = { id: id, managerId: document.getElementById('in-f-managerId').value, details: [] };
            fields.forEach(f => { if (inputs[f]) p[f] = inputs[f].value.trim(); });
            document.querySelectorAll('.oa-detail-row').forEach(row => {
                p.details.push({
                    vendorName: row.querySelector('.dt-vendor').value.trim(),
                    content: row.querySelector('.dt-content').value.trim(),
                    detailCurrency: row.querySelector('.dt-currency').value,
                    amount: row.querySelector('.dt-amount').value.trim()
                });
            });
            if (!p.label || p.details.length === 0) return alert("必填項缺失！");
            const idx = projects.findIndex(x => x && x.id === id);
            if (idx > -1) projects[idx] = p; else projects.push(p);
            await chrome.storage.local.set({ 'oa_projects_v13': projects });
            alert("✅ 保存成功及自動匯出"); exportToCSV(); syncData(); tabFill.click(); resetForm();
        };

        document.getElementById('btn-f-save-as-new').onclick = async () => {
            const p = { id: "p_" + Date.now(), managerId: document.getElementById('in-f-managerId').value, details: [] };
            fields.forEach(f => { if (inputs[f]) p[f] = inputs[f].value.trim(); });
            document.querySelectorAll('.oa-detail-row').forEach(row => {
                p.details.push({
                    vendorName: row.querySelector('.dt-vendor').value.trim(),
                    content: row.querySelector('.dt-content').value.trim(),
                    detailCurrency: row.querySelector('.dt-currency').value,
                    amount: row.querySelector('.dt-amount').value.trim()
                });
            });
            if (!p.label || p.details.length === 0) return alert("必填項缺失！");
            projects.push(p);
            await chrome.storage.local.set({ 'oa_projects_v13': projects });
            alert("✅ 已另存為新項目及自動匯出"); exportToCSV(); syncData(); tabFill.click(); resetForm();
        };

        document.getElementById('btn-f-cancel').onclick = () => { resetForm(); tabFill.click(); };
        document.getElementById('btn-f-export').onclick = () => exportToCSV();

        const fileIn = document.getElementById('in-f-file');
        document.getElementById('btn-f-import').onclick = () => fileIn.click();
        fileIn.onchange = (e) => {
            if (!e.target.files.length) return;
            const reader = new FileReader();
            reader.onload = async (ev) => {
                const content = ev.target.result;
                const lines = content.split(/\r?\n/).filter(l => l.trim());
                if (lines.length < 2) return;

                const grouped = {};
                lines.slice(1).forEach(l => {
                    const cols = [];
                    let c = "", q = false;
                    for (let i = 0; i < l.length; i++) {
                        if (l[i] === '"') q = !q;
                        else if (l[i] === ',' && !q) { cols.push(c); c = ""; }
                        else c += l[i];
                    }
                    cols.push(c);

                    // 使用 項目標題 + 報價編號 作為 Key 進行分組
                    const label = (cols[0] || "").trim();
                    const reportNo = (cols[2] || "").trim();
                    if (!label && !reportNo) return;

                    const k = label + "_" + reportNo;
                    if (!grouped[k]) {
                        grouped[k] = {
                            id: "p_" + Date.now() + "_" + Math.floor(Math.random() * 1000),
                            label: label,
                            propertyName: (cols[1] || "").trim(),
                            reportNo: reportNo,
                            manager: (cols[3] || "").trim(),
                            projectContent: (cols[4] || "").trim(),
                            budget: (cols[5] || "").trim(),
                            total: (cols[6] || "").trim(),
                            quoteType: (cols[7] || "").trim(),
                            buyType: (cols[8] || "").trim(),
                            currency: (cols[9] || "0").trim(),
                            details: [],
                            inviteCount: (cols[14] || "").trim(),
                            replyCount: (cols[15] || "").trim(),
                            reason: (cols[16] || "").trim(),
                            winnerName: (cols[17] || "").trim(),
                            contractCurrency: (cols[18] || "0").trim(),
                            contractAmount: (cols[19] || "").trim()
                        };
                    }

                    // 添加明細（承判商、內容、幣種、金額）
                    if (cols[10] || cols[11] || cols[13]) {
                        grouped[k].details.push({
                            vendorName: (cols[10] || "").trim(),
                            content: (cols[11] || "").trim(),
                            detailCurrency: (cols[12] || "0").trim(),
                            amount: (cols[13] || "").trim()
                        });
                    }
                });

                const imported = Object.values(grouped);
                if (imported.length > 0) {
                    projects = imported;
                    await chrome.storage.local.set({ 'oa_projects_v13': projects });
                    alert(`✅ 成功匯入 ${imported.length} 個項目`);
                    syncData();
                } else {
                    alert("⚠️ 未發現有效數據");
                }
            };
            reader.readAsText(e.target.files[0]);
            fileIn.value = ""; // 重置 input 以便下次選擇同一個文件
        };
    }

    async function syncData() {
        const res = await chrome.storage.local.get(['oa_projects_v13']);
        projects = res.oa_projects_v13 || [];
        const searchInput = document.getElementById('oa-search-input');
        renderList(searchInput ? searchInput.value.trim() : "");
    }

    function renderList(query = "") {
        const list = document.getElementById('oa-disp-list'); if (!list) return;

        let filtered = projects;
        if (query) {
            const q = query.toLowerCase();
            filtered = projects.filter(p =>
                (p.label && p.label.toLowerCase().includes(q)) ||
                (p.manager && p.manager.toLowerCase().includes(q)) ||
                (p.vendorName && p.vendorName.toLowerCase().includes(q))
            );
        }

        list.innerHTML = filtered.length === 0 ? `<p style="padding:40px; color:#999; text-align:center;">${query ? '找不到匹配項目' : '尚無記錄'}</p>` : '';
        filtered.forEach(p => {
            const div = document.createElement('div'); div.className = 'oa-p-item';
            div.innerHTML = `
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <div style="flex:1;"><strong>${p.label}</strong></div>
                    <div style="display:flex; gap:8px;">
                        <div class="edit-btn" style="cursor:pointer; padding:2px 5px;">✎</div>
                        <div class="del-btn" style="cursor:pointer; padding:2px 5px;">🗑️</div>
                    </div>
                </div>`;
            div.querySelector('div[style*="flex:1"]').onclick = () => { console.log("Filling Project:", p); runFill(p); };
            div.querySelector('.edit-btn').onclick = (e) => {
                e.stopPropagation(); document.getElementById('in-f-id').value = p.id; document.getElementById('in-f-managerId').value = p.managerId || "";
                fields.forEach(f => { if (inputs[f]) inputs[f].value = p[f] || ""; });
                const dList = document.getElementById('oa-detail-list'); dList.innerHTML = "";
                if (p.details) p.details.forEach(dt => addDetailRow(dt));
                document.getElementById('tab-v-manage').click(); document.getElementById('btn-f-save').textContent = "更新項目";
                document.getElementById('btn-f-cancel').classList.remove('hidden');
                document.getElementById('btn-f-save-as-new').classList.remove('hidden');
                syncManagerFeedback(p.manager);
            };
            div.querySelector('.del-btn').onclick = async (e) => {
                e.stopPropagation();
                if (confirm(`🗑️ 確定刪除「${p.label}」嗎？`)) {
                    projects = projects.filter(x => x && x.id !== p.id);
                    await chrome.storage.local.set({ 'oa_projects_v13': projects });
                    syncData();
                }
            };
            list.appendChild(div);
        });
    }

    function runFill(p) {
        // 先備份當前數據
        const fieldIds = ['field1366311', 'field1366312', 'field1366309', 'field1366275', 'field1366310', 'field1366280', 'field1366276', 'field1366278', 'field1366279', 'field1366290', 'field1366291', 'field1366292', 'field1366313', 'field1366294', 'field1366295'];
        const bak = {};
        fieldIds.forEach(id => { const e = document.getElementById(id) || document.querySelector(`[name="${id}"]`); if (e) bak[id] = e.value; });
        lastBackup = bak;

        const msg = { type: "OA_LINK_FILL", project: JSON.parse(JSON.stringify(p)) };
        window.postMessage(msg, "*");
        const bc = (win) => { for (let i = 0; i < win.frames.length; i++) { try { win.frames[i].postMessage(msg, "*"); bc(win.frames[i]); } catch (e) { } } };
        bc(window);
    }

    function resetForm() {
        document.getElementById('in-f-id').value = ""; document.getElementById('in-f-managerId').value = "";
        fields.forEach(f => { if (inputs[f]) inputs[f].value = (f === 'currency' || f === 'contractCurrency') ? '0' : ''; });
        document.getElementById('oa-detail-list').innerHTML = ""; addDetailRow();
        document.getElementById('btn-f-save').textContent = "儲存項目";
        document.getElementById('btn-f-cancel').classList.add('hidden');
        document.getElementById('btn-f-save-as-new').classList.add('hidden');
        document.getElementById('in-f-manager').style.backgroundColor = "";
    }

    async function syncManagerFeedback(name) {
        if (!name) return;
        const res = await chrome.storage.local.get(['oa_manager_cache']);
        const cache = res.oa_manager_cache || {};
        if (cache[name.trim()]) {
            document.getElementById('in-f-manager').style.backgroundColor = "#e8f5e9";
            document.getElementById('in-f-managerId').value = cache[name.trim()];
        } else {
            document.getElementById('in-f-manager').style.backgroundColor = "";
        }
    }

    function startPickerWatcher() {
        // 優化：使用 MutationObserver 監測 DOM 變化，減少頻繁輪詢的 CPU 消耗
        const observeFrame = (win) => {
            try {
                const target = win.document.getElementById('field1366309span');
                if (!target) return;

                const observer = new MutationObserver(async () => {
                    const valEl = win.document.getElementById('field1366309');
                    const spanEl = win.document.getElementById('field1366309span');
                    if (valEl && valEl.value && spanEl && spanEl.innerText.trim()) {
                        const name = spanEl.innerText.trim();
                        const id = valEl.value;
                        const nameIn = document.getElementById('in-f-manager');
                        const idIn = document.getElementById('in-f-managerId');

                        if (nameIn && nameIn.value !== name) {
                            nameIn.value = name;
                            idIn.value = id;
                            const res = await chrome.storage.local.get(['oa_manager_cache']);
                            const cache = res.oa_manager_cache || {};
                            cache[name] = id;
                            await chrome.storage.local.set({ 'oa_manager_cache': cache });
                            nameIn.style.backgroundColor = "#e8f5e9";
                        }
                    }
                });
                observer.observe(target, { childList: true, subtree: true, characterData: true });
            } catch (e) { }
        };

        // 初始觀察所有 Frame
        const scanFrames = (win) => {
            observeFrame(win);
            for (let i = 0; i < win.frames.length; i++) {
                try { scanFrames(win.frames[i]); } catch (e) { }
            }
        };
        scanFrames(window);

        // 每隔 5 秒檢查是否有新 Frame 產生（較低頻率）
        setInterval(() => scanFrames(window), 5000);

        const mgrIn = document.getElementById('in-f-manager');
        if (mgrIn) {
            mgrIn.addEventListener('input', async function () {
                const name = this.value.trim(), idIn = document.getElementById('in-f-managerId');
                if (!name || !idIn) { this.style.backgroundColor = ""; return; }
                const res = await chrome.storage.local.get(['oa_manager_cache']);
                const cache = res.oa_manager_cache || {};
                if (cache[name]) {
                    idIn.value = cache[name]; console.log("OA Bubble: Matched ->", cache[name]);
                    this.style.backgroundColor = "#e8f5e9";
                } else {
                    this.style.backgroundColor = ""; idIn.value = "";
                }
            });
        }
    }
    init();
})();
