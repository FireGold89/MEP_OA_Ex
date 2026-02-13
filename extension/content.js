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
        syncData();
        startPickerWatcher();
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
                        <span>Projects</span>
                        <span style="font-size:10px; opacity:0.5;">▼</span>
                    </div>
                </div>
                <div class="oa-header-right">
                    <span class="oa-header-icon" title="更多設置">···</span>
                    <span class="oa-header-icon" id="oa-btn-refresh" title="同步數據">↻</span>
                    <span class="oa-header-icon close" id="oa-panel-close" title="隱藏視窗">✕</span>
                </div>
            </div>

            <div class="oa-domain-card">
                <div class="oa-domain-info">
                    <div class="oa-domain-icon">H</div>
                    <span>oa.copm.com.cn</span>
                </div>
                <button class="oa-btn-save-top" id="btn-quick-save">保存</button>
            </div>

            <div class="oa-p-tabs">
                <div class="oa-p-tab active" id="tab-v-fill">快速填充</div>
                <div class="oa-p-tab" id="tab-v-manage">項目管理</div>
            </div>

            <div id="view-v-fill" class="oa-panel-body">
                <div class="oa-section-title">📖 填充列表</div>
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
        ball = document.createElement('div');
        ball.id = 'oa-float-ball';
        ball.innerHTML = `<span>un</span>`;
        document.body.appendChild(ball);
        ball.onclick = () => { panel.classList.toggle('open'); if (panel.classList.contains('open')) syncData(); };
    }

    function bindEvents() {
        document.getElementById('oa-btn-refresh').onclick = syncData;
        document.getElementById('oa-panel-close').onclick = () => panel.classList.remove('open');
        document.getElementById('btn-quick-save').onclick = () => document.getElementById('btn-f-save').click();
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
            alert("✅ 保存成功"); syncData(); tabFill.click(); resetForm();
        };

        document.getElementById('btn-f-cancel').onclick = () => { resetForm(); tabFill.click(); };
        document.getElementById('btn-f-export').onclick = () => {
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
        };

        const fileIn = document.getElementById('in-f-file');
        document.getElementById('btn-f-import').onclick = () => fileIn.click();
        fileIn.onchange = (e) => {
            const reader = new FileReader(); reader.onload = async (ev) => {
                const lines = ev.target.result.split(/\r?\n/).filter(l => l.trim());
                const grouped = {};
                lines.slice(1).forEach(l => {
                    const cols = []; let c = "", q = false;
                    for (let i = 0; i < l.length; i++) { if (l[i] === '"') q = !q; else if (l[i] === ',' && !q) { cols.push(c); c = ""; } else c += l[i]; } cols.push(c);
                    const k = cols[0] + "_" + cols[2];
                    if (!grouped[k]) grouped[k] = { id: "p_" + Date.now() + Math.random(), label: cols[0], propertyName: cols[1], reportNo: cols[2], manager: cols[3], projectContent: cols[4], budget: cols[5], total: cols[6], quoteType: cols[7], buyType: cols[8], currency: cols[9], details: [], inviteCount: cols[14], replyCount: cols[15], reason: cols[16], winnerName: cols[17], contractCurrency: cols[18], contractAmount: cols[19] };
                    if (cols[10]) grouped[k].details.push({ vendorName: cols[10], content: cols[11], detailCurrency: cols[12], amount: cols[13] });
                });
                projects = Object.values(grouped); await chrome.storage.local.set({ 'oa_projects_v13': projects }); syncData();
            };
            reader.readAsText(e.target.files[0]);
        };
    }

    async function syncData() { const res = await chrome.storage.local.get(['oa_projects_v13']); projects = res.oa_projects_v13 || []; renderList(); }

    function renderList() {
        const list = document.getElementById('oa-disp-list'); if (!list) return;
        list.innerHTML = projects.length === 0 ? '<p style="padding:40px;">尚無記錄</p>' : '';
        projects.forEach(p => {
            const div = document.createElement('div'); div.className = 'oa-p-item';
            div.innerHTML = `<div style="display:flex; justify-content:space-between;"><div><strong>${p.label}</strong><br><small>${p.manager || ''}</small></div><div class="edit-btn">✎</div></div>`;
            div.onclick = () => { console.log("Filling Project:", p); runFill(p); };
            div.querySelector('.edit-btn').onclick = (e) => {
                e.stopPropagation(); document.getElementById('in-f-id').value = p.id; document.getElementById('in-f-managerId').value = p.managerId || "";
                fields.forEach(f => { if (inputs[f]) inputs[f].value = p[f] || ""; });
                const dList = document.getElementById('oa-detail-list'); dList.innerHTML = "";
                if (p.details) p.details.forEach(dt => addDetailRow(dt));
                document.getElementById('tab-v-manage').click(); document.getElementById('btn-f-save').textContent = "更新項目"; document.getElementById('btn-f-cancel').classList.remove('hidden');
                syncManagerFeedback(p.manager);
            };
            list.appendChild(div);
        });
    }

    function runFill(p) {
        const msg = { type: "OA_LINK_FILL", project: JSON.parse(JSON.stringify(p)) };
        window.postMessage(msg, "*");
        const bc = (win) => { for (let i = 0; i < win.frames.length; i++) { try { win.frames[i].postMessage(msg, "*"); bc(win.frames[i]); } catch (e) { } } };
        bc(window);
    }

    function resetForm() {
        document.getElementById('in-f-id').value = ""; document.getElementById('in-f-managerId').value = "";
        fields.forEach(f => { if (inputs[f]) inputs[f].value = (f === 'currency' || f === 'contractCurrency') ? '0' : ''; });
        document.getElementById('oa-detail-list').innerHTML = ""; addDetailRow();
        document.getElementById('btn-f-save').textContent = "儲存項目"; document.getElementById('btn-f-cancel').classList.add('hidden');
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
        setInterval(() => {
            async function checkPicker(win) {
                try {
                    const val = win.document.getElementById('field1366309'), span = win.document.getElementById('field1366309span');
                    if (val && val.value && span && span.innerText.trim()) {
                        const name = span.innerText.trim(), id = val.value;
                        const nameIn = document.getElementById('in-f-manager'), idIn = document.getElementById('in-f-managerId');
                        if (nameIn && nameIn.value !== name) {
                            nameIn.value = name; idIn.value = id; console.log("OA Bubble: Manager Cached ->", name, id);
                            const res = await chrome.storage.local.get(['oa_manager_cache']);
                            const cache = res.oa_manager_cache || {}; cache[name] = id;
                            await chrome.storage.local.set({ 'oa_manager_cache': cache });
                            nameIn.style.backgroundColor = "#e8f5e9";
                            return true;
                        }
                    }
                } catch (e) { } return false;
            }
            if (checkPicker(window)) return;
            for (let i = 0; i < window.frames.length; i++) { try { if (checkPicker(window.frames[i])) return; } catch (e) { } }
        }, 1200);

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
