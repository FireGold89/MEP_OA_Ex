/**
 * Version: 4.12.4 - Browser Field UI Fix
 * 與懸浮面板 100% 統一版本的 popup.js
 */

document.addEventListener('DOMContentLoaded', async function () {
    let projects = [];
    const fields = ['label', 'propertyName', 'reportNo', 'manager', 'projectContent', 'budget', 'total', 'quoteType', 'buyType', 'currency', 'inviteCount', 'replyCount', 'reason', 'winnerName', 'contractCurrency', 'contractAmount'];
    const inputs = {};
    let lastBackup = null;

    const tabFill = document.getElementById('tab-v-fill'), tabManage = document.getElementById('tab-v-manage');
    const viewFill = document.getElementById('view-v-fill'), viewManage = document.getElementById('view-v-manage');

    function init() {
        fields.forEach(f => {
            const el = document.getElementById('in-f-' + f);
            if (el) {
                inputs[f] = el;
                if (['budget', 'total', 'contractAmount'].includes(f)) {
                    el.addEventListener('blur', function () {
                        if (this.value && !isNaN(this.value)) this.value = parseFloat(this.value).toFixed(2);
                    });
                }
            }
        });
        bindEvents();
        initSearch();
        syncData();
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

    function bindEvents() {
        document.getElementById('oa-btn-refresh').onclick = syncData;

        // 快速填充工具列
        document.getElementById('oa-btn-undo').onclick = () => {
            if (lastBackup) {
                const msg = { type: "OA_LINK_FILL", project: { label: "恢復備份", details: [], ...lastBackup } };
                sendMessageToActiveTab(msg);
            } else {
                alert("⚠️ 尚無可恢復的備份");
            }
        };
        document.getElementById('oa-btn-clear').onclick = () => {
            sendMessageToActiveTab({ type: "OA_LINK_CLEAR" });
        };

        // 快捷工具列事件
        document.getElementById('oa-btn-add').onclick = () => { resetForm(); tabManage.click(); };
        document.getElementById('oa-btn-save').onclick = () => document.getElementById('btn-f-save').click();
        document.getElementById('oa-btn-export').onclick = () => document.getElementById('btn-f-export').click();
        document.getElementById('oa-btn-import').onclick = () => document.getElementById('btn-f-import').click();

        const tabFill = document.getElementById('tab-v-fill'), tabManage = document.getElementById('tab-v-manage'), tabSettings = document.getElementById('tab-v-settings');
        const viewFill = document.getElementById('view-v-fill'), viewManage = document.getElementById('view-v-manage'), viewSettings = document.getElementById('view-v-settings');
        
        tabFill.onclick = () => { tabFill.classList.add('active'); tabManage.classList.remove('active'); tabSettings.classList.remove('active'); viewFill.classList.remove('hidden'); viewManage.classList.add('hidden'); viewSettings.classList.add('hidden'); };
        tabManage.onclick = () => { tabManage.classList.add('active'); tabFill.classList.remove('active'); tabSettings.classList.remove('active'); viewManage.classList.remove('hidden'); viewFill.classList.add('hidden'); viewSettings.classList.add('hidden'); };
        tabSettings.onclick = () => { tabSettings.classList.add('active'); tabFill.classList.remove('active'); tabManage.classList.remove('active'); viewSettings.classList.remove('hidden'); viewFill.classList.add('hidden'); viewManage.classList.add('hidden'); };

        // ===== 懸浮輔助球開關 =====
        const checkBall = document.getElementById('oa-check-ball');
        if (checkBall) {
            chrome.storage.local.get(['oa_show_float_ball'], (res) => {
                checkBall.checked = (res.oa_show_float_ball !== false);
            });
            checkBall.onchange = () => {
                const visible = checkBall.checked;
                chrome.storage.local.set({ 'oa_show_float_ball': visible });
                sendMessageToActiveTab({ type: "TOGGLE_FLOAT_BALL", visible: visible });
            };
        }

        // ===== 標題字體大小滑桿 =====
        const titleSizeSlider = document.getElementById('set-title-size');
        const titleSizeVal = document.getElementById('set-title-size-val');
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
                // 實時宣傳到活躍分頁
                sendMessageToActiveTab({ type: "UPDATE_SETTING", key: "title_size", value: parseInt(v) });
                // 實時重新渲染本地列表
                const q = document.getElementById('oa-search-input')?.value.trim() || "";
                renderList(q);
            };
        }

        // ===== 副標題顯示模式 =====
        const subtitleCtrl = document.getElementById('set-subtitle-mode');
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
                    sendMessageToActiveTab({ type: "UPDATE_SETTING", key: "subtitle_mode", value: btn.dataset.value });
                    // 實時重新渲染本地列表
                    const q = document.getElementById('oa-search-input')?.value.trim() || "";
                    renderList(q);
                };
            });
        }

        // ===== 緊湊模式 =====
        const compactCheck = document.getElementById('set-compact-mode');
        if (compactCheck) {
            chrome.storage.local.get(['oa_setting_compact'], (res) => {
                compactCheck.checked = !!res.oa_setting_compact;
            });
            compactCheck.onchange = () => {
                chrome.storage.local.set({ 'oa_setting_compact': compactCheck.checked });
                sendMessageToActiveTab({ type: "UPDATE_SETTING", key: "compact", value: compactCheck.checked });
                // 實時重新渲染本地列表
                const q = document.getElementById('oa-search-input')?.value.trim() || "";
                renderList(q);
            };
        }

        // ===== 匯出確認開關 =====
        const exportConfirmCheck = document.getElementById('set-export-confirm');
        if (exportConfirmCheck) {
            chrome.storage.local.get(['oa_setting_export_confirm'], (res) => {
                exportConfirmCheck.checked = (res.oa_setting_export_confirm !== false);
            });
            exportConfirmCheck.onchange = () => {
                chrome.storage.local.set({ 'oa_setting_export_confirm': exportConfirmCheck.checked });
            };
        }

        document.getElementById('btn-f-pick').onclick = () => {
            alert("⚠️ 彈窗介面無法直接讀取頁面中的人員選擇器，請在網頁上的 OA 懸浮面板使用此功能，或手動輸入經理名稱。");
        };

        document.getElementById('btn-add-detail').onclick = () => addDetailRow();

        function buildCsvDataUrl() {
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
            chrome.downloads.download({
                url: dataUrl,
                filename: filename || 'OA_Projects.csv',
                conflictAction: conflictAction || 'overwrite',
                saveAs: false
            }, () => {});
        }

        async function exportToCSV() {
            if (projects.length === 0) return alert("尚無數據可匯出");

            const setting = await chrome.storage.local.get(['oa_setting_export_confirm']);
            const needsConfirm = (setting.oa_setting_export_confirm !== false);

            if (!needsConfirm) {
                triggerDownload('OA_Projects.csv', 'overwrite');
                return;
            }

            // 創建確認對話框
            const existingModal = document.getElementById('oa-export-modal-popup');
            if (existingModal) existingModal.remove();

            const modal = document.createElement('div');
            modal.id = 'oa-export-modal-popup';
            modal.style.cssText = `
                position:fixed; inset:0; z-index:9999;
                display:flex; align-items:center; justify-content:center;
                background:rgba(0,0,0,0.45);
            `;
            modal.innerHTML = `
                <div style="background:#fff; border-radius:16px; padding:24px 28px; width:90%;
                            box-shadow:0 10px 40px rgba(0,0,0,0.2); font-family:-apple-system,sans-serif;">
                    <div style="font-size:20px; text-align:center; margin-bottom:6px">📤</div>
                    <div style="font-weight:700; font-size:15px; text-align:center; margin-bottom:6px; color:#111">匯出確認</div>
                    <div style="font-size:12px; color:#666; text-align:center; margin-bottom:18px; line-height:1.6">
                        可能已存在 <strong>OA_Projects.csv</strong><br>請選擇處理方式：
                    </div>
                    <button id="pop-exp-overwrite" style="width:100%; padding:12px; margin-bottom:8px;
                        border-radius:10px; border:none; background:#1d1d1f; color:#fff;
                        font-size:13px; font-weight:600; cursor:pointer;">✅ 覆蓋舊檔</button>
                    <button id="pop-exp-new" style="width:100%; padding:12px; margin-bottom:8px;
                        border-radius:10px; border:1px solid #ddd; background:#f7f7f7; color:#333;
                        font-size:13px; font-weight:600; cursor:pointer;">📅 另存新檔（加日期時間）</button>
                    <button id="pop-exp-cancel" style="width:100%; padding:8px;
                        border:none; background:none; color:#999; font-size:12px; cursor:pointer;">取消</button>
                </div>
            `;
            document.body.appendChild(modal);

            const close = () => modal.remove();
            modal.addEventListener('click', (e) => { if (e.target === modal) close(); });
            document.getElementById('pop-exp-cancel').onclick = close;
            document.getElementById('pop-exp-overwrite').onclick = () => { close(); triggerDownload('OA_Projects.csv', 'overwrite'); };
            document.getElementById('pop-exp-new').onclick = () => {
                close();
                const ts = new Date().toISOString().slice(0,16).replace('T','_').replace(':','-');
                triggerDownload(`OA_Projects_${ts}.csv`, 'uniquify');
            };
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

    async function renderList(query = "") {
        const list = document.getElementById('oa-disp-list'); if (!list) return;

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
            filtered = projects.filter(p =>
                (p.label && p.label.toLowerCase().includes(q)) ||
                (p.manager && p.manager.toLowerCase().includes(q)) ||
                (p.details && p.details.some(dt => dt.vendorName && dt.vendorName.toLowerCase().includes(q)))
            );
        }

        list.innerHTML = filtered.length === 0 ? `<p style="padding:40px; color:#999; text-align:center;">${query ? '找不到匹配項目' : '尚無記錄'}</p>` : '';
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
                    <div style="display:flex; gap:12px; margin-left:12px; flex-shrink:0;">
                        <div class="edit-btn" style="cursor:pointer; font-size:16px; opacity:0.3; transition:opacity 0.2s;">✎</div>
                        <div class="del-btn" style="cursor:pointer; font-size:16px; color:#ff4d4f; opacity:0.3; transition:opacity 0.2s;">🗑</div>
                    </div>
                </div>`;
            div.onclick = (e) => {
                if (e.target.classList.contains('edit-btn') || e.target.classList.contains('del-btn')) return;
                console.log("Filling Project:", p);
                runFill(p);
                const originalBg = div.style.background;
                div.style.background = '#f2f2f7';
                setTimeout(() => div.style.background = originalBg || '#fff', 200);
            };
            div.querySelector('.edit-btn').onclick = (e) => {
                e.stopPropagation(); document.getElementById('in-f-id').value = p.id; document.getElementById('in-f-managerId').value = p.managerId || "";
                fields.forEach(f => { if (inputs[f]) inputs[f].value = p[f] || ""; });
                const dList = document.getElementById('oa-detail-list'); dList.innerHTML = "";
                if (p.details) p.details.forEach(dt => addDetailRow(dt));
                document.getElementById('tab-v-manage').click(); document.getElementById('btn-f-save').textContent = "更新項目";
                document.getElementById('btn-f-cancel').classList.remove('hidden');
                document.getElementById('btn-f-save-as-new').classList.remove('hidden');
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

    // 發送給 content.js 進行實體填充
    function sendMessageToActiveTab(msg) {
        chrome.tabs.query({ active: true, currentWindow: true }, function (tabs) {
            if (tabs && tabs[0]) {
                chrome.tabs.sendMessage(tabs[0].id, msg, function (response) {
                    if (chrome.runtime.lastError) {
                        console.error('Popup message failed. Are you on a valid page?', chrome.runtime.lastError.message);
                        alert("⚠️ 填充失敗。此頁面可能不支持自動填充，請確保在正確的 OA 頁籤中。");
                    }
                });
            }
        });
    }

    function runFill(p) {
        // 先建立備份以利未來如果需要在 popup_backup 恢復 (由於 popup 環境限制，目前無法直接讀取 DOM)
        // 故在此簡化備份為直接重置
        lastBackup = null;

        const msg = { type: "OA_LINK_FILL", project: JSON.parse(JSON.stringify(p)) };
        sendMessageToActiveTab(msg);
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

    init();
});
