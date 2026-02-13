/**
 * Version: 3.9_Final_Stable
 * 核心解決：彈窗失焦關閉後的數據持久化與異步人員選取回填
 */

document.addEventListener('DOMContentLoaded', async function () {
    const statusEl = document.getElementById('status');
    const vendorListEl = document.getElementById('vendor-list');
    const toggleFloatBtn = document.getElementById('toggle-float-ball');

    // Tabs & Views
    const tabFill = document.getElementById('tab-fill'), tabManage = document.getElementById('tab-manage');
    const viewFill = document.getElementById('view-fill'), viewManage = document.getElementById('view-manage');

    // 16 欄位定義
    const fields = ['label', 'propertyName', 'reportNo', 'manager', 'projectContent', 'budget', 'total', 'quoteType', 'buyType', 'currency', 'vendorName', 'content', 'amount', 'inviteCount', 'replyCount', 'reason'];
    const inputs = {};
    fields.forEach(f => inputs[f] = document.getElementById('input-' + f));

    const editIdInput = document.getElementById('edit-id'), formTitle = document.getElementById('form-title');
    const btnSave = document.getElementById('btn-save'), btnCancel = document.getElementById('btn-cancel');
    const btnManagerPick = document.getElementById('btn-oa-manager-pick');
    const inputManagerId = document.getElementById('input-manager-id');
    const btnExport = document.getElementById('btn-export'), btnImport = document.getElementById('btn-import-trigger');
    const inputImport = document.getElementById('input-import');

    let lastBackup = null;

    // --- 1. 草稿持久化機制 ---
    async function saveDraft() {
        const draft = { id: editIdInput.value, managerId: inputManagerId.value };
        fields.forEach(f => draft[f] = inputs[f] ? inputs[f].value : "");
        chrome.storage.local.set({ 'oa_form_draft': draft });
    }

    async function loadDraft() {
        const res = await chrome.storage.local.get(['oa_form_draft', 'oa_picker_result']);

        // 優先填入草稿
        if (res.oa_form_draft) {
            const d = res.oa_form_draft;
            editIdInput.value = d.id || "";
            inputManagerId.value = d.managerId || "";
            fields.forEach(f => { if (inputs[f]) inputs[f].value = d[f] || ""; });
            if (editIdInput.value) { formTitle.textContent = "編輯項目"; btnCancel.classList.remove('hidden'); btnSave.textContent = "更新項目"; }
        }

        // 檢查並回填人員選擇結果 (來自 content.js 背景捕捉)
        if (res.oa_picker_result) {
            inputs['manager'].value = res.oa_picker_result.name;
            inputManagerId.value = res.oa_picker_result.id;
            chrome.storage.local.remove('oa_picker_result');
            saveDraft(); // 立即同步到草稿
            updateStatus("✅ 已存入選取人員: " + res.oa_picker_result.name, "#2ecc71");
        }
    }

    // 為所有輸入框綁定自動備份
    Object.values(inputs).forEach(el => { if (el) el.oninput = saveDraft; });
    loadDraft();

    // --- 2. 懸浮球與權責數據 ---
    if (toggleFloatBtn) {
        chrome.storage.local.get(['oa_show_float_ball'], (res) => { toggleFloatBtn.checked = (res.oa_show_float_ball !== false); });
        toggleFloatBtn.onchange = () => { chrome.storage.local.set({ 'oa_show_float_ball': toggleFloatBtn.checked }); updateStatus(toggleFloatBtn.checked ? "✅ 懸浮球已開啟" : "🌑 懸浮球已隱藏", "#2ecc71"); };
    }

    const DB = {
        primaryKey: 'oa_projects_v13',
        async get() {
            const res = await new Promise(r => chrome.storage.local.get(this.primaryKey, r));
            return Array.isArray(res[this.primaryKey]) ? res[this.primaryKey] : [];
        },
        async save(data) {
            const o = {}; o[this.primaryKey] = data;
            await chrome.storage.local.set(o);
            chrome.storage.local.remove('oa_form_draft'); // 儲存完畢清除草稿
            return true;
        }
    };

    let projects = await DB.get();
    renderList();

    // --- 3. UI 交互 ---
    tabFill.onclick = () => { viewFill.classList.remove('hidden'); viewManage.classList.add('hidden'); tabFill.classList.add('active'); tabManage.classList.remove('active'); resetForm(); chrome.storage.local.remove('oa_form_draft'); };
    tabManage.onclick = () => { viewManage.classList.remove('hidden'); viewFill.classList.add('hidden'); tabManage.classList.add('active'); tabFill.classList.remove('active'); };

    if (btnManagerPick) {
        btnManagerPick.onclick = () => {
            saveDraft(); // 先存草稿
            chrome.storage.local.set({ 'oa_picker_active': true });
            runInTab(() => { const b = document.getElementById('field1366309_browserbtn'); if (b) b.click(); });
            updateStatus("🔎 選人後重新打開此視窗即可...", "#3498db");
            // 此時彈窗會關閉，捕捉交給 content.js
        };
    }

    if (btnSave) btnSave.onclick = async () => {
        const p = { id: editIdInput.value || "p_" + Date.now(), managerId: inputManagerId.value };
        fields.forEach(f => p[f] = inputs[f] ? inputs[f].value.trim() : "");
        if (!p.label || !p.vendorName) return updateStatus("⚠️ 標題和公司名必填", "#e74c3c");
        const idx = projects.findIndex(x => x && x.id === p.id);
        if (idx > -1) projects[idx] = p; else projects.push(p);
        await DB.save(projects);
        renderList(); resetForm(); updateStatus("✅ 已儲存項目", "#2ecc71");
    };

    if (btnCancel) btnCancel.onclick = () => { resetForm(); chrome.storage.local.remove('oa_form_draft'); };

    // --- 4. 填充功能 ---
    function runFill(p) {
        updateStatus("🚀 正在填充...", "#3498db");
        runInTab((ids) => {
            const bak = {};
            ids.forEach(id => { const e = document.getElementById(id) || document.querySelector(`[name="${id}"]`); bak[id] = e ? e.value : ""; });
            return bak;
        }, [['field1366311', 'field1366312', 'field1366309', 'field1366275', 'field1366310', 'field1366280', 'field1366276', 'field1366278', 'field1366286_0', 'field1366287_0', 'field1366289_0', 'field1366290', 'field1366291', 'field1366292', 'field1366313', 'field1366294']], (res) => {
            if (res?.[0]?.result) lastBackup = res[0].result;
            const msg = { type: "OA_LINK_FILL", project: JSON.parse(JSON.stringify(p)) };
            runInTab((m) => {
                window.postMessage(m, "*");
                function b(w) { for (let i = 0; i < w.frames.length; i++) { try { w.frames[i].postMessage(m, "*"); b(w.frames[i]); } catch (e) { } } }
                b(window);
            }, [msg], () => updateStatus("✅ 已完成填充", "#2ecc71"));
        });
    }

    // --- 5. 匯入匯出 ---
    if (btnExport) btnExport.onclick = () => {
        const h = ["項目標題", "物業名稱", "報價編號", "項目經理", "項目內容", "立項預算", "合約總價", "報價形式", "採購類別", "幣種", "承判商", "報價內容", "金額", "邀請公司", "有效報價", "推荐理由"];
        const r = projects.map(x => fields.map(f => `"${(x[f] || '').replace(/"/g, '""')}"`).join(","));
        const b = new Blob(["\uFEFF" + h.join(",") + "\n" + r.join("\n")], { type: 'text/csv;charset=utf-8' });
        const a = document.createElement('a'); a.href = URL.createObjectURL(b); a.download = `Projects_${Date.now()}.csv`; a.click();
    };

    if (btnImport) btnImport.onclick = () => inputImport.click();
    if (inputImport) inputImport.onchange = (e) => {
        const reader = new FileReader();
        reader.onload = async (ev) => {
            const lines = ev.target.result.split(/\r?\n/).filter(l => l.trim().length > 0);
            const imp = lines.slice(1).map((l, idx) => {
                const cols = []; let cur = "", inQ = false;
                for (let i = 0; i < l.length; i++) { if (l[i] === '"') { if (inQ && l[i + 1] === '"') { cur += '"'; i++; } else { inQ = !inQ; } } else if (l[i] === ',' && !inQ) { cols.push(cur); cur = ""; } else { cur += l[i]; } }
                cols.push(cur);
                if (cols.length < 3) return null;
                const p = { id: "p_" + Date.now() + "_" + idx }; fields.forEach((f, i) => p[f] = (cols[i] || "").trim()); return p;
            }).filter(x => x);
            if (imp.length > 0) { projects = (confirm(`匯入 ${imp.length} 筆？`)) ? imp : [...projects, ...imp]; await DB.save(projects); renderList(); updateStatus("✅ 匯入完畢", "#2ecc71"); }
            inputImport.value = "";
        };
        reader.readAsText(e.target.files[0]);
    };

    function renderList() {
        vendorListEl.innerHTML = projects.length === 0 ? '<p class="hint">尚無項目 (請點擊匯入 CSV)</p>' : '';
        projects.forEach(p => {
            const div = document.createElement('div'); div.className = 'vendor-item';
            div.innerHTML = `<div class="vendor-info"><span class="btn-title">${esc(p.label)}</span><span class="btn-subtitle">${esc(p.vendorName)}</span></div><div class="actions"><button class="icon-btn ed-btn">✏️</button><button class="icon-btn del-btn">🗑️</button></div>`;
            div.querySelector('.vendor-info').onclick = () => runFill(p);
            div.querySelector('.ed-btn').onclick = (e) => { e.stopPropagation(); tabManage.click(); formTitle.textContent = "編輯項目"; editIdInput.value = p.id; fields.forEach(f => { if (inputs[f]) inputs[f].value = p[f] || ""; }); inputManagerId.value = p.managerId || ""; btnCancel.classList.remove('hidden'); btnSave.textContent = "更新項目"; saveDraft(); };
            div.querySelector('.del-btn').onclick = async (e) => { e.stopPropagation(); if (confirm(`刪除「${p.label}」？`)) { projects = projects.filter(x => x && x.id !== p.id); await DB.save(projects); renderList(); } };
            vendorListEl.appendChild(div);
        });
    }

    function resetForm() { formTitle.textContent = "新增項目"; editIdInput.value = ""; inputManagerId.value = ""; fields.forEach(f => { if (inputs[f]) inputs[f].value = (f === 'currency' ? 'HKD' : ""); }); btnCancel.classList.add('hidden'); btnSave.textContent = "儲存"; }
    function updateStatus(m, c) { statusEl.textContent = m; statusEl.style.color = c; setTimeout(() => { statusEl.textContent = "準備就緒"; statusEl.style.color = "#95a5a6"; }, 3000); }
    function esc(s) { const d = document.createElement('div'); d.textContent = s || ""; return d.innerHTML; }
    function runInTab(f, a, c) { chrome.tabs.query({ active: true, currentWindow: true }, (t) => { if (t[0]) chrome.scripting.executeScript({ target: { tabId: t[0].id, allFrames: true }, func: f, args: a || [] }, c); }); }
});
