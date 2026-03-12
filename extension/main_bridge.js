/**
 * Weaver OA Main Bridge
 * Version: 4.9.3 - Multi-Frame Broadcast Fix
 * 核心功能：全自動增減行 + 強化版人員選取器 + 跨 Iframe 廣播轉發
 */

(function () {
    const isTop = (window.self === window.top);
    console.log(`%cOA Bridge v4.11.0: Active [${isTop ? 'TOP' : 'IFRAME'}]`, "color:white; background:#764ba2; padding:2px 5px; border-radius:3px;");

    // 全球暫存，供人員選擇器自動匹配
    window.__OA_LAST_PROJECT = window.__OA_LAST_PROJECT || null;

    // 監聽來自內容腳本或其他窗口的 postMessage
    window.addEventListener("message", function (event) {
        // 安全校驗：非本頁發出且非廣播訊息則跳過
        if (!event.data || (event.data.type !== "OA_LINK_FILL" && event.data.type !== "OA_LINK_CLEAR" && event.data.type !== "OA_BRIDGE_BROADCAST")) return;

        // 如果收到的是廣播包裹，解包
        const payload = event.data.type === "OA_BRIDGE_BROADCAST" ? event.data.payload : event.data;
        if (isTop && event.data.type !== "OA_BRIDGE_BROADCAST") {
            // 🌟 頂層視窗核心邏輯：向所有子 iframe 進行廣播
            console.log("%cOA Bridge: Top-Level Signal Captured. Broadcasting to frames...", "color:blue; font-weight:bold;");
            
            // 診斷日誌：掃描當前頁面所有 field 欄位
            diagnoseFields();
            
            broadcastToFrames(window, { type: "OA_BRIDGE_BROADCAST", payload: payload });
        }

        // 執行填充檢測
        processSignal(payload);
    });

    function broadcastToFrames(win, msg) {
        for (let i = 0; i < win.frames.length; i++) {
            try {
                const subWin = win.frames[i];
                subWin.postMessage(msg, "*");
                
                // 子視窗診斷
                if (subWin.diagnoseFields) subWin.diagnoseFields(); 
                
                broadcastToFrames(subWin, msg);
            } catch (e) {
                // 忽略跨域 iframe 導致的權限錯誤
            }
        }
    }

    /**
     * 診斷日誌：打印頁面上發現的所有 field 欄位，協助定位 ID
     */
    function diagnoseFields() {
        const fields = Array.from(document.querySelectorAll('[id^="field"], [name^="field"], select'));
        if (fields.length > 0) {
            console.log(`%c[OA Bridge Diagnostics] Scan in ${window.location.pathname}`, "color:purple; font-weight:bold;");
            fields.forEach(el => {
                // 如果是 Select，嘗試找出它的標籤文字
                if (el.tagName === 'SELECT') {
                    const parentTd = el.closest('td');
                    const labelTd = parentTd ? parentTd.previousElementSibling : null;
                    const labelText = labelTd ? labelTd.innerText.trim() : "Unknown";
                    console.log(`%cFound SELECT: ID=${el.id || el.name}, Label Context="${labelText}"`, "color: brown;");
                }
            });
        }
    }
    window.diagnoseFields = diagnoseFields; // 暴露給全局方便遞迴調用

    /**
     * 處理填充或清理邏輯
     */
    function processSignal(data) {
        if (data.type === "OA_LINK_CLEAR") {
            console.log("OA Bridge: Clearing fields locally...");
            clearAllFields();
            return;
        }

        const project = data.project;
        if (!project) return;

        window.__OA_LAST_PROJECT = project;

        // 核心檢查：當前窗口是否有目標欄位 (擴展 ID 列表以提高檢測機率)
        const hasTargets = ['field1366311', 'field1366275', 'field1366310', 'field1366277', 'field1366278'].some(id => !!(document.getElementById(id) || document.querySelector('[name="' + id + '"]')));
        
        if (!hasTargets) {
            // 僅在存在 project 數據時輸出弱提示，減少控制台噪音
            if (isTop) console.log("OA Bridge: (Top) No fields here, waiting for iframe processing.");
            return; 
        }

        console.log("%cOA Bridge: TARGET FOUND! Filling -> " + project.label, "color:green; font-weight:bold; padding:4px; border:2px solid green;");

        // 1. 填充主表
        const staticMap = {
            'field1366311': project.propertyName, 'field1366312': project.reportNo, 'field1366309': project.manager,
            'field1366275': project.projectContent, 'field1366310': project.total, 'field1366280': project.budget,
            'field1366276': project.quoteType, 'field1366278': project.buyType, 'field1366279': project.currency,
            'field1366290': project.inviteCount, 'field1366291': project.replyCount, 'field1366292': project.reason,
            'field1366313': project.winnerName, 'field1366294': project.contractAmount, 'field1366295': project.contractCurrency
        };
        
        for (let id in staticMap) {
            // 採購類別 ID 容錯：深度挖掘
            if (id === 'field1366278') {
                const buyTypeVal = staticMap[id];
                // 方案 A: 嘗試已知 ID
                const knownIds = ['field1366278', 'field1366277', 'field1366279']; 
                let filled = false;
                for (let aid of knownIds) {
                    if (fillField(aid, buyTypeVal)) { filled = true; break; }
                }
                
                // 方案 B: 模糊匹配 (如果 A 失敗)
                if (!filled) {
                    console.log("OA Bridge: Known IDs failed for BuyType, attempting fuzzy scan...");
                    const selects = Array.from(document.querySelectorAll('select'));
                    for (let s of selects) {
                        const pTd = s.closest('td');
                        const lTd = pTd ? pTd.previousElementSibling : null;
                        if (lTd && (lTd.innerText.includes('採購類別') || lTd.innerText.includes('採購類'))) {
                            console.log(`OA Bridge: Fuzzy match found BuyType at ID=${s.id || s.name}`);
                            if (fillField(s.id || s.name, buyTypeVal)) { filled = true; break; }
                        }
                    }
                }
                continue;
            }

            const mId = (id === 'field1366309') ? project.managerId : null;
            fillField(id, staticMap[id], mId);
        }

        // 2. 輔助人員選擇器
        if (project.label !== "恢復備份") {
            setTimeout(() => {
                const mField = document.getElementById('field1366309');
                if (mField && (!mField.value || mField.value === "")) {
                    const btn = document.getElementById('field1366309_browserbtn');
                    if (btn) btn.click();
                }
            }, 1000);
        }

        // 3. 填充明細
        if (project.details && project.details.length > 0) {
            handleDetails(project.details);
        }
    }

    async function handleDetails(details) {
        const requiredCount = details.length;
        for (let i = 0; i < requiredCount; i++) {
            const rowId = `field1366286_${i}`;
            let targetEl = document.getElementById(rowId);
            
            if ((!targetEl || targetEl.tagName !== 'INPUT') && i > 0) {
                const addBtn = findActionBtnAcrossFrames(window.top, 'add');
                if (addBtn) {
                    addBtn.click();
                    await new Promise(r => setTimeout(r, i === 1 ? 1200 : 800));
                    targetEl = document.getElementById(rowId);
                }
            }
            const row = details[i];
            fillField(`field1366286_${i}`, row.vendorName);
            fillField(`field1366287_${i}`, row.content);
            fillField(`field1366288_${i}`, row.detailCurrency);
            fillField(`field1366289_${i}`, row.amount);
        }
        
        // 刪除多餘行
        let excessIdx = requiredCount;
        let hasExcess = false;
        while (document.getElementById(`field1366286_${excessIdx}`)) {
            const chk = findCheckboxForRow(excessIdx);
            if (chk) { if (!chk.checked) chk.click(); hasExcess = true; }
            excessIdx++;
            if (excessIdx > 50) break;
        }
        if (hasExcess) {
            const delBtn = findActionBtnAcrossFrames(window.top, 'del');
            if (delBtn) {
                const org = window.confirm; window.confirm = () => true; delBtn.click();
                setTimeout(() => { window.confirm = org; }, 1000);
            }
        }
    }

    function fillField(id, val, mId = null) {
        const el = document.getElementById(id) || document.querySelector('[name="' + id + '"]');
        if (!el) return;

        const isBrowser = (id === 'field1366309') || el.classList.contains('e8_browser') || !!document.getElementById(id + 'span');

        if (el.tagName === 'SELECT') {
            let ok = "";
            const valStr = String(val).trim();
            const optionsInfo = Array.from(el.options).map(o => `[val:${o.value}, txt:${o.text.trim()}]`).join(", ");
            console.log(`OA Bridge: [SELECT] Fill ${id} with "${valStr}". Options: ${optionsInfo}`);

            for (let o of el.options) {
                const optVal = String(o.value).trim();
                const optText = o.text.trim();
                
                // 擴大匹配範圍：值相符、文字相符、或關鍵字包含
                if (optVal === valStr || optText === valStr || (valStr.length > 0 && optText.includes(valStr)) || (optVal.length > 0 && valStr.includes(optVal))) {
                    ok = o.value; 
                    console.log(`OA Bridge: [SELECT] Match Found! -> "${optText}" (Value: ${ok})`);
                    break; 
                }
            }
            if (ok !== "") {
                el.value = ok;
                // 觸發 OA 內部邏輯
                try { if (window.onsh && typeof window.onsh === 'function') window.onsh(el); } catch(e){}
                return true; 
            }
            console.warn(`OA Bridge: [SELECT] No match for "${valStr}" in field ${id}`);
            return false;
        } else {
            const finalVal = (mId && isBrowser) ? mId : (val || "");
            el.value = finalVal;
            const elShow = document.getElementById(id + '__');
            if (elShow) elShow.value = finalVal;
        }

        if (isBrowser) {
            const displayVal = val || "";
            const spans = [id + 'span', id + '_span'];
            spans.forEach(sid => {
                const s = document.getElementById(sid);
                if (s) {
                    if (displayVal === "") { s.innerHTML = ""; s.style.display = "none"; }
                    else {
                        s.innerHTML = `<span class="e8_showNameClass"><a href="javascript:void(0)" onclick="return false;">${displayVal}</a></span>`;
                        s.style.display = "inline-block";
                    }
                }
            });
        }

        ['change', 'input', 'blur'].forEach(t => {
            const ev = new Event(t, { bubbles: true });
            el.dispatchEvent(ev);
        });
    }

    function clearAllFields() {
        const ids = ['field1366311', 'field1366312', 'field1366309', 'field1366275', 'field1366310', 'field1366280', 'field1366276', 'field1366278', 'field1366279', 'field1366290', 'field1366291', 'field1366292', 'field1366313', 'field1366294', 'field1366295'];
        ids.forEach(id => fillField(id, ""));
    }

    function findActionBtnAcrossFrames(win, type) {
        const id = type === 'add' ? 'addbutton0' : 'delbutton0';
        try {
            const doc = win.document;
            let b = doc.getElementById(id) || doc.querySelector(`[name="${id}"]`);
            if (b) return b;
            for (let i = 0; i < win.frames.length; i++) {
                const found = findActionBtnAcrossFrames(win.frames[i], type);
                if (found) return found;
            }
        } catch (e) { } return null;
    }

    function findCheckboxForRow(idx) {
        const cell = document.getElementById(`field1366286_${idx}`);
        if (cell) {
            const tr = cell.closest('tr');
            if (tr) return tr.querySelector('input[type="checkbox"]');
        }
        return null;
    }

    function autoClickPickerTarget() {
        setInterval(() => {
            const data = window.__OA_LAST_PROJECT;
            if (!data || !data.manager) return;
            const targetName = data.manager.replace(/\s+/g, '').toLowerCase();

            function searchAndClick(win) {
                try {
                    const fDoc = win.document;
                    if (!fDoc || fDoc.body.getAttribute('oa-processed')) return false;
                    const elements = Array.from(fDoc.querySelectorAll('td, span, div.e8_browser_item'));
                    const targetEl = elements.find(el => {
                        const txt = (el.innerText || "").replace(/\s+/g, '').toLowerCase();
                        return txt && (txt === targetName || txt.includes(targetName)) && txt.length < 20;
                    });
                    if (targetEl) {
                        fDoc.body.setAttribute('oa-processed', 'true');
                        const row = targetEl.closest('tr') || targetEl;
                        row.click();
                        const dblEvent = new MouseEvent('dblclick', { bubbles: true, view: win });
                        row.dispatchEvent(dblEvent);
                        return true;
                    }
                    for (let i = 0; i < win.frames.length; i++) {
                        if (searchAndClick(win.frames[i])) return true;
                    }
                } catch (e) { } return false;
            }
            const frames = Array.from(document.querySelectorAll('iframe[id*="DialogFrame"]'));
            frames.forEach(frame => {
                if (frame.contentWindow) searchAndClick(frame.contentWindow);
            });
        }, 1000);
    }

    autoClickPickerTarget();
})();
