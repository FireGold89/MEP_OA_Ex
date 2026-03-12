/**
 * Weaver OA Main Bridge
 * Version: 4.7.1
 * 核心功能：全自動增減行 + 強化版人員選取器 (e8_browser) 深度回寫與彈窗輔助
 */

(function () {
    console.log("%cOA Bridge v4.7.1: Power Active.", "color:white; background:#764ba2; padding:2px 5px; border-radius:3px;");

    window.addEventListener("message", function (event) {
        if (!event.data || event.data.type !== "OA_LINK_FILL") return;

        const data = event.data.project;
        if (!data) return;

        // 全球暫存，供彈窗路徑追蹤
        window.__OA_LAST_PROJECT = data;

        const hasTargets = ['field1366311', 'field1366275', 'field1366310'].some(id => !!(document.getElementById(id) || document.querySelector('[name="' + id + '"]')));
        if (!hasTargets) return;

        console.log("OA Bridge: Power Filling for " + data.label, data);

        // 1. 填充主表 (包含深度回寫邏輯)
        const staticMap = {
            'field1366311': data.propertyName, 'field1366312': data.reportNo, 'field1366309': data.manager,
            'field1366275': data.projectContent, 'field1366310': data.total, 'field1366280': data.budget,
            'field1366276': data.quoteType, 'field1366278': data.buyType, 'field1366279': data.currency,
            'field1366290': data.inviteCount, 'field1366291': data.replyCount, 'field1366292': data.reason,
            'field1366313': data.winnerName, 'field1366294': data.contractAmount, 'field1366295': data.contractCurrency
        };
        for (let id in staticMap) {
            const mId = (id === 'field1366309') ? data.managerId : null;
            fillField(id, staticMap[id], mId);
        }

        // 2. 輔助觸發：若人員欄位仍空，1.5秒後自動點擊放大鏡觸發彈窗觀察者 (恢復備份時跳過)
        if (data.label !== "恢復備份") {
            setTimeout(() => {
                const mField = document.getElementById('field1366309');
                if (mField && (!mField.value || mField.value === "")) {
                    const btn = document.getElementById('field1366309_browserbtn');
                    if (btn) {
                        console.log("OA Bridge: Still empty, launching picker assistant...");
                        btn.click();
                    }
                }
            }, 1500);
        }

        // 3. 填充明細與同步行數
        if (data.details && data.details.length > 0) {
            (async function () {
                const requiredCount = data.details.length;
                for (let i = 0; i < requiredCount; i++) {
                    const rowId = `field1366286_${i}`;
                    let targetEl = document.getElementById(rowId);
                    if ((!targetEl || targetEl.tagName !== 'INPUT') && i > 0) {
                        const addBtn = findActionBtnAcrossFrames(window.top, 'add');
                        if (addBtn) {
                            addBtn.click();
                            await new Promise(r => setTimeout(r, i === 1 ? 1200 : 700));
                            targetEl = document.getElementById(rowId);
                        }
                    }
                    const row = data.details[i];
                    fillField(`field1366286_${i}`, row.vendorName);
                    fillField(`field1366287_${i}`, row.content);
                    fillField(`field1366288_${i}`, row.detailCurrency);
                    fillField(`field1366289_${i}`, row.amount);
                }

                let hasExcess = false;
                let excessIdx = requiredCount;
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
                console.log("OA Bridge: All tasks finished.");
            })();
        }
    });

    window.addEventListener("message", function (event) {
        if (!event.data || event.data.type !== "OA_LINK_CLEAR") return;
        console.log("OA Bridge: Clearing all fields...");
        clearAllFields();
    });

    function clearAllFields() {
        const ids = ['field1366311', 'field1366312', 'field1366309', 'field1366275', 'field1366310', 'field1366280', 'field1366276', 'field1366278', 'field1366279', 'field1366290', 'field1366291', 'field1366292', 'field1366313', 'field1366294', 'field1366295'];
        ids.forEach(id => fillField(id, ""));

        // 清理明細行
        let idx = 0;
        while (document.getElementById(`field1366286_${idx}`)) {
            fillField(`field1366286_${idx}`, "");
            fillField(`field1366287_${idx}`, "");
            fillField(`field1366289_${idx}`, "");
            const chk = findCheckboxForRow(idx);
            if (chk && !chk.checked) chk.click();
            idx++;
        }
    }

    function findActionBtnAcrossFrames(win, type) {
        const id = type === 'add' ? 'addbutton0' : 'delbutton0';
        const selectors = [`#${id}`, `[name="${id}"]`, `[id^="${id.slice(0, -1)}"]`, type === 'add' ? 'input[onclick*="addRow"]' : 'input[onclick*="deleteRow"]'];
        try {
            const doc = win.document;
            for (let s of selectors) { const b = doc.querySelector(s); if (b) return b; }
            for (let i = 0; i < win.frames.length; i++) { const found = findActionBtnAcrossFrames(win.frames[i], type); if (found) return found; }
        } catch (e) { } return null;
    }

    function findCheckboxForRow(idx) {
        let chk = document.getElementById(`check_node_0_${idx}`) || document.querySelector(`input[name="check_node_0"][value="${idx}"]`);
        if (chk) return chk;
        const input = document.getElementById(`field1366286_${idx}`);
        if (input) { const tr = input.closest('tr'); if (tr) return tr.querySelector('input[type="checkbox"]'); }
        return null;
    }

    function fillField(id, val, mId = null) {
        const el = document.getElementById(id) || document.querySelector('[name="' + id + '"]');
        if (!el) return;

        const isBrowser = (id === 'field1366309') || el.classList.contains('e8_browser') || !!document.getElementById(id + 'span');

        // 處理下拉選單
        if (el.tagName === 'SELECT') {
            let ok = "";
            for (let o of el.options) { if (o.text.trim() === val || o.value === val || o.text.includes(val)) { ok = o.value; break; } }
            if (ok !== "") el.value = ok;
        } else {
            // 處理普通輸入框與人員 ID
            const finalVal = (mId && isBrowser) ? mId : (val || "");
            el.value = finalVal;

            // 核心修正：處理泛微 OA 的雙下劃線隱藏欄位 (如 field1366309__)
            const elShow = document.getElementById(id + '__') || document.querySelector('[name="' + id + '__"]');
            if (elShow) elShow.value = finalVal;
        }

        // 如果是人員選取器，嘗試用多種方式回寫數據與顯示
        if (isBrowser) {
            const displayVal = val || "";
            // 1. 更新顯示容器 (核心：Weaver OA 的人員顯示)
            const spans = [id + 'span', id + '_span', id + 'spanimg'];
            spans.forEach(sid => {
                const s = document.getElementById(sid);
                if (s) {
                    if (displayVal === "") { s.innerHTML = ""; s.style.display = "none"; }
                    else {
                        // 構建與系統原生完全一致的 HTML 結構
                        s.innerHTML = `<span class="e8_showNameClass"><a href="javascript:void(0)" onclick="return false;">${displayVal}</a><span class="e8_delClass" style="visibility:visible; cursor:pointer;" onclick="if(window.clearFullbrow) window.clearFullbrow('${id.replace('field', '')}', '${id}span', '1');"> x </span></span>`;
                        s.style.display = "inline-block";
                    }
                }
            });
            if (displayVal !== "") {
                const img = document.getElementById(id + 'spanimg'); if (img) img.style.display = 'none';
            }

            // 2. 深入嘗試：調用 OA 原生回寫函數 (針對某些版本的生態系統)
            try {
                const fieldNum = id.replace('field', '');
                // 嘗試常用的 OA 數據回填黑科技
                if (mId && displayVal) {
                    if (window._writeBackData) window._writeBackData(id, displayVal, mId);
                    else if (window._writeBackBrowser) window._writeBackBrowser(id, displayVal, mId);
                }
            } catch (e) { }
        }

        // 觸發 OA 內部校驗與邏輯同步
        try {
            // 嘗試觸發系統原生的 value 變動回調
            const fieldNum = id.replace('field', '');
            if (window.wfbrowvaluechange) window.wfbrowvaluechange(el, fieldNum);
            if (window.checkinput2) {
                const viewType = el.getAttribute('viewtype') || '1';
                window.checkinput2(id, id + 'span', viewType);
            }
        } catch (e) { console.warn("OA Bridge: Validation Trigger Error ->", e); }

        // 觸發標準事件
        ['change', 'input', 'blur'].forEach(t => {
            const ev = new Event(t, { bubbles: true });
            el.dispatchEvent(ev);
            const elShow = document.getElementById(id + '__');
            if (elShow) elShow.dispatchEvent(ev);
        });
    }

    /**
     * 新增：深度監控彈窗並自動點擊目標
     */
    function autoClickPickerTarget() {
        setInterval(() => {
            const frames = Array.from(document.querySelectorAll('iframe[id*="DialogFrame"]'));
            if (frames.length === 0) return;

            const data = window.__OA_LAST_PROJECT;
            if (!data || !data.manager) return;
            const targetName = data.manager;

            frames.forEach(frame => {
                try {
                    const fDoc = frame.contentDocument || frame.contentWindow.document;
                    if (fDoc.body.getAttribute('oa-processed')) return;

                    // 1. 定義目標查找 (Weaver 通常有 Dennis Chan 的文字)
                    const elements = Array.from(fDoc.querySelectorAll('td, span, div, a'));
                    const targetEl = elements.find(el => {
                        const txt = el.innerText.trim();
                        return txt === targetName || txt.startsWith(targetName) || (txt.includes(targetName) && txt.length < 20);
                    });

                    if (targetEl) {
                        console.log("OA Bridge: Picker Found Power Target ->", targetName);
                        const row = targetEl.closest('tr') || targetEl;

                        // 2. 模擬點擊：如果是單選，Weaver 常需雙擊 (dbclick) 或 點擊後按 OK
                        row.click();
                        // 觸發雙擊，很多彈窗雙擊即選中並關閉
                        const dblEvent = new MouseEvent('dblclick', { bubbles: true, cancelable: true, view: frame.contentWindow });
                        row.dispatchEvent(dblEvent);

                        // 3. 備用方案：尋找「確定」按鈕
                        setTimeout(() => {
                            const btns = Array.from(fDoc.querySelectorAll('button, input[type="button"]'));
                            const okBtn = btns.find(b => b.value?.includes('確定') || b.innerText?.includes('確定') || b.id === 'btnok');
                            if (okBtn) {
                                okBtn.click();
                                // 嘗試調用 Weaver OA 可能存在的全局回調函數
                                if (frame.contentWindow.doCallback) {
                                    frame.contentWindow.doCallback();
                                } else if (frame.contentWindow.opener && frame.contentWindow.opener.doCallback) {
                                    frame.contentWindow.opener.doCallback();
                                }
                            }
                        }, 500);

                        fDoc.body.setAttribute('oa-processed', 'true');
                    }
                } catch (e) { }
            });
        }, 800);
    }

    autoClickPickerTarget();
})();
