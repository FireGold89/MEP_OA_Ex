/**
 * Weaver OA Main Bridge
 * Version: 4.6.2
 * 核心功能：全自動增減行 + 深度修復人員選取器 (e8_browser) 填充
 */

(function () {
    console.log("%cOA Bridge v4.6.2: Active.", "color:white; background:#00BCD4; padding:2px 5px; border-radius:3px;");

    window.addEventListener("message", function (event) {
        if (!event.data || event.data.type !== "OA_LINK_FILL") return;

        const data = event.data.project;
        if (!data) return;

        const hasTargets = ['field1366311', 'field1366275', 'field1366310'].some(id => !!(document.getElementById(id) || document.querySelector('[name="' + id + '"]')));
        if (!hasTargets) return;

        console.log("OA Bridge: Filling Data for " + data.label, data);

        // 1. 填充主表 (包含增強的人員 ID 處理)
        const staticMap = {
            'field1366311': data.propertyName, 'field1366312': data.reportNo, 'field1366309': data.manager,
            'field1366275': data.projectContent, 'field1366310': data.total, 'field1366280': data.budget,
            'field1366276': data.quoteType, 'field1366278': data.buyType, 'field1366279': data.currency,
            'field1366290': data.inviteCount, 'field1366291': data.replyCount, 'field1366292': data.reason,
            'field1366313': data.winnerName, 'field1366294': data.contractAmount, 'field1366295': data.contractCurrency
        };
        for (let id in staticMap) {
            // 傳遞 managerId 供人員欄位使用
            const mId = (id === 'field1366309') ? data.managerId : null;
            fillField(id, staticMap[id], mId);
        }

        // 2. 填充明細與同步行數
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

        // 處理下拉選單
        if (el.tagName === 'SELECT') {
            let ok = "";
            for (let o of el.options) { if (o.text.trim() === val || o.value === val || o.text.includes(val)) { ok = o.value; break; } }
            if (ok !== "") el.value = ok;
        } else {
            // 處理普通輸入框與人員 ID
            if (mId && id === 'field1366309') {
                el.value = mId; // 填入真正的系統 ID
            } else {
                el.value = val || "";
            }
        }

        // 核心：深度修復人员選取器 (Browser 欄位) 的顯示狀態
        if (id === 'field1366309') {
            const displayVal = val || "";
            // 1. 更新 Span 顯示內容
            const spans = [id + 'span', id + '_span', id + 'spanimg'];
            spans.forEach(sid => {
                const s = document.getElementById(sid);
                if (s) {
                    s.innerHTML = '<span class="e8_showNameClass"><a href="javascript:void(0);">' + displayVal + '</a><span class="e8_delClass" onclick="clearFullbrow(' + id.replace('field', '') + ')"> x </span></span>';
                    s.style.display = (displayVal === "") ? "" : "inline-block";
                }
            });
            // 2. 移除 OA 的必填警告驚嘆號
            const img = document.getElementById(id + 'spanimg');
            if (img) img.style.display = 'none';
        }

        // 觸發 OA 內部邏輯
        try {
            if (window.wfbrowvaluechange) window.wfbrowvaluechange(el, id.replace('field', ''));
            if (window.checkinput2) window.checkinput2(id, id + 'span', el.getAttribute('viewtype') || '1');
        } catch (e) { }

        // 觸發標準事件
        ['change', 'input', 'blur'].forEach(t => el.dispatchEvent(new Event(t, { bubbles: true })));
    }
})();
