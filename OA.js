// ==UserScript==
// @name         採購立項自動化
// @match        *://*/workflow/request/AddRequest.jsp* /* 這是泛微新建流程的網址特徵 */
// @grant        none
// @run-at       document-end  /* 確保頁面載入完成後才執行 */
// ==/UserScript==

(function () {
    'use strict';

    const vendorDatabase = {
        "1": {
            label: "信成 - 電器材料",
            vendorName: "信成電器材料有限公司",
            content: "電器材料及配件報價",
            amount: "12152.60",
            reason: "長期惠顧的供應商，貨真價實，配送快速"
        }
    };

    const baseConfig = {
        'field1366311': '九龍灣車廠',
        'field1366312': 'MS/Q1241/24/kp',
        'field1366275': '康樂中心及健身室擴建工程',
        'field1366310': '7400000.00',
        'field1366280': '12152.60',
        'field1366276': '0',
        'field1366278': '11'
    };

    function runAutoFill(selectedData) {
        const finalConfig = {
            ...baseConfig,
            'field1366286_0': selectedData.vendorName,
            'field1366287_0': selectedData.content,
            'field1366289_0': selectedData.amount,
            'field1366292': selectedData.reason,
            'field1366313': selectedData.vendorName,
            'field1366294': selectedData.amount
        };

        function process(win) {
            let doc = win.document;

            /* 1. 填充普通欄位 */
            for (let id in finalConfig) {
                let el = doc.getElementById(id);
                if (el) {
                    el.focus(); // 先獲取焦點
                    el.value = finalConfig[id];
                    if (el.onchange) el.onchange();
                    el.blur();  // 失去焦點觸發校驗
                    let s = doc.getElementById(id + 'span');
                    if (s && id.indexOf('_') === -1) s.innerHTML = '';
                }
            }

            /* 2. 核心修正：項目經理 (field1366309) 模擬手動觸發 */
            let mId = 'field1366309';
            let mInput = doc.getElementById(mId);
            let mInputShow = doc.getElementById(mId + '__');

            if (mInput) {
                // 模擬真實點擊流程
                if (mInputShow) mInputShow.focus();

                mInput.value = '48313';
                if (mInputShow) mInputShow.value = '48313';

                let mSpan = doc.getElementById(mId + 'span');
                if (mSpan) {
                    mSpan.innerHTML = `
                        <span class="e8_showNameClass">
                            <a href="javascript:void(0)">Dennis Chan 陳家俊</a>
                            <span class="e8_delClass" style="visibility:visible;">&nbsp;x&nbsp;</span>
                        </span>`;
                }

                // 強制觸發系統校驗函數
                try {
                    if (typeof win.checkinput2 === 'function') {
                        win.checkinput2(mId, mId + 'span', '1');
                    }
                    // 模擬分派事件，讓系統感覺到「值變了」
                    let event = new Event('change', { bubbles: true });
                    mInput.dispatchEvent(event);
                    if (mInputShow) {
                        mInputShow.dispatchEvent(event);
                        mInputShow.blur(); // 觸發失去焦點校驗
                    }
                } catch (e) { }
            }
        }

        process(window);
        for (let i = 0; i < window.frames.length; i++) {
            try { process(window.frames[i]); } catch (e) { }
        }
    }

    const choice = prompt("請選擇供應商編號:", "1");
    if (choice && vendorDatabase[choice]) {
        runAutoFill(vendorDatabase[choice]);
    }
})();