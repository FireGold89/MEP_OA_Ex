/**
 * 數據緊急救回工具 v1.0
 * 請在擴充功能圖示上按右鍵 ->「檢查」 (Inspect)，在 Console 分頁貼上此代碼。
 */
(async function () {
    console.log("%c--- 數據救回診斷開始 ---", "color: #3498db; font-size: 16px; font-weight: bold;");

    // 1. 檢查 chrome.storage.local
    if (typeof chrome !== 'undefined' && chrome.storage && chrome.storage.local) {
        const storageData = await new Promise(r => chrome.storage.local.get(null, r));
        console.log("Chrome Storage 內容:", storageData);
        for (let key in storageData) {
            if (Array.isArray(storageData[key])) {
                console.log(`%cFound data in key [${key}]:`, "color: #2ecc71; font-weight: bold;", storageData[key]);
            }
        }
    }

    // 2. 檢查 localStorage
    console.log("LocalStorage 內容:");
    for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        const val = localStorage.getItem(key);
        console.log(`Key: ${key}`, val);
    }

    alert("診斷完成！請查看瀏覽器控制台 (Console) 中的輸出，截圖給工程師。");
})();
