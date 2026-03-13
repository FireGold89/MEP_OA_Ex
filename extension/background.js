/**
 * OA Bridge - Background Service Worker
 * Version: 4.11.7
 */

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
    if (msg.type === 'DOWNLOAD_CSV') {
        // content script 無法直接使用 chrome.downloads，由後台代理
        // msg.filename: 'OA_Projects.csv' 或帶時間戳的備用名稱
        // msg.conflictAction: 'overwrite' 或 'uniquify'
        chrome.downloads.download({
            url: msg.dataUrl,
            filename: msg.filename || 'OA_Projects.csv',
            conflictAction: msg.conflictAction || 'overwrite',
            saveAs: false
        }, (downloadId) => {
            if (chrome.runtime.lastError) {
                console.warn('OA Background: Download error -', chrome.runtime.lastError.message);
                sendResponse({ success: false, error: chrome.runtime.lastError.message });
            } else {
                console.log('OA Background: Download started, id =', downloadId);
                sendResponse({ success: true });
            }
        });
        return true;
    }
});
