/**
 * OA Bridge - Background Service Worker
 * Version: 4.11.5
 * 作為代理者幫助 content script 調用 chrome.downloads API
 */

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
    if (msg.type === 'DOWNLOAD_CSV') {
        // content script 無法直接使用 chrome.downloads，由後台代理
        chrome.downloads.download({
            url: msg.dataUrl,
            filename: 'OA_Projects.csv',
            conflictAction: 'overwrite',
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
        return true; // 保持非同步回應通道
    }
});
