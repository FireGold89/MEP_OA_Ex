/**
 * OA Bridge - Background Service Worker
 * Version: 4.13.1 - 飛書 Bitable 同步
 */

// ===== 下載 CSV =====
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
    if (msg.type === 'DOWNLOAD_CSV') {
        // content script 無法直接使用 chrome.downloads，由後台代理
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

    // ===== 飛書同步 =====
    if (msg.type === 'FEISHU_SYNC') {
        handleFeishuSync(msg.config, msg.projects)
            .then(result => sendResponse({ success: true, ...result }))
            .catch(err => sendResponse({ success: false, error: err.message }));
        return true; // 保持 channel 開放等待 async 回應
    }
});

// ===== 飛書 API 工具函數 =====

/**
 * 取得 tenant_access_token
 */
async function getFeishuToken(appId, appSecret) {
    if (!appId || !appSecret) {
        throw new Error('App ID 或 App Secret 為空');
    }
    console.log('OA Bridge: 正在獲取飛書 Token，appId:', appId);
    const resp = await fetch('https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json; charset=utf-8' },
        body: JSON.stringify({ app_id: appId, app_secret: appSecret })
    });
    const data = await resp.json();
    console.log('OA Bridge: 飛書 Token 響應:', data);
    if (data.code !== 0) throw new Error(`Token 錯誤(${data.code}): ${data.msg}`);
    return data.tenant_access_token;
}

/**
 * 讀取多維表格的欄位清單
 * 返回 Set<fieldName>
 */
async function fetchTableFields(token, appToken, tableId) {
    const url = `https://open.feishu.cn/open-apis/bitable/v1/apps/${appToken}/tables/${tableId}/fields?page_size=100`;
    const resp = await fetch(url, {
        headers: { 'Authorization': `Bearer ${token}` }
    });
    const data = await resp.json();
    if (data.code !== 0) throw new Error(`讀取欄位清單失敗(${data.code}): ${data.msg}`);
    const names = new Set();
    (data.data?.items || []).forEach(f => names.add(f.field_name));
    return names;
}

/**
 * 讀取多維表格所有記錄（自動分頁）
 * 返回 Map<projectId, recordId>
 */
async function fetchAllRecords(token, appToken, tableId) {
    const map = new Map(); // projectId -> recordId
    let pageToken = '';
    do {
        const url = new URL(`https://open.feishu.cn/open-apis/bitable/v1/apps/${appToken}/tables/${tableId}/records`);
        url.searchParams.set('page_size', '500');
        if (pageToken) url.searchParams.set('page_token', pageToken);

        const resp = await fetch(url.toString(), {
            headers: { 'Authorization': `Bearer ${token}` }
        });
        const data = await resp.json();
        if (data.code !== 0) throw new Error(`讀取記錄失敗(${data.code}): ${data.msg}`);

        (data.data?.items || []).forEach(item => {
            const pid = item.fields?.['項目ID'];
            if (pid) map.set(String(pid).trim(), item.record_id);
        });

        pageToken = data.data?.has_more ? data.data.page_token : '';
    } while (pageToken);

    return map;
}

/**
 * 讀取多維表格所有記錄（自動分頁）並轉換為本地項目格式
 * 返回完整項目陣列
 */
async function fetchAllRecordsAsProjects(token, appToken, tableId) {
    const projects = [];
    let pageToken = '';
    do {
        const url = new URL(`https://open.feishu.cn/open-apis/bitable/v1/apps/${appToken}/tables/${tableId}/records`);
        url.searchParams.set('page_size', '500');
        if (pageToken) url.searchParams.set('page_token', pageToken);

        const resp = await fetch(url.toString(), {
            headers: { 'Authorization': `Bearer ${token}` }
        });
        const data = await resp.json();
        if (data.code !== 0) throw new Error(`讀取記錄失敗(${data.code}): ${data.msg}`);

        (data.data?.items || []).forEach(item => {
            const fields = item.fields || {};
            // 只處理有項目ID的記錄
            const projectId = fields['項目ID'];
            if (!projectId) return;

            // 採購類別反轉（文字 → code）
            const buyTypeText = fields['採購類別'] || '';
            const REVERSE_BUY_TYPE = {
                "保養維修材料": "9", "保養合約分判": "10",
                "工程材料": "11", "分判工程": "12",
                "後加工程": "14", "固定資產": "4", "其他": "13"
            };
            const buyType = REVERSE_BUY_TYPE[buyTypeText] || buyTypeText;

            // 幣種反轉（文字 → code）
            const currencyText = fields['幣種'] || 'HKD';
            const REVERSE_CURRENCY = { "HKD": "0", "RMB": "1", "USD": "2", "MOP": "3" };
            const currency = REVERSE_CURRENCY[currencyText] || '0';

            // 時間戳轉 ISO 字串（飛書是毫秒）
            const createdAt = fields['新增時間'] ? new Date(fields['新增時間']).toISOString() : null;
            const updatedAt = fields['修改時間'] ? new Date(fields['修改時間']).toISOString() : null;

            const p = {
                id: "p_" + Date.now() + "_" + Math.floor(Math.random() * 1000),
                projectId: String(projectId).trim(),
                label: fields['項目標題'] || '',
                propertyName: fields['物業名稱'] || '',
                reportNo: fields['美博報價編號'] || '',
                manager: fields['項目經理'] || '',
                projectContent: fields['項目內容'] || '',
                buyType: buyType,
                budget: fields['立項預算'] != null ? String(fields['立項預算']) : '',
                contractAmount: fields['合約金額'] != null ? String(fields['合約金額']) : '',
                currency: currency,
                quoteType: fields['報價形式'] || '',
                winnerName: fields['中標公司'] || '',
                reason: fields['推薦理由'] || '',
                createdAt: createdAt,
                updatedAt: updatedAt,
                details: []
            };
            projects.push(p);
        });

        pageToken = data.data?.has_more ? data.data.page_token : '';
    } while (pageToken);

    return projects;
}

/**
 * 採購類別 code → 文字
 */
const BUY_TYPE_MAP = {
    "9": "保養維修材料", "10": "保養合約分判",
    "11": "工程材料", "12": "分判工程",
    "14": "後加工程", "4": "固定資產", "13": "其他"
};

const CURRENCY_MAP = { "0": "HKD", "1": "RMB", "2": "USD", "3": "MOP" };

/**
 * 代碼期望寫入的所有欄位（用於診斷比對）
 */
const EXPECTED_FIELDS = [
    '項目ID', '項目標題', '物業名稱', '美博報價編號',
    '項目經理', '項目內容', '採購類別', '立項預算',
    '合約金額', '幣種', '報價形式', '中標公司',
    '推薦理由', '新增時間', '修改時間'
];

/**
 * 把一個 project 轉換為飛書欄位
 * @param {object} p - 項目資料
 * @param {Set<string>} [allowedFields] - 若提供，只輸出存在於飛書的欄位
 */
function projectToFields(p, allowedFields) {
    const all = {};
    if (p.projectId) all['項目ID'] = p.projectId;
    if (p.label) all['項目標題'] = p.label;
    if (p.propertyName) all['物業名稱'] = p.propertyName;
    if (p.reportNo) all['美博報價編號'] = p.reportNo;
    if (p.manager) all['項目經理'] = p.manager;
    if (p.projectContent) all['項目內容'] = p.projectContent;
    if (p.buyType) all['採購類別'] = BUY_TYPE_MAP[p.buyType] || p.buyType;
    if (p.budget) all['立項預算'] = parseFloat(p.budget) || 0;
    if (p.contractAmount) all['合約金額'] = parseFloat(p.contractAmount) || 0;
    if (p.currency != null) all['幣種'] = CURRENCY_MAP[p.currency] || 'HKD';
    if (p.quoteType) all['報價形式'] = p.quoteType;
    if (p.winnerName) all['中標公司'] = p.winnerName;
    if (p.reason) all['推薦理由'] = p.reason;
    // 日期時間欄位：飛書使用毫秒時間戳
    if (p.createdAt) all['新增時間'] = new Date(p.createdAt).getTime();
    if (p.updatedAt) all['修改時間'] = new Date(p.updatedAt).getTime();

    // 若有 allowedFields 限制，過濾掉不存在的欄位
    if (!allowedFields) return all;
    const filtered = {};
    Object.keys(all).forEach(k => { if (allowedFields.has(k)) filtered[k] = all[k]; });
    return filtered;
}

/**
 * 批次新增（每批最多 500 條）
 */
async function batchCreate(token, appToken, tableId, records, allowedFields) {
    let created = 0;
    const batchSize = 500;
    for (let i = 0; i < records.length; i += batchSize) {
        const batch = records.slice(i, i + batchSize);
        const resp = await fetch(
            `https://open.feishu.cn/open-apis/bitable/v1/apps/${appToken}/tables/${tableId}/records/batch_create`,
            {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json; charset=utf-8'
                },
                body: JSON.stringify({ records: batch.map(p => ({ fields: projectToFields(p, allowedFields) })) })
            }
        );
        const data = await resp.json();
        if (data.code !== 0) throw new Error(`批次新增失敗(${data.code}): ${data.msg}`);
        created += (data.data?.records?.length || 0);
    }
    return created;
}

/**
 * 批次更新（每批最多 500 條）
 */
async function batchUpdate(token, appToken, tableId, records, recordIdMap, allowedFields) {
    let updated = 0;
    const batchSize = 500;
    for (let i = 0; i < records.length; i += batchSize) {
        const batch = records.slice(i, i + batchSize);
        const resp = await fetch(
            `https://open.feishu.cn/open-apis/bitable/v1/apps/${appToken}/tables/${tableId}/records/batch_update`,
            {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json; charset=utf-8'
                },
                body: JSON.stringify({
                    records: batch.map(p => ({
                        record_id: recordIdMap.get(p.projectId),
                        fields: projectToFields(p, allowedFields)
                    }))
                })
            }
        );
        const data = await resp.json();
        if (data.code !== 0) throw new Error(`批次更新失敗(${data.code}): ${data.msg}`);
        updated += (data.data?.records?.length || 0);
    }
    return updated;
}

/**
 * 把明細轉換為飛書欄位（明細表）
 */
function detailToFields(d) {
    // 初始化欄位物件
    const f = {};
    // 只寫入文字欄位（A₁），跳過關聯欄位（項目ID 為 Linked Record，不支援直接寫入文字）
    if (d.vendorName) f['承判商'] = String(d.vendorName);
    if (d.content) f['報價內容'] = String(d.content);
    if (d.detailCurrency) f['明細幣種'] = String(d.detailCurrency);
    // 金額為 Number 欄位，傳入數字
    if (d.amount != null && d.amount !== '') {
        const num = parseFloat(d.amount);
        f['金額'] = isNaN(num) ? 0 : num; // 若無法解析則傳 0
    }
    return f;
}


/**
 * 批次新增明細（每批最多 500 條）
 */
async function batchCreateDetails(token, appToken, detailTableId, details) {
    let created = 0;
    const batchSize = 500;
    for (let i = 0; i < details.length; i += batchSize) {
        const batch = details.slice(i, i + batchSize);
        const resp = await fetch(
            `https://open.feishu.cn/open-apis/bitable/v1/apps/${appToken}/tables/${detailTableId}/records/batch_create`,
            {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json; charset=utf-8'
                },
                body: JSON.stringify({ records: batch.map(d => ({ fields: detailToFields(d) })) })
            }
        );
        const data = await resp.json();
        if (data.code !== 0) throw new Error(`明細批次新增失敗(${data.code}): ${data.msg}`);
        created += (data.data?.records?.length || 0);
    }
    return created;
}

/**
 * 讀取某個表全部 record_id（自動分頁）
 */
async function fetchAllRecordIds(token, appToken, tableId) {
    const ids = [];
    let pageToken = '';
    do {
        const url = new URL(`https://open.feishu.cn/open-apis/bitable/v1/apps/${appToken}/tables/${tableId}/records`);
        url.searchParams.set('page_size', '500');
        if (pageToken) url.searchParams.set('page_token', pageToken);

        const resp = await fetch(url.toString(), {
            headers: { 'Authorization': `Bearer ${token}` }
        });
        const data = await resp.json();
        if (data.code !== 0) throw new Error(`讀取明細記錄失敗(${data.code}): ${data.msg}`);

        (data.data?.items || []).forEach(item => {
            if (item.record_id) ids.push(item.record_id);
        });
        pageToken = data.data?.has_more ? data.data.page_token : '';
    } while (pageToken);
    return ids;
}

/**
 * 批次刪除記錄（每批最多 500 條）
 */
async function batchDeleteByIds(token, appToken, tableId, recordIds) {
    let deleted = 0;
    const batchSize = 500;
    for (let i = 0; i < recordIds.length; i += batchSize) {
        const batch = recordIds.slice(i, i + batchSize);
        const resp = await fetch(
            `https://open.feishu.cn/open-apis/bitable/v1/apps/${appToken}/tables/${tableId}/records/batch_delete`,
            {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json; charset=utf-8'
                },
                body: JSON.stringify({ records: batch })
            }
        );
        const data = await resp.json();
        if (data.code !== 0) throw new Error(`明細批次刪除失敗(${data.code}): ${data.msg}`);
        deleted += batch.length;
    }
    return deleted;
}

/**
 * 主同步流程
 */
async function handleFeishuSync(config, projects, options = {}) {
    const { appId, appSecret, appToken, tableId, detailTableId } = config;
    const { downloadOnly = false } = options;

    if (!appId || !appSecret || !appToken || !tableId) {
        throw new Error('請先填寫完整的飛書設定（App ID、App Secret、App Token、Table ID）');
    }

    // 1. 取得 Token
    const token = await getFeishuToken(appId, appSecret);

    // 2. 抓取飛書表格現有欄位清單，診斷缺失欄位
    const tableFields = await fetchTableFields(token, appToken, tableId);
    const missingFields = EXPECTED_FIELDS.filter(f => !tableFields.has(f));

    // 3. 讀取飛書現有記錄
    const existingMap = await fetchAllRecords(token, appToken, tableId);

    // 4. 如果是僅下載模式，直接返回飛書資料
    if (downloadOnly) {
        const feishuProjects = await fetchAllRecordsAsProjects(token, appToken, tableId);
        return {
            created: 0, updated: 0, total: feishuProjects.length,
            downloaded: feishuProjects.length,
            feishuProjects: feishuProjects,
            missingFields: missingFields,
            detailCreated: 0, detailDeleted: 0, detailCount: 0
        };
    }

    // 5. 分類：需要新增 vs 需要更新（僅上傳模式）
    const toCreate = [];
    const toUpdate = [];
    projects.forEach(p => {
        if (!p.projectId) return; // 跳過沒有 projectId 的
        if (existingMap.has(p.projectId)) {
            toUpdate.push(p);
        } else {
            toCreate.push(p);
        }
    });

    // 6. 收集所有明細資料
    const allDetails = [];
    projects.forEach(p => {
        if (!p.projectId || !p.details) return;
        p.details.forEach(d => {
            allDetails.push({ projectId: p.projectId, ...d });
        });
    });

    // 7. 執行批次操作（傳入 tableFields 以自動過濾不存在的欄位）
    let created = 0, updated = 0;
    if (toCreate.length > 0) created = await batchCreate(token, appToken, tableId, toCreate, tableFields);
    if (toUpdate.length > 0) updated = await batchUpdate(token, appToken, tableId, toUpdate, existingMap, tableFields);

    // 8. 明細採用「覆蓋式同步」：先刪除明細表全部舊資料，再重建，避免累加
    let detailCreated = 0;
    let detailDeleted = 0;
    if (detailTableId) {
        const existingDetailIds = await fetchAllRecordIds(token, appToken, detailTableId);
        if (existingDetailIds.length > 0) {
            detailDeleted = await batchDeleteByIds(token, appToken, detailTableId, existingDetailIds);
        }
        if (allDetails.length > 0) {
            detailCreated = await batchCreateDetails(token, appToken, detailTableId, allDetails);
        }
    }

    // 9. 抓取飛書中新下載的項目（本地沒有的）
    const feishuProjects = await fetchAllRecordsAsProjects(token, appToken, tableId);
    const localProjectIds = new Set(projects.map(p => p.projectId).filter(Boolean));
    const newFromFeishu = feishuProjects.filter(fp => !localProjectIds.has(fp.projectId));

    return {
        created, updated, total: projects.length,
        downloaded: newFromFeishu.length,
        feishuProjects: newFromFeishu,
        missingFields,
        detailCreated, detailDeleted, detailCount: allDetails.length
    };
}
