/**
 * Version: 4.13.1 - Project ID + Timestamp Export + 飛書 Bitable 同步
 * 與懸浮面板 100% 統一版本的 popup.js
 */

document.addEventListener('DOMContentLoaded', async function () {
    let projects = [];
    const fields = ['label', 'propertyName', 'reportNo', 'manager', 'projectContent', 'budget', 'total', 'quoteType', 'buyType', 'currency', 'inviteCount', 'replyCount', 'reason', 'winnerName', 'contractCurrency', 'contractAmount'];
    const inputs = {};
    let lastBackup = null;
    const feishuFields = ['app-id', 'app-secret', 'app-token', 'table-id', 'detail-table-id'];
    const feishuKeys = ['oa_feishu_app_id', 'oa_feishu_app_secret', 'oa_feishu_app_token', 'oa_feishu_table_id', 'oa_feishu_detail_table_id'];

    // ===== 項目 ID 生成（方案 B：Q1322-24JC-MAT-001）=====
    const BUY_TYPE_CODE = { "9": "MNT", "10": "SVC", "11": "MAT", "12": "SUB", "14": "ADD", "4": "FAR", "13": "OTH" };
    function generateProjectId(buyType, reportNo, createdAt, allProjects) {
        const typeCode = BUY_TYPE_CODE[buyType] || 'OTH';

        // 解析報價編號，例：MS/Q1322/24/jc → Q1322-24JC
        let reportPart;
        const rn = (reportNo || '').trim();
        if (rn) {
            const parts = rn.split('/');
            if (parts.length >= 3) {
                const qNum = parts[1] || '';
                const yr = parts[2] || '';
                const init = (parts[3] || '').toUpperCase();
                reportPart = `${qNum}-${yr}${init}`;
            } else {
                // 非標準格式：清理特殊字符
                reportPart = rn.replace(/[\/\\:*?"<>| ]/g, '').toUpperCase().slice(0, 12);
            }
        } else {
            // 沒有報價編號：用日期代替
            const d = createdAt ? new Date(createdAt) : new Date();
            const pad = n => String(n).padStart(2, '0');
            reportPart = `NA-${String(d.getFullYear()).slice(2)}${pad(d.getMonth() + 1)}`;
        }

        const fullPfx = `${reportPart}-${typeCode}-`;
        let maxSeq = 0;
        (allProjects || []).forEach(p => {
            if (p.projectId && p.projectId.startsWith(fullPfx)) {
                const seq = parseInt(p.projectId.slice(fullPfx.length)) || 0;
                if (seq > maxSeq) maxSeq = seq;
            }
        });
        return fullPfx + String(maxSeq + 1).padStart(3, '0');
    }

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
        // 初始化排序選單
        const sortSel = document.getElementById('oa-sort-select');
        if (sortSel) {
            chrome.storage.local.get(['oa_setting_sort_mode'], (res) => {
                sortSel.value = res.oa_setting_sort_mode || 'default';
            });
            sortSel.onchange = () => {
                chrome.storage.local.set({ 'oa_setting_sort_mode': sortSel.value });
                const q = document.getElementById('oa-search-input')?.value.trim() || '';
                renderList(q);
            };
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
        document.getElementById('oa-btn-feishu').onclick = () => {
            // 即跳到設定頁飛書區塊
            tabSettings.click();
            setTimeout(() => {
                const el = document.getElementById('set-feishu-app-id');
                if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }, 100);
        };

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

        const exportConfirmCheck = document.getElementById('set-export-confirm');
        if (exportConfirmCheck) {
            chrome.storage.local.get(['oa_setting_export_confirm'], (res) => {
                exportConfirmCheck.checked = (res.oa_setting_export_confirm !== false);
            });
            exportConfirmCheck.onchange = () => {
                chrome.storage.local.set({ 'oa_setting_export_confirm': exportConfirmCheck.checked });
            };
        }

        // ===== 匯出子目錄 =====
        const exportDirIn = document.getElementById('set-export-dir');
        if (exportDirIn) {
            chrome.storage.local.get(['oa_setting_export_dir'], (res) => {
                exportDirIn.value = res.oa_setting_export_dir || '';
            });
            exportDirIn.addEventListener('blur', () => {
                const val = exportDirIn.value.trim().replace(/[\/:*?"<>|]/g, '');
                exportDirIn.value = val;
                chrome.storage.local.set({ 'oa_setting_export_dir': val });
                sendMessageToActiveTab({ type: "UPDATE_SETTING", key: "export_dir", value: val });
            });
        }

        // ===== 飛書設定讀寫（local 為主 + sync 備援，避免更新後需重填）=====
        const feishuInputs = feishuFields.map(field => document.getElementById('set-feishu-' + field));
        let feishuSaveTimer = null;

        const collectFeishuValues = () => {
            const obj = {};
            feishuKeys.forEach((key, i) => {
                obj[key] = (feishuInputs[i]?.value || '').trim();
            });
            return obj;
        };

        const fillFeishuInputs = (values) => {
            feishuKeys.forEach((key, i) => {
                if (feishuInputs[i]) feishuInputs[i].value = values[key] || '';
            });
        };

        const persistFeishuValues = (values) => {
            chrome.storage.local.set(values);
            chrome.storage.sync.set(values);
        };

        const schedulePersistFeishuValues = () => {
            if (feishuSaveTimer) clearTimeout(feishuSaveTimer);
            feishuSaveTimer = setTimeout(() => {
                persistFeishuValues(collectFeishuValues());
                feishuSaveTimer = null;
            }, 300);
        };

        const loadFeishuValues = async () => {
            const [localRes, syncRes] = await Promise.all([
                chrome.storage.local.get(feishuKeys),
                chrome.storage.sync.get(feishuKeys)
            ]);
            const merged = {};
            feishuKeys.forEach((key) => {
                const localVal = (localRes[key] || '').trim();
                const syncVal = (syncRes[key] || '').trim();
                merged[key] = localVal || syncVal || '';
            });
            fillFeishuInputs(merged);

            // 自動雙向補齊，避免不同裝置/區域造成資料缺口
            if (feishuKeys.some(k => merged[k])) persistFeishuValues(merged);
        };

        loadFeishuValues();

        feishuInputs.forEach((el) => {
            if (!el) return;
            el.addEventListener('input', schedulePersistFeishuValues);
            el.addEventListener('change', schedulePersistFeishuValues);
            el.addEventListener('blur', () => persistFeishuValues(collectFeishuValues()));
        });
        window.addEventListener('beforeunload', () => {
            if (feishuSaveTimer) clearTimeout(feishuSaveTimer);
            persistFeishuValues(collectFeishuValues());
        });

        // ===== 飛書同步按鈕 =====
        const feishuSyncBtn = document.getElementById('btn-feishu-sync');
        const feishuStatus = document.getElementById('feishu-sync-status');

        function showFeishuStatus(type, html) {
            feishuStatus.style.display = 'block';
            if (type === 'loading') {
                feishuStatus.style.background = 'rgba(51,112,255,0.08)';
                feishuStatus.style.color = '#3370FF';
                feishuStatus.style.border = '1px solid rgba(51,112,255,0.2)';
            } else if (type === 'success') {
                feishuStatus.style.background = 'rgba(52,199,89,0.08)';
                feishuStatus.style.color = '#1a8a35';
                feishuStatus.style.border = '1px solid rgba(52,199,89,0.3)';
            } else {
                feishuStatus.style.background = 'rgba(255,59,48,0.08)';
                feishuStatus.style.color = '#c0392b';
                feishuStatus.style.border = '1px solid rgba(255,59,48,0.2)';
            }
            feishuStatus.innerHTML = html;
        }

        if (feishuSyncBtn) {
            feishuSyncBtn.onclick = async () => {
                // 優先讀取輸入框當前值，避免「尚未 blur 就按同步」丟失更新
                const liveCfg = collectFeishuValues();
                persistFeishuValues(liveCfg);
                const config = {
                    appId: liveCfg.oa_feishu_app_id || '',
                    appSecret: liveCfg.oa_feishu_app_secret || '',
                    appToken: liveCfg.oa_feishu_app_token || '',
                    tableId: liveCfg.oa_feishu_table_id || '',
                    detailTableId: liveCfg.oa_feishu_detail_table_id || ''
                };

                if (!config.appId || !config.appSecret || !config.appToken || !config.tableId) {
                    showFeishuStatus('error', '❌ 請先填寫完整的飛書設定（App ID、App Secret、App Token、Table ID）');
                    return;
                }

                if (projects.length === 0) {
                    showFeishuStatus('error', '❌ 尚無項目資料可同步');
                    return;
                }

                feishuSyncBtn.disabled = true;
                feishuSyncBtn.textContent = '⏳ 同步中...';
                showFeishuStatus('loading', `🐦 正在連線飛書並同步 ${projects.length} 個項目，請稍候...`);

                try {
                    const result = await new Promise((resolve, reject) => {
                        chrome.runtime.sendMessage(
                            { type: 'FEISHU_SYNC', config: config, projects: projects },
                            (res) => {
                                if (chrome.runtime.lastError) reject(new Error(chrome.runtime.lastError.message));
                                else if (!res.success) reject(new Error(res.error || '未知錯誤'));
                                else resolve(res);
                            }
                        );
                    });
                    // 組合成功訊息
                    let successHtml =
                        `✅ <strong>同步成功</strong><br>` +
                        `📥 新增：${result.created} 條` +
                        `&emsp;✏️ 更新：${result.updated} 條` +
                        `&emsp;📄 共計：${result.total} 條`;

                    // 若有從飛書下載的新項目
                    if (result.downloaded > 0) {
                        successHtml += `<br>📥 從飛書下載：${result.downloaded} 條（新項目已合併到本地）`;
                    }

                    // 若有明細寫入，附加明細統計
                    if (result.detailCreated > 0) {
                        successHtml += `<br>🔸 明細新增：${result.detailCreated} 條（共 ${result.detailCount} 條明細）`;
                    } else if (result.detailCount > 0) {
                        successHtml += `<br>⚠️ 有 ${result.detailCount} 條明細但未寫入（請確認明細表 Table ID 設定）`;
                    }
                    if (result.detailDeleted > 0) {
                        successHtml += `<br>🧹 已刪除舊明細：${result.detailDeleted} 條（覆蓋式同步）`;
                    }

                    // 若有缺失欄位，附加黃色警告
                    if (result.missingFields && result.missingFields.length > 0) {
                        successHtml +=
                            `<div style="margin-top:10px; padding:8px 10px; background:rgba(255,204,0,0.12);
                             border:1px solid rgba(255,170,0,0.4); border-radius:8px; color:#7a5c00; font-size:11px; line-height:1.8;">` +
                            `⚠️ <strong>以下 ${result.missingFields.length} 個欄位在飛書表格中找不到</strong>，相關數據未被寫入：<br>` +
                            result.missingFields.map(f => `&nbsp;• ${f}`).join('<br>') +
                            `<br><span style="opacity:0.75">▶ 請在多維表格中新增對應欄位後重新同步</span></div>`;
                    }
                    showFeishuStatus('success', successHtml);

                    // 如果有從飛書下載的新項目，自動合併到本地存儲
                    if (result.feishuProjects && result.feishuProjects.length > 0) {
                        const mergedProjects = [...projects, ...result.feishuProjects];
                        await chrome.storage.local.set({ 'oa_projects_v13': mergedProjects });
                        syncData(); // 重新讀取並渲染列表
                    }
                } catch (err) {
                    showFeishuStatus('error', `❌ <strong>同步失敗</strong><br>${err.message}`);
                } finally {
                    feishuSyncBtn.disabled = false;
                    feishuSyncBtn.textContent = '🐦 立即同步到飛書';
                }
            };
        }

        document.getElementById('btn-f-pick').onclick = () => {
            alert("⚠️ 彈窗介面無法直接讀取頁面中的人員選擇器，請在網頁上的 OA 懸浮面板使用此功能，或手動輸入經理名稱。");
        };

        document.getElementById('btn-add-detail').onclick = () => addDetailRow();

        function buildCsvDataUrl() {
            const buyTypeMap = {
                "9": "保養維修材料",
                "10": "保養合約分判",
                "11": "工程材料",
                "12": "分判工程",
                "14": "後加工程",
                "4": "固定資產",
                "13": "其他"
            };
            const h = ["項目標題", "物業名稱", "報價編號", "項目經理", "項目內容", "立項預算", "合約總價", "報價形式", "採購類別", "幣種", "承判商", "報價內容", "明細幣種", "金額", "邀請公司", "有效報價", "推荐理由", "中標公司", "合約幣種", "合約金額", "項目ID", "新增日期時間", "修改日期時間"];
            const rows = [];
            projects.forEach(p => {
                const buyTypeText = buyTypeMap[p.buyType] || p.buyType || "";
                (p.details || [{}]).forEach(dt => {
                    const row = [p.label, p.propertyName, p.reportNo, p.manager, p.projectContent, p.budget, p.total, p.quoteType, buyTypeText, p.currency, dt.vendorName, dt.content, dt.detailCurrency, dt.amount, p.inviteCount, p.replyCount, p.reason, p.winnerName, p.contractCurrency, p.contractAmount,
                    p.projectId || '', p.createdAt || '', p.updatedAt || ''];
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
            }, () => { });
        }

        async function exportToCSV() {
            if (projects.length === 0) return alert("尚無數據可匯出");

            const setting = await chrome.storage.local.get(['oa_setting_export_confirm', 'oa_setting_export_dir']);
            const needsConfirm = (setting.oa_setting_export_confirm !== false);
            const exportDir = (setting.oa_setting_export_dir || "").trim();
            const basePath = exportDir ? (exportDir + '/') : '';

            if (!needsConfirm) {
                triggerDownload(basePath + 'OA_Projects.csv', 'overwrite');
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
            document.getElementById('pop-exp-overwrite').onclick = () => { close(); triggerDownload(basePath + 'OA_Projects.csv', 'overwrite'); };
            document.getElementById('pop-exp-new').onclick = () => {
                close();
                const ts = new Date().toISOString().slice(0, 16).replace('T', '_').replace(':', '-');
                triggerDownload(basePath + `OA_Projects_${ts}.csv`, 'uniquify');
            };
        }

        document.getElementById('btn-f-save').onclick = async () => {
            const id = document.getElementById('in-f-id').value || "p_" + Date.now();
            const now = new Date().toISOString();
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
            if (idx > -1) {
                // 更新現有項目：保留原 createdAt 和 projectId，更新 updatedAt
                p.createdAt = projects[idx].createdAt || now;
                p.updatedAt = now;
                p.projectId = projects[idx].projectId || generateProjectId(p.buyType, p.reportNo, p.createdAt, projects);
                projects[idx] = p;
            } else {
                // 新增項目：生成 projectId
                p.createdAt = now;
                p.updatedAt = now;
                projects.push(p); // 先加入再生成，避免編號重複
                p.projectId = generateProjectId(p.buyType, p.reportNo, p.createdAt, projects);
            }
            // 更新表單顯示
            const pidDisplay = document.getElementById('in-f-projectId-display');
            const pidHidden = document.getElementById('in-f-projectId');
            if (pidDisplay) pidDisplay.textContent = p.projectId;
            if (pidHidden) pidHidden.value = p.projectId;
            await chrome.storage.local.set({ 'oa_projects_v13': projects });
            alert("✅ 保存成功及自動匯出"); exportToCSV(); syncData(); tabFill.click(); resetForm();
        };

        document.getElementById('btn-f-save-as-new').onclick = async () => {
            const now = new Date().toISOString();
            const p = { id: "p_" + Date.now(), managerId: document.getElementById('in-f-managerId').value, details: [], createdAt: now, updatedAt: now };
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
            p.projectId = generateProjectId(p.buyType, p.reportNo, p.createdAt, projects);
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
                        const importTime = new Date().toISOString();
                        // 讀取項目 ID 和時間戳（如果 CSV 有這些欄）
                        const csvProjectId = (cols[20] || '').trim();
                        const csvCreatedAt = (cols[21] || '').trim();
                        const csvUpdatedAt = (cols[22] || '').trim();
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
                            contractAmount: (cols[19] || "").trim(),
                            projectId: csvProjectId || '',
                            createdAt: csvCreatedAt || importTime,
                            updatedAt: csvUpdatedAt || importTime
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

        // 自動遷移：為沒有時間戳的舊項目補上時間
        let needsSave = false;
        projects.forEach(p => {
            if (!p.createdAt) {
                // 嘗試從 id（格式 "p_1234567890"）反推時間
                const tsMatch = p.id && p.id.match(/p_(\d+)/);
                const ts = tsMatch ? parseInt(tsMatch[1]) : null;
                const guessedTime = (ts && ts > 1000000000000) ? new Date(ts).toISOString() : null;
                p.createdAt = guessedTime || new Date().toISOString();
                p.updatedAt = p.createdAt;
                needsSave = true;
            }
        });

        // 自動遷移：為沒有 projectId 的舊項目自動補填 ID
        // 順序處理，使序號能正確逐步累加
        projects.forEach(p => {
            if (!p.projectId) {
                p.projectId = generateProjectId(p.buyType, p.reportNo, p.createdAt, projects);
                needsSave = true;
            }
        });

        if (needsSave) {
            await chrome.storage.local.set({ 'oa_projects_v13': projects });
        }

        const searchInput = document.getElementById('oa-search-input');
        renderList(searchInput ? searchInput.value.trim() : "");
    }

    async function renderList(query = "") {
        const list = document.getElementById('oa-disp-list'); if (!list) return;

        // 讀取設置
        const settings = await chrome.storage.local.get([
            'oa_setting_title_size',
            'oa_setting_subtitle_mode',
            'oa_setting_compact',
            'oa_setting_sort_mode'
        ]);

        const titleSize = settings.oa_setting_title_size || 15;
        const subMode = settings.oa_setting_subtitle_mode || 'both';
        const isCompact = !!settings.oa_setting_compact;
        const sortMode = settings.oa_setting_sort_mode || 'default';

        // 同步排序選單顯示
        const sortSel = document.getElementById('oa-sort-select');
        if (sortSel && sortSel.value !== sortMode) sortSel.value = sortMode;

        let filtered = projects;
        if (query) {
            const q = query.toLowerCase();
            filtered = projects.filter(p =>
                (p.label && p.label.toLowerCase().includes(q)) ||
                (p.manager && p.manager.toLowerCase().includes(q)) ||
                (p.details && p.details.some(dt => dt.vendorName && dt.vendorName.toLowerCase().includes(q)))
            );
        }

        // ===== 排序邏輯 =====
        const getTime = (p, field) => p[field] || p.createdAt || '';
        if (sortMode !== 'default') {
            filtered = [...filtered];
            if (sortMode === 'created_desc') filtered.sort((a, b) => getTime(b, 'createdAt') > getTime(a, 'createdAt') ? 1 : -1);
            if (sortMode === 'created_asc') filtered.sort((a, b) => getTime(a, 'createdAt') > getTime(b, 'createdAt') ? 1 : -1);
            if (sortMode === 'updated_desc') filtered.sort((a, b) => getTime(b, 'updatedAt') > getTime(a, 'updatedAt') ? 1 : -1);
            if (sortMode === 'updated_asc') filtered.sort((a, b) => getTime(a, 'updatedAt') > getTime(b, 'updatedAt') ? 1 : -1);
        }

        list.innerHTML = filtered.length === 0 ? `<p style="padding:40px; color:#999; text-align:center;">${query ? '找不到匹配項目' : '尚無記錄'}</p>` : '';

        const grouped = {};
        const unGrouped = [];
        filtered.forEach(p => {
            const rNo = (p.reportNo || '').trim();
            if (rNo) {
                if (!grouped[rNo]) grouped[rNo] = [];
                grouped[rNo].push(p);
            } else {
                unGrouped.push(p);
            }
        });

        // 格式化日期時間顯示（短格式：MM-DD HH:mm）
        function formatDateTime(isoStr) {
            if (!isoStr) return '';
            try {
                const d = new Date(isoStr);
                const pad = n => String(n).padStart(2, '0');
                return `${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
            } catch (e) { return ''; }
        }

        const createItemEl = (p) => {
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

            // 構建日期時間行（獨立第三行）
            const createdStr = p.createdAt ? formatDateTime(p.createdAt) : '';
            const updatedStr = p.updatedAt ? formatDateTime(p.updatedAt) : '';
            const isModified = p.createdAt && p.updatedAt && p.createdAt !== p.updatedAt;
            let dateHtml = '';
            if (createdStr) {
                dateHtml = `<div style="font-size:10px; color:#b0b0ba; margin-top:2px;">🕐 ${createdStr}`;
                if (isModified) dateHtml += `　✏️ ${updatedStr}`;
                dateHtml += `</div>`;
            }

            div.innerHTML = `
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <div style="flex:1; overflow:hidden;">
                        <div style="font-weight:700; font-size:${titleSize}px; color:#1d1d1f; margin-bottom:4px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${p.label}</div>
                        <div style="font-size:11px; color:#86868b; letter-spacing:0.5px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${subContent}</div>
                        ${dateHtml}
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
                // 更新項目ID顯示
                const pidDisp = document.getElementById('in-f-projectId-display');
                const pidHid = document.getElementById('in-f-projectId');
                if (pidDisp) pidDisp.textContent = p.projectId || '（儲存後生成）';
                if (pidHid) pidHid.value = p.projectId || '';
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
            return div;
        };

        // ===== 資料夾排序 =====
        const getFolderTime = (rNo, field) =>
            grouped[rNo].reduce((mx, p) => { const t = p[field] || p.createdAt || ''; return t > mx ? t : mx; }, '');

        let folderKeys = Object.keys(grouped);
        if (sortMode === 'created_desc') folderKeys.sort((a, b) => getFolderTime(b, 'createdAt') > getFolderTime(a, 'createdAt') ? 1 : -1);
        if (sortMode === 'created_asc') folderKeys.sort((a, b) => getFolderTime(a, 'createdAt') > getFolderTime(b, 'createdAt') ? 1 : -1);
        if (sortMode === 'updated_desc') folderKeys.sort((a, b) => getFolderTime(b, 'updatedAt') > getFolderTime(a, 'updatedAt') ? 1 : -1);
        if (sortMode === 'updated_asc') folderKeys.sort((a, b) => getFolderTime(a, 'updatedAt') > getFolderTime(b, 'updatedAt') ? 1 : -1);

        folderKeys.forEach(rNo => {
            const groupItems = grouped[rNo];
            const folder = document.createElement('div');
            // 預設為收合
            folder.className = 'oa-folder';
            folder.innerHTML = `
                <div class="oa-folder-header">
                    <div class="oa-folder-title">
                        <span>📁</span>
                        <span>${rNo}</span>
                        <span class="oa-folder-count">${groupItems.length}</span>
                    </div>
                    <span class="oa-folder-arrow">▼</span>
                </div>
                <div class="oa-folder-content"></div>
            `;
            folder.querySelector('.oa-folder-header').onclick = () => folder.classList.toggle('open');
            const content = folder.querySelector('.oa-folder-content');
            groupItems.forEach(p => content.appendChild(createItemEl(p)));
            list.appendChild(folder);
        });

        unGrouped.forEach(p => list.appendChild(createItemEl(p)));
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
