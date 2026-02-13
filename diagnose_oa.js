/**
 * 診斷腳本：請在 OA 頁面按 F12，貼上以下代碼並執行
 * 它會列出所有可能的欄位 ID 及其當前值
 */
(function () {
    const results = [];
    const inputs = document.querySelectorAll('input, textarea');
    inputs.forEach(i => {
        if (i.value && i.value.trim().length > 0) {
            results.push({
                id: i.id,
                name: i.name,
                value: i.value,
                placeholder: i.placeholder
            });
        }
    });
    console.table(results);
    alert("掃描完成，請查看 Console 控制台 (F12) 的表格。");
})();
