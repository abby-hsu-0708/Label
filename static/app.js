let detectedPdfFile = null;
let detectedExcelFile = null;
let currentExcelFile = null; // 用於重新生成代碼

document.addEventListener("DOMContentLoaded", () => {
    initPreviewUploader();
    initUnifiedUploader();
    initDateChangeHandlers();
    initCopyCode();
});

// 1. 初始化左欄：預覽標籤箱數
function initPreviewUploader() {
    const dropzone = document.getElementById("preview-dropzone");
    const input = document.getElementById("preview-file-input");

    setupDropEvents(dropzone, (files) => {
        if (files.length > 0 && isExcel(files[0])) {
            handlePreviewExcel(files[0]);
        }
    });

    input.addEventListener("change", (e) => {
        if (e.target.files.length > 0) {
            handlePreviewExcel(e.target.files[0]);
        }
    });
}

// 2. 初始化右欄：智慧標籤自動套印 (單一入口)
function initUnifiedUploader() {
    const dropzone = document.getElementById("unified-dropzone");
    const input = document.getElementById("unified-file-input");
    const manualBtn = document.getElementById("manual-run-btn");
    const autoRunCb = document.getElementById("auto-run-cb");

    setupDropEvents(dropzone, (files) => {
        handleIncomingFiles(files);
    });

    input.addEventListener("change", (e) => {
        handleIncomingFiles(e.target.files);
    });

    autoRunCb.addEventListener("change", () => {
        checkAndTriggerProcess();
    });

    manualBtn.addEventListener("click", () => {
        executeLabelProcess();
    });
}

function handleIncomingFiles(fileList) {
    for (const f of fileList) {
        const fn = f.name.toLowerCase();
        if (fn.endsWith(".pdf") && !detectedPdfFile) {
            detectedPdfFile = f;
            updateBadge("pdf-badge", "pdf-badge-title", "pdf-badge-desc", `Label PDF: ${f.name}`, `${(f.size/1024).toFixed(1)} KB`, true);
        } else if ((fn.endsWith(".xls") || fn.endsWith(".xlsx")) && !detectedExcelFile) {
            detectedExcelFile = f;
            updateBadge("excel-badge", "excel-badge-title", "excel-badge-desc", `Packing List: ${f.name}`, `${(f.size/1024).toFixed(1)} KB`, true);
            // 同時解析 Excel 供方案 A 顯示
            handlePreviewExcel(f);
        }
    }
    checkAndTriggerProcess();
}

function updateBadge(badgeId, titleId, descId, title, desc, ready) {
    const badge = document.getElementById(badgeId);
    badge.className = `file-badge ${ready ? 'ready' : 'waiting'}`;
    document.getElementById(titleId).textContent = title;
    document.getElementById(descId).textContent = desc;
}

function checkAndTriggerProcess() {
    const autoRun = document.getElementById("auto-run-cb").checked;
    const manualBtn = document.getElementById("manual-run-btn");

    if (detectedPdfFile && detectedExcelFile) {
        if (autoRun) {
            manualBtn.style.display = "none";
            executeLabelProcess();
        } else {
            manualBtn.style.display = "inline-flex";
        }
    } else {
        manualBtn.style.display = "none";
    }
}

// 3. 處理 Excel 預覽解析與方案 A
async function handlePreviewExcel(file) {
    currentExcelFile = file;
    document.getElementById("preview-filename").textContent = file.name;
    const statusMsg = document.getElementById("preview-status");
    statusMsg.textContent = "正在解析裝箱明細中...";

    const formData = new FormData();
    formData.append("excel_file", file);

    try {
        const res = await fetch("/api/preview_excel", {
            method: "POST",
            body: formData
        });
        const data = await res.json();

        if (data.success) {
            statusMsg.textContent = `✅ 已解析 ${data.po_rows.length} 組 PO，共 ${data.total_boxes} 箱。`;
            renderPlanA(data);
            renderPOTable(data.po_rows);
        } else {
            statusMsg.textContent = `❌ ${data.message || '解析失敗'}`;
        }
    } catch (err) {
        statusMsg.textContent = `❌ 連線伺服器出錯: ${err.message}`;
    }
}

function renderPlanA(data) {
    const planA = document.getElementById("plan-a-section");
    planA.style.display = "block";

    // 格式化 YYYY/MM/DD 為 YYYY-MM-DD 供 date input 賦值
    const startVal = data.start_date.replace(/\//g, "-");
    const endVal = data.end_date.replace(/\//g, "-");

    document.getElementById("date-start").value = startVal;
    document.getElementById("date-end").value = endVal;

    updateBookmarkletUI(data.bookmarklet_url, data.raw_js);
}

function updateBookmarkletUI(bookmarkletUrl, rawJs) {
    const link = document.getElementById("bookmarklet-link");
    link.href = bookmarkletUrl;
    link.onclick = (e) => {
        // 防止使用者直接點擊跳轉，引導其拖曳
        alert("請用滑鼠按住此按鈕【拖曳至瀏覽器書籤列】！\n在客戶出貨管理頁點一下即可一鍵全自動填表。");
        e.preventDefault();
    };

    document.getElementById("js-code-display").textContent = rawJs;
}

function initDateChangeHandlers() {
    const startInput = document.getElementById("date-start");
    const endInput = document.getElementById("date-end");

    const refreshScript = async () => {
        if (!currentExcelFile) return;
        const sVal = startInput.value.replace(/-/g, "/");
        const eVal = endInput.value.replace(/-/g, "/");

        const formData = new FormData();
        formData.append("start_date", sVal);
        formData.append("end_date", eVal);
        formData.append("excel_file", currentExcelFile);

        try {
            const res = await fetch("/api/generate_script", { method: "POST", body: formData });
            const data = await res.json();
            if (data.success) {
                updateBookmarkletUI(data.bookmarklet_url, data.raw_js);
            }
        } catch (e) {
            console.error("更新代碼失敗:", e);
        }
    };

    startInput.addEventListener("change", refreshScript);
    endInput.addEventListener("change", refreshScript);
}

function renderPOTable(rows) {
    const section = document.getElementById("po-table-section");
    section.style.display = "block";
    document.getElementById("po-count-tag").textContent = `已解析 ${rows.length} 組 PO`;

    const tbody = document.getElementById("po-table-body");
    tbody.innerHTML = "";

    rows.forEach(r => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td style="font-weight: 600; color: #a78bfa;">${r.po}</td>
            <td>${r.product_no || '-'}</td>
            <td style="color: #cbd5e1; font-size: 0.82rem;">${r.description || '-'}</td>
            <td style="text-align: right; font-weight: 600; color: #38bdf8;">${r.box_count} 箱</td>
            <td style="text-align: right; color: #94a3b8;">${r.total_qty} pcs</td>
        `;
        tbody.appendChild(tr);
    });
}

// 4. 執行核心 PDF 套印
async function executeLabelProcess() {
    if (!detectedPdfFile || !detectedExcelFile) return;

    const overlay = document.getElementById("loading-overlay");
    overlay.style.display = "flex";

    const formData = new FormData();
    formData.append("pdf_file", detectedPdfFile);
    formData.append("excel_file", detectedExcelFile);

    try {
        const res = await fetch("/api/process", {
            method: "POST",
            body: formData
        });
        const data = await res.json();
        overlay.style.display = "none";

        if (data.success) {
            renderResult(data);
        } else {
            alert(`❌ 處理失敗: ${data.message || '未知錯誤'}`);
        }
    } catch (err) {
        overlay.style.display = "none";
        alert(`❌ 伺服器錯誤: ${err.message}`);
    }
}

function renderResult(data) {
    const resultSec = document.getElementById("result-section");
    resultSec.style.display = "block";

    const stats = data.stats || {};
    document.getElementById("stat-pages").textContent = `${stats.total_pages || 0} 頁`;
    document.getElementById("stat-matched").textContent = `${stats.matched_count || 0} 個`;
    document.getElementById("stat-missing").textContent = `${stats.missing_labels || 0} 個`;
    document.getElementById("stat-unmatched").textContent = `${stats.unmatched_labels || 0} 個`;

    const warnBox = document.getElementById("warning-box");
    const warnList = document.getElementById("warning-list");
    warnList.innerHTML = "";
    if (stats.warnings && stats.warnings.length > 0) {
        warnBox.style.display = "block";
        stats.warnings.forEach(w => {
            const li = document.createElement("li");
            li.textContent = w;
            warnList.appendChild(li);
        });
    } else {
        warnBox.style.display = "none";
    }

    if (data.original_preview) {
        document.getElementById("img-original").src = data.original_preview;
    }
    if (data.labeled_preview) {
        document.getElementById("img-labeled").src = data.labeled_preview;
    }

    const dlBtn = document.getElementById("download-btn");
    dlBtn.href = `/api/download/${data.task_id}`;

    // 平滑滾動至成果區域
    resultSec.scrollIntoView({ behavior: "smooth" });
}

// 5. 複製代碼小功能
function initCopyCode() {
    const copyBtn = document.getElementById("copy-code-btn");
    copyBtn.addEventListener("click", () => {
        const codeText = document.getElementById("js-code-display").textContent;
        navigator.clipboard.writeText(codeText).then(() => {
            copyBtn.textContent = "✅ 已複製！";
            setTimeout(() => { copyBtn.textContent = "📋 複製代碼"; }, 2500);
        }).catch(err => {
            alert("複製失敗，請手動全選複製。");
        });
    });
}

function setupDropEvents(dropzone, onFiles) {
    ['dragenter', 'dragover'].forEach(eventName => {
        dropzone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropzone.classList.add('dragover');
        });
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropzone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropzone.classList.remove('dragover');
        });
    });

    dropzone.addEventListener('drop', (e) => {
        if (e.dataTransfer && e.dataTransfer.files) {
            onFiles(Array.from(e.dataTransfer.files));
        }
    });
}

function isExcel(file) {
    const fn = file.name.toLowerCase();
    return fn.endsWith(".xls") || fn.endsWith(".xlsx");
}
