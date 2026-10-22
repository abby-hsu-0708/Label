import streamlit as st
import os
import fitz  # PyMuPDF
import tempfile
import base64
import re
import pandas as pd
import importlib
import core_logic
importlib.reload(core_logic)
from core_logic import parse_excel, label_pdf, find_delivery_date, calculate_date_range, generate_portal_bookmarklet


# 1. 設置頁面配置
st.set_page_config(page_title="智慧出貨標籤自動套印系統", layout="centered", page_icon="🏷️")

# 2. 注入奢華的極光與玻璃擬物化 (Glassmorphism) 全域 CSS
st.markdown("""
    <style>
    /* 全域字型與背景 */
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;700&family=Noto+Sans+TC:wght@300;400;500;700&display=swap');
    
    .stApp {
        background-color: #121826 !important;
        color: #f3f4f6 !important;
        font-family: 'Outfit', 'Noto Sans TC', sans-serif !important;
    }

    /* 書籤按鈕專屬樣式 */
    .bookmarklet-btn {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        background: linear-gradient(135deg, #8b5cf6, #3b82f6);
        color: #ffffff !important;
        font-weight: 600;
        font-size: 0.95rem;
        padding: 10px 20px;
        border-radius: 12px;
        text-decoration: none !important;
        box-shadow: 0 4px 14px rgba(139, 92, 246, 0.35);
        cursor: grab;
        transition: all 0.2s ease;
        border: 1px solid rgba(255, 255, 255, 0.15);
    }
    .bookmarklet-btn:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(139, 92, 246, 0.55);
        color: #ffffff !important;
        background: linear-gradient(135deg, #9d76fa, #4f94fd);
    }
    .bookmarklet-btn:active {
        cursor: grabbing;
        transform: translateY(0);
    }
    
    /* 調整 st.expander 樣式，防止背景在點擊或某些狀態下變白 */
    div[data-testid="stExpander"] {
        background: rgba(17, 24, 39, 0.45) !important;
        border: 1px solid rgba(255, 255, 255, 0.08) !important;
        border-radius: 16px !important;
        overflow: hidden !important;
    }
    
    div[data-testid="stExpander"] > details {
        border: none !important;
        background: transparent !important;
    }
    
    div[data-testid="stExpander"] > details > summary {
        background-color: rgba(17, 24, 39, 0.7) !important;
        color: #f3f4f6 !important;
        border-bottom: 1px solid rgba(255, 255, 255, 0.04) !important;
        padding: 12px 16px !important;
        transition: background-color 0.2s ease, color 0.2s ease !important;
    }
    
    div[data-testid="stExpander"] > details > summary:hover {
        background-color: rgba(17, 24, 39, 0.9) !important;
        color: #ffffff !important;
    }
    
    div[data-testid="stExpander"] > details > summary:focus,
    div[data-testid="stExpander"] > details > summary:active,
    div[data-testid="stExpander"] > details[open] > summary {
        background-color: rgba(17, 24, 39, 0.8) !important;
        color: #ffffff !important;
    }
    
    div[data-testid="stExpander"] > details > summary * {
        color: #f3f4f6 !important;
    }
    div[data-testid="stExpander"] > details > summary:hover * {
        color: #ffffff !important;
    }
    
    /* 放大主容器寬度，讓左右欄的方塊拉寬變大 */
    .block-container {
        max-width: 950px !important;
        padding-left: 2rem !important;
        padding-right: 2rem !important;
    }
    
    /* 隱藏 Streamlit 預設的頁首與頁尾 */
    header { visibility: hidden !important; }
    footer { visibility: hidden !important; }
    #MainMenu { visibility: hidden !important; }
    
    /* 極光背景發光效果 */
    .stApp::before {
        content: "";
        position: absolute;
        width: 600px;
        height: 600px;
        border-radius: 50%;
        background: #8b5cf6;
        filter: blur(140px);
        opacity: 0.12;
        top: -150px;
        left: -150px;
        pointer-events: none;
        z-index: 0;
    }
    
    .stApp::after {
        content: "";
        position: absolute;
        width: 600px;
        height: 600px;
        border-radius: 50%;
        background: #3b82f6;
        filter: blur(140px);
        opacity: 0.12;
        bottom: -150px;
        right: -150px;
        pointer-events: none;
        z-index: 0;
    }

    /* 頂部標題客製化樣式 */
    .custom-header {
        text-align: center;
        margin-bottom: 35px;
        z-index: 10;
        position: relative;
    }
    .custom-logo {
        display: inline-flex;
        align-items: center;
        gap: 12px;
        margin-bottom: 12px;
    }
    .custom-logo h1 {
        font-size: 2.2rem !important;
        font-weight: 700 !important;
        margin: 0 !important;
        background: linear-gradient(135deg, #fff 40%, #9ca3af);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .custom-badge {
        background: linear-gradient(135deg, #8b5cf6, #3b82f6);
        font-size: 0.75rem;
        font-weight: 600;
        padding: 3px 8px;
        border-radius: 20px;
        color: white;
        box-shadow: 0 0 15px rgba(139, 92, 246, 0.4);
    }
    .custom-subtitle {
        color: #9ca3af;
        font-size: 1rem;
        font-weight: 300;
        max-width: 600px;
        margin: 0 auto;
        line-height: 1.6;
    }

    /* 覆寫 Streamlit 原生上傳區域，使其變身為玻璃擬物化卡片 */
    /* 覆寫 Streamlit 原生上傳區域，使其變身為精緻緊湊的扁平科技感卡片 */
    div[data-testid="stFileUploader"] {
        background: #090d16 !important;
        border: 1px solid rgba(255, 255, 255, 0.08) !important;
        border-radius: 16px !important;
        padding: 12px 16px !important;
        backdrop-filter: blur(20px) !important;
        -webkit-backdrop-filter: blur(20px) !important;
        box-shadow: 0 8px 18px rgba(0, 0, 0, 0.3) !important;
        transition: all 0.25s ease !important;
        margin-bottom: 12px !important;
        margin-left: 0px !important;
        margin-right: 0px !important;
        min-height: auto !important;
        display: flex !important;
        flex-direction: column !important;
        justify-content: flex-start !important;
    }
    
    div[data-testid="stFileUploader"] > div,
    div[data-testid="stFileUploader"] > div > div,
    div[data-testid="stFileUploader"] > div > div > div {
        width: 100% !important;
        max-width: 100% !important;
    }
    
    div[data-testid="stFileUploader"] label {
        min-height: auto !important;
        margin-bottom: 6px !important;
        display: inline-flex !important;
        align-items: center !important;
    }
    
    div[data-testid="stFileUploader"] label p {
        color: #f3f4f6 !important;
        font-weight: 600 !important;
        font-size: 0.92rem !important;
        margin: 0 !important;
    }
    
    div[data-testid="stFileUploader"]:hover {
        border-color: rgba(255, 255, 255, 0.2) !important;
        box-shadow: 0 10px 22px rgba(0, 0, 0, 0.4) !important;
    }

    /* 上傳框內部 Dropzone - 改為緊湊的白底方塊 */
    div[data-testid="stFileUploaderDropzone"] {
        background-color: #f8fafc !important;
        border: 1px dashed #cbd5e1 !important;
        border-radius: 10px !important;
        height: auto !important;
        min-height: 48px !important;
        width: 100% !important;
        max-width: 100% !important;
        display: flex !important;
        flex-direction: row !important;
        justify-content: space-between !important;
        align-items: center !important;
        padding: 6px 14px !important;
    }
    
    div[data-testid="stFileUploaderDropzone"] > div {
        background-color: transparent !important;
        display: flex !important;
        flex-direction: row !important;
        align-items: center !important;
        justify-content: flex-start !important;
        gap: 10px !important;
        width: auto !important;
    }
    
    div[data-testid="stFileUploaderDropzone"] h3, 
    div[data-testid="stFileUploaderDropzone"] span, 
    div[data-testid="stFileUploaderDropzone"] small,
    div[data-testid="stFileUploaderDropzone"] p {
        color: #334155 !important;
        font-weight: 500 !important;
        font-size: 0.82rem !important;
        margin: 0 !important;
    }

    /* 針對 Streamlit 內置的 "Upload" / "Browse files" 按鈕 */
    div[data-testid="stFileUploaderDropzone"] button,
    div[data-testid="stFileUploader"] button,
    .stFileUploader button {
        background-color: #000000 !important;
        color: #ffffff !important;
        border: 1px solid rgba(255, 255, 255, 0.3) !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        font-size: 0.82rem !important;
        padding: 5px 12px !important;
        transition: all 0.2s ease !important;
    }

    /* 強制按鈕內部的 text span 與 icon svg 顏色為純白色 */
    div[data-testid="stFileUploaderDropzone"] button *,
    div[data-testid="stFileUploader"] button *,
    .stFileUploader button * {
        color: #ffffff !important;
        fill: #ffffff !important;
        stroke: #ffffff !important;
    }

    div[data-testid="stFileUploaderDropzone"] button:hover,
    div[data-testid="stFileUploader"] button:hover {
        background-color: #1f2937 !important;
        border-color: rgba(255, 255, 255, 0.5) !important;
        transform: translateY(-1px) !important;
    }
    
    div[data-testid="stFileUploaderDropzone"] button:active,
    div[data-testid="stFileUploader"] button:active {
        transform: translateY(0px) !important;
    }
    
    /* 覆寫 Streamlit 按鈕，將其改造為帶有發光漸層的奢華按鈕 */
    div.stButton > button,
    div[data-testid="stDownloadButton"] button,
    div[data-testid="stDownloadButton"] > button {
        background: linear-gradient(135deg, #8b5cf6, #3b82f6) !important;
        color: #ffffff !important;
        border: none !important;
        padding: 14px 28px !important;
        font-size: 1.05rem !important;
        font-weight: 600 !important;
        border-radius: 16px !important;
        width: 100% !important;
        transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1) !important;
        box-shadow: 0 8px 20px rgba(139, 92, 246, 0.25) !important;
        margin-top: 10px !important;
        text-decoration: none !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
    }
    
    div.stButton > button:hover,
    div[data-testid="stDownloadButton"] button:hover,
    div[data-testid="stDownloadButton"] > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 12px 25px rgba(139, 92, 246, 0.45) !important;
        border: none !important;
        color: #ffffff !important;
        background: linear-gradient(135deg, #9d76fa, #4f94fd) !important;
    }
    
    div.stButton > button:active,
    div[data-testid="stDownloadButton"] button:active,
    div[data-testid="stDownloadButton"] > button:active {
        transform: translateY(1px) !important;
    }

    /* 成果面板客製化樣式 */
    .glass-panel {
        background: rgba(17, 24, 39, 0.65) ;
        border: 1px solid rgba(255, 255, 255, 0.08) ;
        border-radius: 24px ;
        padding: 30px ;
        backdrop-filter: blur(20px) ;
        -webkit-backdrop-filter: blur(20px) ;
        box-shadow: 0 20px 40px rgba(0, 0, 0, 0.3) ;
        margin-top: 20px;
    }
    
    .success-header {
        display: flex;
        align-items: center;
        gap: 20px;
        margin-bottom: 25px;
        border-bottom: 1px solid rgba(255, 255, 255, 0.08);
        padding-bottom: 20px;
    }
    
    .success-icon {
        width: 50px;
        height: 50px;
        background: rgba(16, 185, 129, 0.1);
        border: 1px solid rgba(16, 185, 129, 0.2);
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        color: #10b981;
    }
    
    .success-header h2 {
        font-size: 1.5rem !important;
        font-weight: 600 !important;
        margin: 0 !important;
        color: #f3f4f6;
    }
    
    .success-header p {
        color: #9ca3af;
        font-size: 0.9rem;
        margin: 4px 0 0 0;
    }
    
    /* 數據統計 Grid */
    .stats-grid {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 15px;
        margin-bottom: 25px;
    }
    
    .stat-card {
        background: rgba(0, 0, 0, 0.2);
        border: 1px solid rgba(255, 255, 255, 0.06);
        border-radius: 16px;
        padding: 16px;
        display: flex;
        flex-direction: column;
    }
    
    .stat-label {
        font-size: 0.8rem;
        color: #9ca3af;
        margin-bottom: 4px;
    }
    
    .stat-value {
        font-size: 1.6rem;
        font-weight: 700;
        color: #f3f4f6;
    }
    
    .stat-card.highlight {
        border-color: rgba(16, 185, 129, 0.25);
        background: rgba(16, 185, 129, 0.03);
    }
    
    .stat-card.highlight .stat-value {
        color: #10b981;
    }
    
    .stat-card.warning .stat-value {
        color: #f59e0b;
    }

    /* 警示區 */
    .warning-box {
        background: rgba(245, 158, 11, 0.04);
        border: 1px solid rgba(245, 158, 11, 0.15);
        border-radius: 16px;
        padding: 16px;
        margin-bottom: 25px;
    }
    .warning-title {
        color: #f59e0b;
        font-weight: 600;
        margin-bottom: 8px;
        font-size: 0.95rem;
    }
    .warning-list {
        margin: 0;
        padding-left: 20px;
        color: #9ca3af;
        font-size: 0.85rem;
    }

    /* 對比預覽區 */
    .preview-section {
        margin-top: 30px;
        background: rgba(255, 255, 255, 0.01);
        border: 1px solid rgba(255, 255, 255, 0.03);
        border-radius: 20px;
        padding: 20px;
    }
    .preview-section h3 {
        font-size: 1.1rem !important;
        font-weight: 600 !important;
        margin-bottom: 4px !important;
    }
    .preview-desc {
        color: #9ca3af;
        font-size: 0.85rem;
        margin-bottom: 15px;
    }
    .preview-container {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 20px;
    }
    .preview-box {
        display: flex;
        flex-direction: column;
        gap: 8px;
    }
    .preview-badge {
        align-self: flex-start;
        font-size: 0.75rem;
        font-weight: 600;
        padding: 3px 8px;
        background: rgba(255, 255, 255, 0.04);
        border: 1px solid rgba(255, 255, 255, 0.06);
        border-radius: 6px;
        color: #9ca3af;
    }
    .preview-badge.success {
        background: rgba(16, 185, 129, 0.1);
        border-color: rgba(16, 185, 129, 0.15);
        color: #10b981;
    }
    .preview-img-wrapper {
        background: white;
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 10px 20px rgba(0, 0, 0, 0.3);
        aspect-ratio: 1 / 1.414;
        display: flex;
        align-items: center;
        justify-content: center;
        transition: transform 0.2s;
    }
    .preview-img-wrapper:hover {
        transform: scale(1.02);
    }
    .preview-img-wrapper img {
        width: 100%;
        height: 100%;
        object-fit: contain;
    }

    /* 檔案識別狀態徽章卡片 */
    .file-status-container {
        display: flex;
        gap: 15px;
        margin-top: 15px;
        margin-bottom: 20px;
    }
    .file-status-badge {
        flex: 1;
        background: rgba(17, 24, 39, 0.7);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 14px;
        padding: 14px 18px;
        display: flex;
        align-items: center;
        gap: 14px;
        transition: all 0.2s ease;
    }
    .file-status-badge.ready {
        border-color: rgba(16, 185, 129, 0.35);
        background: rgba(16, 185, 129, 0.06);
    }
    .file-status-badge.waiting {
        border-color: rgba(245, 158, 11, 0.3);
        background: rgba(245, 158, 11, 0.04);
    }
    .file-status-icon {
        font-size: 1.5rem;
    }
    .file-status-info {
        display: flex;
        flex-direction: column;
        overflow: hidden;
    }
    .file-status-title {
        font-weight: 600;
        color: #f3f4f6;
        font-size: 0.95rem;
    }
    .file-status-desc {
        font-size: 0.82rem;
        color: #9ca3af;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        max-width: 320px;
        margin-top: 2px;
    }

    /* 調整 st.tabs 標籤頁樣式 */
    div[data-testid="stTabs"] button {
        color: #9ca3af !important;
        font-size: 1rem !important;
        font-weight: 500 !important;
        padding: 10px 18px !important;
        border-radius: 10px 10px 0 0 !important;
        transition: all 0.2s ease !important;
    }
    div[data-testid="stTabs"] button[aria-selected="true"] {
        color: #ffffff !important;
        border-bottom-color: #8b5cf6 !important;
        font-weight: 700 !important;
        background: rgba(139, 92, 246, 0.08) !important;
    }

    /* 強制 checkbox 標籤文字為高亮淺白，解決深色主題下字體太黑看不清的問題 */
    div[data-testid="stCheckbox"] label p,
    div[data-testid="stCheckbox"] label span,
    div[data-testid="stCheckbox"] p,
    div[data-testid="stCheckbox"] span {
        color: #f3f4f6 !important;
        font-weight: 500 !important;
        font-size: 0.95rem !important;
    }
    </style>
""", unsafe_allow_html=True)

# 3. 渲染奢華的主標題 (HTML)
st.markdown("""
    <div class="custom-header">
        <div class="custom-logo">
            <span class="custom-badge">PRO</span>
            <h1>智慧出貨標籤自動套印系統</h1>
        </div>
    </div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### ⚙️ 系統設定參數")
    st.info("💡 請於主介面上傳對應的 Label PDF 與 Packing List Excel 檔案以執行標記對應。")

# 4. 通用標籤套印與成果渲染流程定義
def execute_label_flow(pdf_input, excel_input, original_pdf_name="labeled_shipping_labels.pdf"):
    """
    通用標籤套印與成果渲染流程。
    可接受 UploadedFile 或本機檔案路徑 (str)。
    """
    with st.spinner("正在進行 PO 號碼匹配與標籤套印中..."):
        tmp_pdf_path = None
        tmp_excel_path = None
        output_pdf_path = None
        clean_tmp_pdf = False
        clean_tmp_excel = False

        try:
            # 處理 PDF 輸入
            if isinstance(pdf_input, str):
                tmp_pdf_path = pdf_input
            else:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                    tmp_pdf.write(pdf_input.read())
                    tmp_pdf_path = tmp_pdf.name
                clean_tmp_pdf = True

            # 處理 Excel 輸入
            if isinstance(excel_input, str):
                tmp_excel_path = excel_input
            else:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as tmp_excel:
                    tmp_excel.write(excel_input.read())
                    tmp_excel_path = tmp_excel.name
                clean_tmp_excel = True

            output_pdf_path = tempfile.mktemp(suffix="_labeled.pdf")

            # 1. 解析 Excel
            excel_data = parse_excel(tmp_excel_path)
            if not excel_data:
                st.error("❌ Excel 檔案中未解析出任何有效的裝箱數據，請確認格式是否正確。")
                return

            # 2. 標示 PDF 檔案
            success, msg, stats = label_pdf(tmp_pdf_path, excel_data, output_pdf_path)
            if not success:
                st.error(f"❌ {msg}")
                return

            # 3. 準備預覽圖 (轉 base64)
            doc = fitz.open(tmp_pdf_path)
            orig_base64 = ""
            if len(doc) > 0:
                pix = doc[0].get_pixmap(dpi=110)
                orig_base64 = f"data:image/png;base64,{base64.b64encode(pix.tobytes('png')).decode('utf-8')}"
            doc.close()

            doc_labeled = fitz.open(output_pdf_path)
            labeled_base64 = ""
            if len(doc_labeled) > 0:
                pix = doc_labeled[0].get_pixmap(dpi=110)
                labeled_base64 = f"data:image/png;base64,{base64.b64encode(pix.tobytes('png')).decode('utf-8')}"
            doc_labeled.close()

            # 4. 準備警示訊息 HTML
            warning_html = ""
            if stats.get("warnings"):
                items = "".join([f"<li>{w}</li>" for w in stats.get("warnings", [])])
                warning_html = f"""
                <div class="warning-box">
                    <div class="warning-title">⚠️ 配對警示訊息 ({len(stats.get('warnings', []))} 條)</div>
                    <ul class="warning-list">{items}</ul>
                </div>
                """

            # 5. 渲染成果卡片
            st.html(f"""
                <div class="glass-panel">
                    <div class="success-header">
                        <div class="success-icon">
                            <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="3" stroke="currentColor" style="width:24px;height:24px;">
                                <path stroke-linecap="round" stroke-linejoin="round" d="M4.5 12.75l6 6 9-13.5" />
                            </svg>
                        </div>
                        <div>
                            <h2>PDF 標籤自動套印成功！</h2>
                            <p>所有裝箱明細與 PDF 條碼 PO 號碼皆已自動配對標記完成。</p>
                        </div>
                    </div>
                    
                    <div class="stats-grid">
                        <div class="stat-card">
                            <span class="stat-label">PDF 總頁數</span>
                            <span class="stat-value">{stats.get('total_pages', 0)} 頁</span>
                        </div>
                        <div class="stat-card highlight">
                            <span class="stat-label">成功匹配標籤</span>
                            <span class="stat-value">{stats.get('matched_count', 0)} 個</span>
                        </div>
                        <div class="stat-card {'warning' if stats.get('missing_labels', 0) > 0 else ''}">
                            <span class="stat-label">缺少標籤 (Excel剩餘)</span>
                            <span class="stat-value">{stats.get('missing_labels', 0)} 個</span>
                        </div>
                        <div class="stat-card {'warning' if stats.get('unmatched_labels', 0) > 0 else ''}">
                            <span class="stat-label">多餘標籤 (PDF多出)</span>
                            <span class="stat-value">{stats.get('unmatched_labels', 0)} 個</span>
                        </div>
                    </div>
                    
                    {warning_html}
                    
                    <div class="preview-section">
                        <h3>效果即時預覽 (第一頁)</h3>
                        <p class="preview-desc">您可以將滑鼠移至圖片上以放大查看套印細節</p>
                        <div class="preview-container">
                            <div class="preview-box">
                                <span class="preview-badge">原始 PDF 標籤</span>
                                <div class="preview-img-wrapper">
                                    <img src="{orig_base64}" alt="原始 PDF 預覽">
                                </div>
                            </div>
                            <div class="preview-box">
                                <span class="preview-badge success">自動標示後 PDF</span>
                                <div class="preview-img-wrapper">
                                    <img src="{labeled_base64}" alt="標示後 PDF 預覽">
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            """)

            st.write("")

            # 6. 提供下載按鈕
            with open(output_pdf_path, "rb") as f:
                pdf_bytes = f.read()

            out_filename = f"labeled_{os.path.basename(original_pdf_name)}" if not original_pdf_name.startswith("labeled_") else original_pdf_name
            st.download_button(
                label="📥 下載標示後的 PDF 檔案",
                data=pdf_bytes,
                file_name=out_filename,
                mime="application/pdf",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"❌ 處理檔案時發生未知錯誤: {e}")
        finally:
            if clean_tmp_pdf and tmp_pdf_path and os.path.exists(tmp_pdf_path):
                os.unlink(tmp_pdf_path)
            if clean_tmp_excel and tmp_excel_path and os.path.exists(tmp_excel_path):
                os.unlink(tmp_excel_path)
            if output_pdf_path and os.path.exists(output_pdf_path):
                os.unlink(output_pdf_path)


# 5. 左右對稱雙欄佈局 (左：步驟 0 預覽箱數 / 右：單一入口智慧套印)
col_left, col_right = st.columns([1, 1], gap="large")

df_preview = None
parsed_excel_data = None
detected_delivery_date = None

with col_left:
    st.markdown("#### 📋 預覽標籤箱數 (選填)")
    st.caption("前往下載標籤前，先丟 Packing List 計算各 PO 需求箱數。")
    preview_excel_file = st.file_uploader(
        "請上傳 Packing List (Excel)",
        type=["xls", "xlsx"],
        key="preview_uploader"
    )

    if preview_excel_file:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as tmp_preview_excel:
            tmp_preview_excel.write(preview_excel_file.read())
            tmp_preview_excel_path = tmp_preview_excel.name
            
        try:
            with st.spinner("計算中..."):
                excel_data = parse_excel(tmp_preview_excel_path)
                if not excel_data:
                    st.error("❌ 未解析出有效裝箱數據。")
                else:
                    parsed_excel_data = excel_data
                    detected_delivery_date = find_delivery_date(tmp_preview_excel_path)
                    po_rows = []
                    for po, boxes in excel_data.items():
                        formatted_po = po
                        if len(po) == 14 and po.startswith("53"):
                            formatted_po = f"{po[0:2]}-{po[2:8]}-{po[8:14]}"
                        
                        product_nos = sorted(list(set(box.get("product_no", "") for box in boxes if box.get("product_no"))))
                        descriptions = sorted(list(set(box.get("description", "") for box in boxes if box.get("description"))))
                        
                        po_rows.append({
                            "PO 號碼": formatted_po,
                            "Product NO": ", ".join(product_nos),
                            "Description": ", ".join(descriptions),
                            "需求箱數": len(boxes),
                            "總件數": sum(box["qty"] for box in boxes)
                        })
                    
                    df_preview = pd.DataFrame(po_rows).sort_values(by="PO 號碼")
                    st.success(f"✅ 已解析 {len(df_preview)} 組 PO，填表小工具與箱數統計已展開於下方 ⬇️")
        except Exception as e:
            st.error(f"❌ 解析失敗: {e}")
        finally:
            if os.path.exists(tmp_preview_excel_path):
                os.unlink(tmp_preview_excel_path)

with col_right:
    st.markdown("#### ⚡ 智慧標籤自動套印 (核心)")
    st.caption("同時拖曳 Packing List 與客戶 Label PDF，自動分類套印。")
    uploaded_files = st.file_uploader(
        "請將 Packing List 與 Label PDF 同時拖曳至此處",
        type=["pdf", "xls", "xlsx"],
        accept_multiple_files=True,
        key="unified_uploader"
    )

    detected_pdf = None
    detected_excel = None

    if uploaded_files:
        for f in uploaded_files:
            fn = f.name.lower()
            if fn.endswith(".pdf") and not detected_pdf:
                detected_pdf = f
            elif (fn.endswith(".xls") or fn.endswith(".xlsx")) and not detected_excel:
                detected_excel = f

    pdf_ready = detected_pdf is not None
    excel_ready = detected_excel is not None

    pdf_badge_class = "ready" if pdf_ready else "waiting"
    pdf_badge_icon = "📄" if pdf_ready else "⏳"
    pdf_badge_title = f"Label PDF：{detected_pdf.name}" if pdf_ready else "客戶 Label PDF"
    pdf_badge_desc = f"{detected_pdf.size / 1024:.1f} KB" if pdf_ready else "等待上傳中 (.pdf)"

    excel_badge_class = "ready" if excel_ready else "waiting"
    excel_badge_icon = "📊" if excel_ready else "⏳"
    excel_badge_title = f"Packing List：{detected_excel.name}" if excel_ready else "倉庫 Packing List"
    excel_badge_desc = f"{detected_excel.size / 1024:.1f} KB" if excel_ready else "等待上傳中 (.xls, .xlsx)"

    # 若左欄未上傳，但右欄有偵測到 Packing List，自動解析以支援方案 A 填表小幫手
    if parsed_excel_data is None and detected_excel is not None:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as tmp_sync_excel:
            tmp_sync_excel.write(detected_excel.getvalue())
            tmp_sync_excel_path = tmp_sync_excel.name
        try:
            excel_data = parse_excel(tmp_sync_excel_path)
            if excel_data:
                parsed_excel_data = excel_data
                detected_delivery_date = find_delivery_date(tmp_sync_excel_path)
                po_rows = []
                for po, boxes in excel_data.items():
                    formatted_po = po
                    if len(po) == 14 and po.startswith("53"):
                        formatted_po = f"{po[0:2]}-{po[2:8]}-{po[8:14]}"
                    product_nos = sorted(list(set(box.get("product_no", "") for box in boxes if box.get("product_no"))))
                    descriptions = sorted(list(set(box.get("description", "") for box in boxes if box.get("description"))))
                    po_rows.append({
                        "PO 號碼": formatted_po,
                        "Product NO": ", ".join(product_nos),
                        "Description": ", ".join(descriptions),
                        "需求箱數": len(boxes),
                        "總件數": sum(box["qty"] for box in boxes)
                    })
                df_preview = pd.DataFrame(po_rows).sort_values(by="PO 號碼")
        except Exception:
            pass
        finally:
            if os.path.exists(tmp_sync_excel_path):
                os.unlink(tmp_sync_excel_path)

    st.markdown(f"""
        <div class="file-status-container" style="margin-top: 8px; margin-bottom: 12px;">
            <div class="file-status-badge {pdf_badge_class}" style="padding: 8px 12px;">
                <div class="file-status-icon">{pdf_badge_icon}</div>
                <div class="file-status-info">
                    <div class="file-status-title" style="font-size: 0.86rem;">{pdf_badge_title}</div>
                    <div class="file-status-desc">{pdf_badge_desc}</div>
                </div>
            </div>
            <div class="file-status-badge {excel_badge_class}" style="padding: 8px 12px;">
                <div class="file-status-icon">{excel_badge_icon}</div>
                <div class="file-status-info">
                    <div class="file-status-title" style="font-size: 0.86rem;">{excel_badge_title}</div>
                    <div class="file-status-desc">{excel_badge_desc}</div>
                </div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    auto_run = st.checkbox("⚡ 檔案辨識齊全後自動觸發標示", value=True, key="auto_run_unified")

    if pdf_ready and excel_ready:
        if not auto_run:
            if st.button("🚀 開始自動對應並標示", key="btn_run_unified", use_container_width=True):
                execute_label_flow(detected_pdf, detected_excel, detected_pdf.name)
    elif uploaded_files:
        missing = []
        if not pdf_ready:
            missing.append("客戶 Label PDF")
        if not excel_ready:
            missing.append("倉庫 Packing List")
        st.warning(f"⚠️ 尚缺少：{' 與 '.join(missing)}。")

# 6. 若有解析出 Packing List，渲染方案 A：一鍵填表小幫手與統計表
if parsed_excel_data is not None and df_preview is not None:
    st.write("")
    
    # 決定基準出貨日
    from datetime import datetime, timedelta
    today = datetime.now()
    if detected_delivery_date is None or pd.isna(detected_delivery_date):
        # 預設為當週或下一個週五
        days_ahead = (4 - today.weekday()) % 7
        base_ship_date = today + timedelta(days=days_ahead)
    else:
        base_ship_date = detected_delivery_date

    rec_start, rec_end = calculate_date_range(base_ship_date)
    rec_start_val = rec_start.date() if hasattr(rec_start, "date") else rec_start
    rec_end_val = rec_end.date() if hasattr(rec_end, "date") else rec_end

    # 方案 A 容器卡片
    st.markdown("""
        <div style="background: rgba(17, 24, 39, 0.65); border: 1px solid rgba(139, 92, 246, 0.3); border-radius: 20px; padding: 22px 24px; margin-top: 10px; margin-bottom: 25px; box-shadow: 0 10px 25px rgba(0, 0, 0, 0.35);">
            <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px; border-bottom: 1px solid rgba(255, 255, 255, 0.08); padding-bottom: 12px;">
                <div style="display: flex; align-items: center; gap: 10px;">
                    <span style="font-size: 1.5rem;">⚡</span>
                    <div>
                        <h3 style="margin: 0; font-size: 1.25rem; font-weight: 700; color: #f3f4f6;">方案 A：客戶出貨平台一鍵自動填表小幫手</h3>
                        <p style="margin: 2px 0 0 0; font-size: 0.85rem; color: #9ca3af;">免手動切換編輯、免逐筆填寫！一鍵自動完成日期篩選、填寫各 PO 箱數與全選。</p>
                    </div>
                </div>
                <span style="background: rgba(16, 185, 129, 0.15); color: #34d399; font-size: 0.78rem; font-weight: 600; padding: 4px 10px; border-radius: 20px; border: 1px solid rgba(16, 185, 129, 0.25);">
                    免安裝任何軟體
                </span>
            </div>
            <div style="font-size: 0.88rem; color: #d1d5db; line-height: 1.6; margin-bottom: 16px;">
                📌 <b>出貨日計算規則</b>：依公司固定<b>週五出貨</b>（遇假日提前），本次篩選截止日自動計算至<b>下次出貨日前一天</b>。
            </div>
        </div>
    """, unsafe_allow_html=True)

    # 日期微調列與書籤生成
    date_col1, date_col2 = st.columns([1, 1])
    with date_col1:
        sel_start = st.date_input("📅 供應商出貨日 (起始日)", value=rec_start_val, key="custom_start_date")
    with date_col2:
        sel_end = st.date_input("📅 篩選截止日 (下次出貨日前一天)", value=rec_end_val, key="custom_end_date")

    # 動態產生對應選定日期的 Bookmarklet 與腳本
    raw_js, bookmarklet_url, summary = generate_portal_bookmarklet(sel_start, sel_end, parsed_excel_data)

    st.markdown(f"""
        <div style="background: rgba(0, 0, 0, 0.25); border: 1px solid rgba(255, 255, 255, 0.06); border-radius: 14px; padding: 16px 20px; margin-top: 10px; margin-bottom: 20px;">
            <div style="font-weight: 600; color: #f3f4f6; margin-bottom: 10px; font-size: 0.95rem;">
                🚀 選擇執行方式 (兩種管道皆可)：
            </div>
            <div style="display: flex; flex-wrap: wrap; align-items: center; gap: 16px;">
                <div>
                    <a href="{bookmarklet_url}" class="bookmarklet-btn" title="拖曳至瀏覽器上方書籤列">
                        ⭐ 拖曳此按鈕至書籤列 (客戶平台自動填表)
                    </a>
                </div>
                <div style="font-size: 0.85rem; color: #9ca3af;">
                    👈 <b>滑鼠拖曳上方按鈕</b> 到 Chrome/Edge 書籤列，登入客戶平台後點一下即刻全自動填寫！
                </div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    with st.expander("💻 或直接在瀏覽器 F12 主控台 (Console) 一鍵貼上執行 (點擊展開代碼)"):
        st.caption("在客戶出貨管理頁面按 F12 鍵 ➡️ 點擊 Console (主控台) ➡️ 貼上下方代碼並按 Enter 即可：")
        st.code(raw_js, language="javascript")

    st.write("")
    st.markdown("### 📊 各 PO 標籤需求箱數統計表")
    st.dataframe(
        df_preview,
        use_container_width=True,
        hide_index=True,
        column_config={
            "PO 號碼": st.column_config.TextColumn("PO 號碼", width=150),
            "Product NO": st.column_config.TextColumn("Product NO", width=130),
            "Description": st.column_config.TextColumn("Description", width=220),
            "需求箱數": st.column_config.NumberColumn("需求箱數", format="%d 箱", width=100),
            "總件數": st.column_config.NumberColumn("總出貨件數", format="%d pcs", width=110),
        }
    )

# 7. 若滿足自動執行條件，在全寬下方執行並渲染成果卡片
if pdf_ready and excel_ready and auto_run:
    st.write("")
    execute_label_flow(detected_pdf, detected_excel, detected_pdf.name)


