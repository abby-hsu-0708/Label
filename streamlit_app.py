import streamlit as st
import os
import fitz  # PyMuPDF
import tempfile
import base64
import re
from core_logic import parse_excel, label_pdf

# 1. 設置頁面配置
st.set_page_config(page_title="MISUMI 標籤自動標示工具", layout="centered", page_icon="🏷️")

# 2. 注入奢華的極光與玻璃擬物化 (Glassmorphism) 全域 CSS
# 這會徹底覆寫 Streamlit 的原生樣式，將其改裝為跟本地版一模一樣的精美介面
st.markdown("""
    <style>
    /* 全域字型與背景 */
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;700&family=Noto+Sans+TC:wght@300;400;500;700&display=swap');
    
    .stApp {
        background-color: #080b11 !important;
        color: #f3f4f6 !important;
        font-family: 'Outfit', 'Noto Sans TC', sans-serif !important;
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
    div[data-testid="stFileUploader"] {
        background: #090d16 !important; /* 更深邃的背景色 */
        border: 1px solid rgba(255, 255, 255, 0.08) !important;
        border-radius: 24px !important;
        padding: 30px 25px !important; /* 調整 padding 使上下高度更合適 */
        backdrop-filter: blur(20px) !important;
        -webkit-backdrop-filter: blur(20px) !important;
        box-shadow: 0 15px 30px rgba(0, 0, 0, 0.4) !important;
        transition: all 0.3s ease !important;
        margin-bottom: 25px !important;
        margin-left: 10px !important; /* 增加左右邊距，讓兩個方塊更隔開 */
        margin-right: 10px !important;
        min-height: 290px !important; /* 強制設定卡片最低高度，確保左右大小完全一致 */
        display: flex !important;
        flex-direction: column !important;
        justify-content: space-between !important;
    }
    
    /* 強制將上傳框的 Label 容器設為固定高度，避免折行影響整體卡片高度 */
    div[data-testid="stFileUploader"] label {
        min-height: 55px !important;
        display: inline-flex !important;
        align-items: center !important;
    }
    
    /* 將上傳框標題字體顏色改為明亮淺灰色 */
    div[data-testid="stFileUploader"] label p {
        color: #f3f4f6 !important;
        font-weight: 600 !important;
        font-size: 1.05rem !important;
        margin: 0 !important;
    }
    
    div[data-testid="stFileUploader"]:hover {
        border-color: rgba(255, 255, 255, 0.2) !important;
        box-shadow: 0 20px 40px rgba(0, 0, 0, 0.5) !important;
    }

    /* 上傳框內部字體顏色與圖示微調 */
    div[data-testid="stFileUploaderDropzone"] {
        background-color: #030712 !important; /* 更深更黑的 Dropzone 背景 */
        border: 1px dashed rgba(255, 255, 255, 0.2) !important;
        border-radius: 16px !important;
        height: 140px !important; /* 固定高度使其呈現標準長方形，解決肥矮胖正方形問題 */
        display: flex !important;
        flex-direction: column !important;
        justify-content: center !important;
        align-items: center !important;
        padding: 15px !important;
    }
    
    /* 確保 dropzone 內部的子 div 不會有預設的亮色背景，並完美居中 */
    div[data-testid="stFileUploaderDropzone"] > div {
        background-color: transparent !important;
        display: flex !important;
        flex-direction: column !important;
        align-items: center !important;
        justify-content: center !important;
        gap: 8px !important;
        width: 100% !important;
    }
    
    div[data-testid="stFileUploaderDropzone"] h3, 
    div[data-testid="stFileUploaderDropzone"] span, 
    div[data-testid="stFileUploaderDropzone"] small,
    div[data-testid="stFileUploaderDropzone"] p {
        color: #e2e8f0 !important; /* 明亮的白灰色 */
        font-weight: 500 !important;
    }

    /* 針對 Streamlit 內置的 "Upload" / "Browse files" 按鈕 */
    div[data-testid="stFileUploaderDropzone"] button,
    div[data-testid="stFileUploader"] button,
    .stFileUploader button {
        background-color: #000000 !important; /* 直接改為純黑色背景 */
        color: #ffffff !important; /* 白色字體 */
        border: 1px solid rgba(255, 255, 255, 0.3) !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        padding: 10px 20px !important;
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
        background-color: #1f2937 !important; /* hover 時稍微變淺一點的黑灰色 */
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
        grid-template-columns: repeat(3, 1fr);
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
    </style>
""", unsafe_allow_html=True)

# 3. 渲染奢華的主標題 (HTML)
st.markdown("""
    <div class="custom-header">
        <div class="custom-logo">
            <span class="custom-badge">PRO</span>
            <h1>MISUMI 標籤自動標示系統</h1>
        </div>
        <p class="custom-subtitle">只需上傳原始標籤 PDF 與倉庫 Packing List，即可在幾秒鐘內精確套印箱號與數量。</p>
    </div>
""", unsafe_allow_html=True)

# 4. 檔案上傳器 (左右排列分布)
col1, col2 = st.columns(2, gap="large")
with col1:
    pdf_file = st.file_uploader("1. 請上傳客戶 Label PDF 檔案", type=["pdf"])
with col2:
    excel_file = st.file_uploader("2. 請上傳倉庫 Packing List Excel 檔案", type=["xls", "xlsx"])

# 5. 開始標示按鈕
if pdf_file and excel_file:
    st.write("")
    if st.button("🚀 開始自動對應並標示", use_container_width=True):
        with st.spinner("正在進行 PO 號碼匹配與標籤套印中..."):
            try:
                # 建立臨時檔案儲存上傳的文件
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                    tmp_pdf.write(pdf_file.read())
                    tmp_pdf_path = tmp_pdf.name
                    
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as tmp_excel:
                    tmp_excel.write(excel_file.read())
                    tmp_excel_path = tmp_excel.name
                    
                output_pdf_path = tmp_pdf_path + "_labeled.pdf"
                
                # 1. 解析 Excel
                excel_data = parse_excel(tmp_excel_path)
                
                if not excel_data:
                    st.error("❌ Excel 檔案中未解析出任何有效的裝箱數據，請確認格式是否正確。")
                else:
                    # 2. 標示 PDF 檔案
                    success, msg, stats = label_pdf(tmp_pdf_path, excel_data, output_pdf_path)
                    
                    if success:
                        # 準備預覽圖 (轉 base64)
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
                        
                        # 準備警示訊息 HTML
                        warning_html = ""
                        if stats["warnings"]:
                            items = "".join([f"<li>{w}</li>" for w in stats["warnings"]])
                            warning_html = f"""
                            <div class="warning-box">
                                <div class="warning-title">⚠️ 配對警示訊息 ({len(stats['warnings'])} 條)</div>
                                <ul class="warning-list">{items}</ul>
                            </div>
                            """
                        
                        # 渲染和本地版一模一樣的 Glassmorphism 成果卡片
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
                                        <span class="stat-value">{stats['total_pages']} 頁</span>
                                    </div>
                                    <div class="stat-card highlight">
                                        <span class="stat-label">成功匹配標籤</span>
                                        <span class="stat-value">{stats['matched_count']} 個</span>
                                    </div>
                                    <div class="stat-card {'warning' if stats['unmatched_labels'] > 0 else ''}">
                                        <span class="stat-label">未匹配標籤</span>
                                        <span class="stat-value">{stats['unmatched_labels']} 個</span>
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
                        
                        # 提供下載按鈕
                        with open(output_pdf_path, "rb") as f:
                            pdf_bytes = f.read()
                        
                        st.download_button(
                            label="📥 下載標示後的 PDF 檔案",
                            data=pdf_bytes,
                            file_name="MISUMI_labeled_labels.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                    else:
                        st.error(f"❌ {msg}")
                        
                # 刪除臨時檔案
                os.unlink(tmp_pdf_path)
                os.unlink(tmp_excel_path)
                if os.path.exists(output_pdf_path):
                    os.unlink(output_pdf_path)
                    
            except Exception as e:
                st.error(f"❌ 處理檔案時發生未知錯誤: {e}")
