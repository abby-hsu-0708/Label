import streamlit as st
import os
import fitz  # PyMuPDF
import tempfile
import re
from core_logic import parse_excel, label_pdf

st.set_page_config(page_title="MISUMI 標籤自動標示工具", layout="centered", page_icon="🏷️")

# 自訂 Streamlit 的 CSS 樣式使其更有現代感
st.markdown("""
    <style>
    .main {
        background-color: #0b0f19;
        color: #f3f4f6;
    }
    h1 {
        background: linear-gradient(135deg, #fff 40%, #9ca3af);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    div[data-testid="stMetricValue"] {
        color: #10b981;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🏷️ MISUMI 標籤自動標示系統")
st.write("上傳原始標籤 PDF 與倉庫 Packing List，系統會自動在指定座標套印箱號與數量。")
st.write("---")

# 建立檔案上傳器
pdf_file = st.file_uploader("1. 請上傳客戶 Label PDF 檔案", type=["pdf"])
excel_file = st.file_uploader("2. 請上傳倉庫 Packing List Excel 檔案", type=["xls", "xlsx"])

if pdf_file and excel_file:
    st.write("")
    if st.button("🚀 開始自動對應並標示", type="primary", use_container_width=True):
        with st.spinner("正在進行 PO 號碼匹配與標籤套印中..."):
            try:
                # 建立臨時檔案來儲存上傳的文件
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                    tmp_pdf.write(pdf_file.read())
                    tmp_pdf_path = tmp_pdf.name
                    
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as tmp_excel:
                    tmp_excel.write(excel_file.read())
                    tmp_excel_path = tmp_excel.name
                    
                output_pdf_path = tmp_pdf_path + "_labeled.pdf"
                
                # 1. 解析 Excel 裝箱資料
                excel_data = parse_excel(tmp_excel_path)
                
                if not excel_data:
                    st.error("❌ Excel 檔案中未解析出任何有效的裝箱數據，請確認格式是否正確。")
                else:
                    # 2. 標示 PDF 檔案
                    success, msg, stats = label_pdf(tmp_pdf_path, excel_data, output_pdf_path)
                    
                    if success:
                        st.success("🎉 標籤自動套印成功！")
                        
                        # 顯示統計數據面版
                        col1, col2, col3 = st.columns(3)
                        col1.metric("PDF 總頁數", f"{stats['total_pages']} 頁")
                        col2.metric("成功匹配標籤數", f"{stats['matched_count']} 個")
                        col3.metric("未匹配標籤數", f"{stats['unmatched_labels']} 個")
                        
                        # 顯示警示訊息 (如果有)
                        if stats["warnings"]:
                            with st.expander("⚠️ 查看配對警示訊息"):
                                for w in stats["warnings"]:
                                    st.warning(w)
                                    
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
                        
                        # 提供第一頁套印預覽
                        doc = fitz.open(output_pdf_path)
                        if len(doc) > 0:
                            pix = doc[0].get_pixmap(dpi=130)
                            img_bytes = pix.tobytes("png")
                            st.write("---")
                            st.subheader("📸 效果即時預覽 (第一頁)")
                            st.image(img_bytes, caption="套印後的 PDF 第一頁預覽", use_column_width=True)
                            doc.close()
                    else:
                        st.error(f"❌ {msg}")
                        
                # 刪除臨時檔案
                os.unlink(tmp_pdf_path)
                os.unlink(tmp_excel_path)
                if os.path.exists(output_pdf_path):
                    os.unlink(output_pdf_path)
                    
            except Exception as e:
                st.error(f"❌ 處理檔案時發生未知錯誤: {e}")
