import os
import uuid
import shutil
import webbrowser
import uvicorn
import base64
import fitz
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from core_logic import parse_excel, label_pdf

app = FastAPI(title="MISUMI 標籤自動標示工具", description="上傳 Label PDF 與 Packing List Excel，自動標示箱號、數量與序號。")

# 建立暫存與輸出目錄
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMP_DIR = os.path.join(BASE_DIR, "temp_data")
STATIC_DIR = os.path.join(BASE_DIR, "static")

os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(STATIC_DIR, exist_ok=True)

# 掛載靜態檔案目錄，前端網頁將放在這裡
if os.path.exists(STATIC_DIR):
    app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

@app.get("/")
async def read_index():
    """首頁路由，直接返回前端 HTML 頁面"""
    index_path = os.path.join(STATIC_DIR, "index.html")
    if os.path.exists(index_path):
        return FileResponse(index_path)
    return JSONResponse(status_code=404, content={"message": "前端 index.html 尚未建立"})

def pdf_page_to_base64(pdf_path, page_idx=0, dpi=150):
    """將 PDF 指定頁面渲染並轉換成 Base64 PNG 圖片，用於前端即時預覽"""
    try:
        doc = fitz.open(pdf_path)
        if len(doc) <= page_idx:
            return ""
        page = doc[page_idx]
        pix = page.get_pixmap(dpi=dpi)
        img_data = pix.tobytes("png")
        doc.close()
        return f"data:image/png;base64,{base64.b64encode(img_data).decode('utf-8')}"
    except Exception as e:
        print(f"渲染 PDF 預覽圖出錯: {e}")
        return ""

@app.post("/api/process")
async def process_files(
    pdf_file: UploadFile = File(...),
    excel_file: UploadFile = File(...)
):
    """
    處理上傳的 PDF 與 Excel 檔案。
    """
    # 產生唯一的隨機 ID 作為此次處理任務的資料夾
    task_id = str(uuid.uuid4())
    task_dir = os.path.join(TEMP_DIR, task_id)
    os.makedirs(task_dir, exist_ok=True)
    
    # 儲存上傳的檔案
    input_pdf_path = os.path.join(task_dir, f"input_{pdf_file.filename}")
    input_excel_path = os.path.join(task_dir, f"input_{excel_file.filename}")
    output_pdf_path = os.path.join(task_dir, "labeled_output.pdf")
    
    try:
        with open(input_pdf_path, "wb") as buffer:
            shutil.copyfileobj(pdf_file.file, buffer)
        with open(input_excel_path, "wb") as buffer:
            shutil.copyfileobj(excel_file.file, buffer)
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"success": False, "message": f"儲存上傳檔案失敗: {str(e)}"}
        )
        
    # 1. 產生原始 PDF 第一頁的預覽圖
    original_preview = pdf_page_to_base64(input_pdf_path, 0, dpi=120)
    
    # 2. 解析 Excel
    try:
        excel_data = parse_excel(input_excel_path)
    except Exception as e:
        return JSONResponse(
            status_code=400,
            content={"success": False, "message": f"解析 Excel Packing List 失敗: {str(e)}"}
        )
        
    if not excel_data:
        return JSONResponse(
            status_code=400,
            content={"success": False, "message": "Excel 檔案中未解析出任何有效的裝箱數據，請確認格式是否正確。"}
        )
        
    # 3. 標示 PDF
    success, msg, stats = label_pdf(input_pdf_path, excel_data, output_pdf_path)
    
    if not success:
        return JSONResponse(
            status_code=500,
            content={"success": False, "message": msg}
        )
        
    # 4. 產生標示後 PDF 第一頁的預覽圖
    labeled_preview = pdf_page_to_base64(output_pdf_path, 0, dpi=120)
    
    # 計算 PO 解析摘要以供前端顯示
    po_summary = []
    for po, boxes in excel_data.items():
        # 取得這一組的 prefix
        po_summary.append({
            "po": po,
            "box_count": len(boxes),
            "prefix": boxes[0]["display_box"] if len(boxes) == 1 else boxes[0]["box_no"][:2]
        })
    
    return {
        "success": True,
        "message": msg,
        "task_id": task_id,
        "stats": stats,
        "po_summary": po_summary,
        "original_preview": original_preview,
        "labeled_preview": labeled_preview
    }

@app.get("/api/download/{task_id}")
async def download_file(task_id: str):
    """
    下載指定任務處理後的 PDF 檔案。
    """
    output_pdf_path = os.path.join(TEMP_DIR, task_id, "labeled_output.pdf")
    if not os.path.exists(output_pdf_path):
        raise HTTPException(status_code=404, detail="找不到處理後的 PDF 檔案，可能已過期，請重新上傳處理。")
        
    return FileResponse(
        output_pdf_path,
        media_type="application/pdf",
        filename="MISUMI_labeled_labels.pdf"
    )

if __name__ == "__main__":
    # 自動在預設瀏覽器中開啟首頁
    url = "http://127.0.0.1:18000"
    print(f"正在啟動 MISUMI 標籤工具服務，將自動開啟網頁：{url}")
    webbrowser.open(url)
    
    # 啟動 uvicorn 服務
    uvicorn.run("main:app", host="127.0.0.1", port=18000, reload=False)
