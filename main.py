import os
import sys
import uuid
import shutil
import webbrowser
import threading
import socket
import time
import base64
import fitz
import pandas as pd
from datetime import datetime, timedelta
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
import uvicorn

from core_logic import (
    parse_excel,
    label_pdf,
    find_delivery_date,
    calculate_date_range,
    generate_portal_bookmarklet
)

app = FastAPI(title="智慧出貨標籤自動套印系統", description="出貨裝箱標籤自動對應標示與平台輔助工具")

# 跨域支援
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 判斷是否為 PyInstaller 打包環境
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
    DATA_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    DATA_DIR = BASE_DIR

TEMP_DIR = os.path.join(DATA_DIR, "temp_data")
STATIC_DIR = os.path.join(BASE_DIR, "static")

os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(STATIC_DIR, exist_ok=True)

if os.path.exists(STATIC_DIR):
    app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

@app.get("/")
async def read_index():
    """首頁路由，返回前端 HTML 頁面"""
    index_path = os.path.join(STATIC_DIR, "index.html")
    if os.path.exists(index_path):
        return FileResponse(index_path)
    return JSONResponse(status_code=404, content={"message": "前端介面正在載入中..."})

def pdf_page_to_base64(pdf_path, page_idx=0, dpi=120):
    """將 PDF 指定頁面轉換為 Base64 預覽圖"""
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
        print(f"渲染預覽圖失敗: {e}")
        return ""

@app.post("/api/preview_excel")
async def preview_excel(excel_file: UploadFile = File(...)):
    """解析 Packing List Excel 並回傳出貨日、篩選區間、各 PO 箱數與書籤代碼"""
    task_id = str(uuid.uuid4())
    task_dir = os.path.join(TEMP_DIR, task_id)
    os.makedirs(task_dir, exist_ok=True)
    temp_excel = os.path.join(task_dir, f"preview_{excel_file.filename}")

    try:
        with open(temp_excel, "wb") as buffer:
            shutil.copyfileobj(excel_file.file, buffer)

        excel_data = parse_excel(temp_excel)
        if not excel_data:
            return JSONResponse(status_code=400, content={"success": False, "message": "未從 Excel 中解析出有效裝箱明細。"})

        deliv_date = find_delivery_date(temp_excel)
        today = datetime.now()
        if deliv_date is None or pd.isna(deliv_date):
            days_ahead = (4 - today.weekday()) % 7
            base_ship_date = today + timedelta(days=days_ahead)
        else:
            base_ship_date = deliv_date

        start_dt, end_dt = calculate_date_range(base_ship_date)
        raw_js, bookmarklet_url, summary = generate_portal_bookmarklet(start_dt, end_dt, excel_data)

        # 整理表格清單
        po_rows = []
        for po, boxes in excel_data.items():
            formatted_po = po
            if len(po) == 14 and po.startswith("53"):
                formatted_po = f"{po[0:2]}-{po[2:8]}-{po[8:14]}"
            product_nos = sorted(list(set(b.get("product_no", "") for b in boxes if b.get("product_no"))))
            descriptions = sorted(list(set(b.get("description", "") for b in boxes if b.get("description"))))
            po_rows.append({
                "po": formatted_po,
                "product_no": ", ".join(product_nos),
                "description": ", ".join(descriptions),
                "box_count": len(boxes),
                "total_qty": sum(b["qty"] for b in boxes)
            })

        po_rows.sort(key=lambda x: x["po"])

        return {
            "success": True,
            "delivery_date": str(deliv_date.date()) if deliv_date and not pd.isna(deliv_date) else None,
            "start_date": summary["start_date"],
            "end_date": summary["end_date"],
            "raw_js": raw_js,
            "bookmarklet_url": bookmarklet_url,
            "po_rows": po_rows,
            "total_boxes": sum(r["box_count"] for r in po_rows)
        }
    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "message": f"解析失敗: {str(e)}"})
    finally:
        shutil.rmtree(task_dir, ignore_errors=True)

@app.post("/api/generate_script")
async def generate_script(
    start_date: str = Form(...),
    end_date: str = Form(...),
    excel_file: UploadFile = File(...)
):
    """手動調整日期後重新生成書籤腳本"""
    task_id = str(uuid.uuid4())
    task_dir = os.path.join(TEMP_DIR, task_id)
    os.makedirs(task_dir, exist_ok=True)
    temp_excel = os.path.join(task_dir, f"script_{excel_file.filename}")

    try:
        with open(temp_excel, "wb") as buffer:
            shutil.copyfileobj(excel_file.file, buffer)
        excel_data = parse_excel(temp_excel)
        raw_js, bookmarklet_url, summary = generate_portal_bookmarklet(start_date, end_date, excel_data)
        return {
            "success": True,
            "raw_js": raw_js,
            "bookmarklet_url": bookmarklet_url
        }
    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "message": str(e)})
    finally:
        shutil.rmtree(task_dir, ignore_errors=True)

@app.post("/api/process")
async def process_files(
    pdf_file: UploadFile = File(...),
    excel_file: UploadFile = File(...)
):
    """標籤對應與套印"""
    task_id = str(uuid.uuid4())
    task_dir = os.path.join(TEMP_DIR, task_id)
    os.makedirs(task_dir, exist_ok=True)
    
    input_pdf_path = os.path.join(task_dir, f"input_{pdf_file.filename}")
    input_excel_path = os.path.join(task_dir, f"input_{excel_file.filename}")
    output_pdf_path = os.path.join(task_dir, "labeled_output.pdf")
    
    try:
        with open(input_pdf_path, "wb") as buffer:
            shutil.copyfileobj(pdf_file.file, buffer)
        with open(input_excel_path, "wb") as buffer:
            shutil.copyfileobj(excel_file.file, buffer)
            
        original_preview = pdf_page_to_base64(input_pdf_path, 0, dpi=120)
        excel_data = parse_excel(input_excel_path)
        if not excel_data:
            return JSONResponse(status_code=400, content={"success": False, "message": "Excel 檔案未解析出有效裝箱數據。"})
            
        success, msg, stats = label_pdf(input_pdf_path, excel_data, output_pdf_path)
        if not success:
            return JSONResponse(status_code=500, content={"success": False, "message": msg})
            
        labeled_preview = pdf_page_to_base64(output_pdf_path, 0, dpi=120)
        
        return {
            "success": True,
            "message": msg,
            "task_id": task_id,
            "stats": stats,
            "original_preview": original_preview,
            "labeled_preview": labeled_preview
        }
    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "message": f"套印出錯: {str(e)}"})

@app.get("/api/download/{task_id}")
async def download_file(task_id: str):
    """下載標示後的 PDF 檔案"""
    output_pdf_path = os.path.join(TEMP_DIR, task_id, "labeled_output.pdf")
    if not os.path.exists(output_pdf_path):
        raise HTTPException(status_code=404, detail="檔案已過期或不存在，請重新上傳處理。")
    return FileResponse(
        output_pdf_path,
        media_type="application/pdf",
        filename="labeled_shipping_labels.pdf"
    )

def find_available_port(start_port=18000, max_attempts=50):
    """尋找本機可用的通訊埠"""
    for p in range(start_port, start_port + max_attempts):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", p)) != 0:
                return p
    return start_port

def open_browser(port):
    time.sleep(1.2)
    webbrowser.open(f"http://127.0.0.1:{port}")

if __name__ == "__main__":
    port = find_available_port(18000)
    threading.Thread(target=open_browser, args=(port,), daemon=True).start()
    print(f"==================================================")
    print(f"  智慧出貨標籤自動套印系統 正在啟動...")
    print(f"  服務網址: http://127.0.0.1:{port}")
    print(f"  請保留此視窗，關閉此視窗將結束服務。")
    print(f"==================================================")
    uvicorn.run(app, host="127.0.0.1", port=port, log_level="warning")
