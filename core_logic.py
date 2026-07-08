import re
import pandas as pd
import fitz
from collections import Counter

def parse_excel(excel_path):
    """
    解析 Packing List Excel 檔案。
    
    回傳格式:
    {
        "53260201049789": [
            {"box_no": "AA0001", "qty": 20, "display_box": "AA", "display_qty": "/20pcs", "seq_str": "1/100"},
            ...
        ]
    }
    """
    # 讀取 Excel 檔案，將所有欄位轉成字串處理，避免浮點數精度或科學記號問題
    df = pd.read_excel(excel_path, dtype=str)
    
    # 找出欄位名稱的對應關係。因為 Packing List 的標題位置可能有些微位移，
    # 我們以含有 'CTN#', 'Po No.', 'Quantity' 的列來作為標題特徵。
    # 這裡我們直接遍歷每一行，並以特定 Regex 特徵來匹配數據行。
    
    raw_boxes = []
    
    # 箱號 Regex：例如 AA0001 或 AA0001 - AA0002
    box_pattern = re.compile(r"^([A-Z]+)(\d+)(?:\s*-\s*([A-Z]+)(\d+))?\s*$")
    
    # PO Regex：14 位純數字，或含有連字號 (e.g. 53-260201-049789)，以及 5P 開頭的 12 位英數字 (e.g. 5P-002A8J-000208)
    po_pattern = re.compile(r"^(?:53-?\d{6}-?\d{6}|5[Pp]-?[A-Za-z0-9]{6}-?[A-Za-z0-9]{6})$")
    
    # 數量 Regex：例如 @30pcs 或 @20pcs
    qty_pattern = re.compile(r"@(\d+)pcs")
    
    for idx, row in df.iterrows():
        # 取出該行的各個單元格值並去除前後空白
        row_vals = [str(val).strip() if pd.notna(val) else "" for val in row.values]
        
        # 尋找非空值
        if len(row_vals) < 6:
            continue
            
        ctn_val = row_vals[0]
        part_val = row_vals[1]
        po_val = row_vals[3]
        qty_val = row_vals[5]
        
        # 檢查是否符合數據行特徵
        ctn_match = box_pattern.match(ctn_val)
        po_match = po_pattern.match(po_val)
        qty_match = qty_pattern.match(qty_val)
        
        if ctn_match and po_match and qty_match:
            po_clean = po_val.replace("-", "")
            qty = int(qty_match.group(1))
            
            # 取得下一行的欄位值，作為 description
            description = ""
            if idx + 1 < len(df):
                next_row = df.iloc[idx + 1]
                next_row_vals = [str(val).strip() if pd.notna(val) else "" for val in next_row.values]
                if len(next_row_vals) > 1:
                    description = next_row_vals[1]
            
            # 解析箱號範圍
            prefix1 = ctn_match.group(1)
            start_num = int(ctn_match.group(2))
            
            prefix2 = ctn_match.group(3)
            end_num = int(ctn_match.group(4)) if ctn_match.group(4) else start_num
            
            # 使用 prefix1 作為前綴，範圍內每一箱展開
            for num in range(start_num, end_num + 1):
                box_no = f"{prefix1}{num:04d}"
                raw_boxes.append({
                    "box_no": box_no,
                    "po": po_clean,
                    "qty": qty,
                    "prefix": prefix1,
                    "num": num,
                    "product_no": part_val,
                    "description": description
                })
                
    # 根據 PO 進行分組
    po_groups = {}
    for box in raw_boxes:
        po = box["po"]
        if po not in po_groups:
            po_groups[po] = []
        po_groups[po].append(box)
        
    # 在每個 PO 群組內計算標示資訊
    processed_data = {}
    for po, boxes in po_groups.items():
        # 依據箱號中的數字排序，確保序號正確
        boxes_sorted = sorted(boxes, key=lambda x: x["num"])
        total_boxes = len(boxes_sorted)
        
        # 計算出現次數最多的數量作為大宗滿箱數量
        qtys = [b["qty"] for b in boxes_sorted]
        majority_qty = Counter(qtys).most_common(1)[0][0]
        
        po_boxes_info = []
        for i, box in enumerate(boxes_sorted):
            seq_num = i + 1  # 1-based 序號
            
            # 決定箱號顯示文字：
            # 如果總箱數不大於 2 箱，或是此箱數量不是大宗滿箱數量，則標示完整箱號；否則僅標示前綴
            if total_boxes <= 2 or box["qty"] != majority_qty:
                display_box = box["box_no"]
            else:
                display_box = box["prefix"]
                
            display_qty = f"/{box['qty']}pcs"
            seq_str = f"{seq_num}/{total_boxes}"
            
            po_boxes_info.append({
                "box_no": box["box_no"],
                "qty": box["qty"],
                "display_box": display_box,
                "display_qty": display_qty,
                "seq_str": seq_str,
                "product_no": box.get("product_no", ""),
                "description": box.get("description", "")
            })
            
        processed_data[po] = po_boxes_info
        
    return processed_data

def get_label_rect(col, row):
    """
    計算給定欄和列的標籤 Rect 區域。
    """
    x0 = 0 if col == 0 else 297
    x1 = 297 if col == 0 else 595
    
    # 根據 Y_AA (100.5, 243.9, 391.5, 536.1, 679.0)
    # 設定高度範圍，每個標籤高約 144
    y0 = row * 144 + 40
    y1 = (row + 1) * 144 + 40
    return fitz.Rect(x0, y0, x1, y1)

def label_pdf(pdf_path, excel_data, output_pdf_path):
    """
    將解析出來的 Excel 資料標示到原始 PDF 檔案上，並另存新檔。
    
    回傳 (success, msg, stats)
    """
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        return False, f"無法開啟 PDF 檔案: {str(e)}", {}
        
    # 用於追蹤各 PO 目前配對到第幾個箱子
    po_counters = {po: 0 for po in excel_data}
    
    # 定義精確的繪製 Y 座標
    Y_AA = [100.5, 243.9, 391.5, 536.1, 679.0]
    Y_PCS = [y + 24 for y in Y_AA]
    Y_P = [187.7, 332.2, 476.8, 621.3, 765.8]
    
    matched_count = 0
    unmatched_labels = 0
    warnings = []
    
    # 逐頁、逐標籤處理
    for page_idx in range(len(doc)):
        page = doc[page_idx]
        
        for row in range(5):
            for col in range(2):
                # 縮小裁剪範圍，只鎖定標籤最上方的 PO 號碼區間 (Y_AA - 40 到 Y_AA + 5)，避開底部條碼
                y_aa = Y_AA[row]
                x0 = 40 if col == 0 else 300
                x1 = 280 if col == 0 else 550
                y0 = y_aa - 40
                y1 = y_aa + 5
                po_rect = fitz.Rect(x0, y0, x1, y1)
                
                # 提取此極窄區域內的文字以識別 PO 號碼
                text = page.get_text("text", clip=po_rect)
                
                # 搜尋 PO 號碼格式：53-260201-049789 或 53260201049789，以及 5P 開頭格式
                po_match = re.search(r"(?:53-?\d{6}-?\d{6})|(?:5[Pp]-?[A-Za-z0-9]{6}-?[A-Za-z0-9]{6})", text)
                
                if po_match:
                    po_clean = po_match.group(0).replace("-", "")
                    
                    if po_clean in excel_data:
                        boxes = excel_data[po_clean]
                        idx = po_counters[po_clean]
                        
                        if idx < len(boxes):
                            box_info = boxes[idx]
                            
                            # 若為 5P 開頭的 PO，為了避開左側長品名文字與下方條碼，
                            # 將 X 座標平移 (左欄 195.0, 右欄 440.0)，且 Y 座標下移 30.0pt，並將箱號與數量合併為單行印出
                            is_5p = po_clean.startswith("5P") or po_clean.startswith("5p")
                            if is_5p:
                                x_aa = 195.0 if col == 0 else 440.0
                                y_aa = Y_AA[row] + 30.0
                            else:
                                x_aa = 205.0 if col == 0 else 450.0
                                y_aa = Y_AA[row]
                            x_pcs = x_aa
                            y_pcs = Y_PCS[row]
                            x_p = 67.3 if col == 0 else 312.1
                            y_p = Y_P[row]
                            
                            # 1. 繪製箱號文字 (AA 或完整箱號如 AC0134)
                            # 若為 5P 標籤，則將箱號與數量合併在同一行印出，防止重疊壓到下方 QTY
                            if is_5p:
                                text_to_draw = f"{box_info['display_box']}  {box_info['display_qty']}"
                            else:
                                text_to_draw = box_info["display_box"]
                                
                            page.insert_text(
                                fitz.Point(x_aa, y_aa + 12),
                                text_to_draw,
                                fontsize=14,
                                fontname="hebo",
                                color=(0, 0, 0)
                            )
                            
                            # 2. 繪製數量文字 (5P 標籤已合併至同一行，故在此跳過)
                            if not is_5p:
                                page.insert_text(
                                    fitz.Point(x_pcs, y_pcs + 12),
                                    box_info["display_qty"],
                                    fontsize=14,
                                    fontname="hebo",
                                    color=(0, 0, 0)
                                )
                            
                            po_counters[po_clean] += 1
                            matched_count += 1
                        else:
                            # 該 PO 的箱子資料已在 Excel 配對完畢，但 PDF 上還有多餘的標籤
                            unmatched_labels += 1
                            warnings.append(f"第 {page_idx+1} 頁標籤 {row*2+col+1}: PO {po_clean} 在 Excel 中的裝箱資料已用完。")
                    else:
                        unmatched_labels += 1
                        warnings.append(f"第 {page_idx+1} 頁標籤 {row*2+col+1}: PO {po_clean} 在 Excel 中無對應的裝箱資料。")
                else:
                    # 該標籤格為空白標籤，跳過不處理
                    pass
                    
    missing_labels = 0
    # 檢查是否所有 Excel 中的箱子都被標示了
    for po, boxes in excel_data.items():
        used = po_counters[po]
        total = len(boxes)
        if used < total:
            missing_labels += (total - used)
            warnings.append(f"警告: Excel 中 PO {po} 共有 {total} 箱，但 PDF 標籤中僅標示了 {used} 箱，有 {total - used} 箱未匹配。")
            
    total_pages = len(doc)
    try:
        doc.save(output_pdf_path)
        doc.close()
    except Exception as e:
        return False, f"儲存處理後的 PDF 失敗: {str(e)}", {}
        
    stats = {
        "total_pages": total_pages,
        "matched_count": matched_count,
        "unmatched_labels": unmatched_labels,
        "missing_labels": missing_labels,
        "warnings": warnings[:15]  # 最多回傳前 15 個警告資訊，避免過長
    }
    
    return True, "成功標示 PDF 標籤！", stats

def find_delivery_date(excel_path):
    """
    從 Packing List Excel 中讀取出貨日期 (Delivery Date)。
    """
    try:
        # 強制用 str 讀取，避免日期格式異常
        df = pd.read_excel(excel_path, header=None, dtype=str)
        # 遍歷前 20 行與前 5 欄尋找關鍵字
        for r_idx in range(min(20, len(df))):
            for c_idx in range(min(5, len(df.columns))):
                val = str(df.iloc[r_idx, c_idx]).strip()
                if "Delivery Date" in val:
                    # 嘗試取得其右側單元格
                    for offset in [1, 2]:
                        if c_idx + offset < len(df.columns):
                            date_str = str(df.iloc[r_idx, c_idx + offset]).strip()
                            if date_str and date_str.lower() != "nan":
                                # 使用 pandas.to_datetime 進行強健的日期格式轉換
                                dt = pd.to_datetime(date_str, errors='coerce')
                                if not pd.isna(dt):
                                    return dt
    except Exception as e:
        print(f"讀取出貨日期時發生錯誤: {e}")
    return None

def is_taiwan_holiday(date_obj):
    """
    判斷指定日期是否為台灣國定假日。
    包含 2026 年與 2027 年主要可能落在週四的假日。
    """
    holidays = {
        # 2026
        "2026-01-01", # 元旦 (週四)
        "2026-02-16", "2026-02-17", "2026-02-18", "2026-02-19", "2026-02-20", # 春節 (2/19 週四)
        "2026-02-27", "2026-02-28", # 228
        "2026-04-02", "2026-04-03", "2026-04-04", "2026-04-05", # 清明/兒童節 (4/2 週四)
        "2026-06-19", # 端午節 (週五)
        "2026-09-25", # 中秋節 (週五)
        "2026-10-09", "2026-10-10", # 國慶日
        
        # 2027
        "2027-01-01", # 元旦
        "2027-02-05", "2027-02-06", "2027-02-07", "2027-02-08", "2027-02-09", "2027-02-10", "2027-02-11", "2027-02-12", # 春節 (2/11 週四)
        "2027-02-28", # 228
        "2027-04-02", "2027-04-03", "2027-04-04", "2027-04-05", # 清明/兒童節
        "2027-06-09", # 端午節 (週三)
        "2027-09-15", # 中秋節 (週三)
        "2027-10-08", "2027-10-10" # 國慶日
    }
    date_str = date_obj.strftime("%Y-%m-%d")
    return date_str in holidays

def calculate_date_range(start_date):
    """
    根據出貨日計算篩選日期區間。
    開始日期 = 出貨日
    結束日期 = 下週四 (若遇台灣國定假日則改為週三)
    """
    from datetime import timedelta
    if not start_date:
        return None, None
    
    # 計算下週四：(10 - 星期幾) % 7 如果為 0 則代表下週四
    # 星期一=0, 星期二=1, ..., 星期日=6
    # 我們使用公式： D + timedelta(days = 10 - D.weekday())
    # 這樣一律可以找到 start_date 所在日曆週下星期的星期四
    days_to_thursday = 10 - start_date.weekday()
    next_thursday = start_date + timedelta(days=days_to_thursday)
    
    # 檢查是否為台灣假日，若是則改為週三 (前移 1 天)
    if is_taiwan_holiday(next_thursday):
        end_date = next_thursday - timedelta(days=1)
    else:
        end_date = next_thursday
        
    return start_date, end_date
