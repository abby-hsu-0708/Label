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
    
    # PO Regex：14 位純數字，或含有連字號 (e.g. 53-260201-049789)
    po_pattern = re.compile(r"^53-?\d{6}-?\d{6}$")
    
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
                    "num": num
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
                "seq_str": seq_str
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
                
                # 搜尋 PO 號碼格式：53-260201-049789 或 53260201049789
                po_match = re.search(r"53-?\d{6}-?\d{6}", text)
                
                if po_match:
                    po_clean = po_match.group(0).replace("-", "")
                    
                    if po_clean in excel_data:
                        boxes = excel_data[po_clean]
                        idx = po_counters[po_clean]
                        
                        if idx < len(boxes):
                            box_info = boxes[idx]
                            
                            # 確定繪製的座標：箱號與數量對齊下方 QTY: 起點 (QTY: 起點約為 205.7 / 450.5)
                            x_aa = 205.0 if col == 0 else 450.0
                            x_pcs = x_aa
                            x_p = 67.3 if col == 0 else 312.1
                            
                            y_aa = Y_AA[row]
                            y_pcs = Y_PCS[row]
                            y_p = Y_P[row]
                            
                            # 1. 繪製箱號文字 (AA 或完整箱號如 AC0134)
                            # 使用 Helvetica-Bold 字型，大小約 14pt (與原版完美契合)
                            page.insert_text(
                                fitz.Point(x_aa, y_aa + 12),
                                box_info["display_box"],
                                fontsize=14,
                                fontname="hebo",
                                color=(0, 0, 0)
                            )
                            
                            # 2. 繪製數量文字 (e.g. /20pcs)
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
                    
    # 檢查是否所有 Excel 中的箱子都被標示了
    for po, boxes in excel_data.items():
        used = po_counters[po]
        total = len(boxes)
        if used < total:
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
        "warnings": warnings[:15]  # 最多回傳前 15 個警告資訊，避免過長
    }
    
    return True, "成功標示 PDF 標籤！", stats
