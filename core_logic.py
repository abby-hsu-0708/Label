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
    
    # PO Regex：支援所有 5 開頭、後接英數字、中間含可選連字號之 14 位 PO 號碼 (如 53, 5P, 5K 等)
    po_pattern = re.compile(r"^5[A-Za-z0-9]-?[A-Za-z0-9]{6}-?[A-Za-z0-9]{6}$")
    
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
                
                # 搜尋 PO 號碼格式：包含 53, 5P, 5K 等所有 5 開頭的 14 位 PO 格式
                po_match = re.search(r"5[A-Za-z0-9]-?[A-Za-z0-9]{6}-?[A-Za-z0-9]{6}", text)
                
                if po_match:
                    po_clean = po_match.group(0).replace("-", "")
                    
                    if po_clean in excel_data:
                        boxes = excel_data[po_clean]
                        idx = po_counters[po_clean]
                        
                        if idx < len(boxes):
                            box_info = boxes[idx]
                            
                            # 若為非 53 開頭的特殊 PO (如 5P, 5K 等)，為了避開左側長品名文字與下方條碼，
                            # 將 X 座標平移 (左欄 195.0, 右欄 440.0)，且 Y 座標下移 30.0pt，並將箱號與數量合併為單行印出
                            is_special_layout = not po_clean.startswith("53")
                            if is_special_layout:
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
                            # 若為特殊標籤，則將箱號與數量合併在同一行印出，防止重疊壓到下方 QTY
                            if is_special_layout:
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
                            
                            # 2. 繪製數量文字 (特殊標籤已合併至同一行，故在此跳過)
                            if not is_special_layout:
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
    判斷指定日期是否為台灣國定假日或週末。
    包含 2025、2026、2027 年主要國定假日與彈性放假。
    """
    # 週末即為休假日 (5: 週六, 6: 週日)
    if date_obj.weekday() in (5, 6):
        return True

    holidays = {
        # 2025
        "2025-01-01", # 元旦
        "2025-01-27", "2025-01-28", "2025-01-29", "2025-01-30", "2025-01-31", # 春節
        "2025-02-28", # 228
        "2025-04-03", "2025-04-04", # 清明兒童連假
        "2025-05-01", # 勞動節
        "2025-05-30", "2025-05-31", # 端午節
        "2025-10-06", # 中秋節
        "2025-10-10", # 國慶日

        # 2026
        "2026-01-01", "2026-01-02", # 元旦連假
        "2026-02-15", "2026-02-16", "2026-02-17", "2026-02-18", "2026-02-19", "2026-02-20", # 春節
        "2026-02-27", "2026-02-28", # 228連假
        "2026-04-02", "2026-04-03", "2026-04-04", "2026-04-05", "2026-04-06", # 清明/兒童節連假
        "2026-05-01", # 勞動節 (週五)
        "2026-06-19", # 端午節 (週五)
        "2026-09-25", # 中秋節 (週五)
        "2026-10-09", "2026-10-10", # 國慶日連假

        # 2027
        "2027-01-01", # 元旦 (週五)
        "2027-02-05", "2027-02-06", "2027-02-07", "2027-02-08", "2027-02-09", "2027-02-10", "2027-02-11", "2027-02-12", # 春節
        "2027-02-26", "2027-02-27", "2027-02-28", # 228連假
        "2027-04-02", "2027-04-03", "2027-04-04", "2027-04-05", # 清明/兒童節
        "2027-04-30", "2027-05-01", # 勞動節
        "2027-06-09", # 端午節
        "2027-09-15", # 中秋節
        "2027-10-08", "2027-10-10" # 國慶日連假
    }
    date_str = date_obj.strftime("%Y-%m-%d")
    return date_str in holidays

def calculate_date_range(start_date):
    """
    根據出貨日計算客戶出貨系統的篩選日期區間。
    業務規則：
    1. 起始日 (Start Date)：當次出貨日。
    2. 結束日 (End Date)：下次出貨日的前一天。
       - 公司固定週五出貨；若遇假日則提前一天或兩天（例如週四、週三）。
       - 下次出貨日預設為下週五，若休假則提前至最近工作日。
       - 結束日 = 下次出貨日 - 1 天。
       例如：10/16 (週五) 出貨，下次出貨為 10/23 (週五)，結束日即為 10/22 (週四)，區間為 10/16 ~ 10/22。
    """
    from datetime import timedelta
    if not start_date:
        return None, None

    # 1. 計算下一個日曆週的週五 (weekday 4)
    # 若 start_date 為週一(0) ~ 週日(6)，計算距離下週五的天數
    days_to_next_friday = 4 - start_date.weekday()
    if days_to_next_friday <= 0:
        days_to_next_friday += 7

    target_friday = start_date + timedelta(days=days_to_next_friday)

    # 2. 判斷下週五是否遇台灣國定假日，若休假則持續往前推算至上班日
    next_shipment = target_friday
    while is_taiwan_holiday(next_shipment):
        next_shipment -= timedelta(days=1)

    # 3. 結束日為下次出貨日的前一天
    end_date = next_shipment - timedelta(days=1)

    # 若結束日小於起始日（極端罕見情況），保底設為起始日當天
    if end_date < start_date:
        end_date = start_date

    return start_date, end_date

def generate_portal_bookmarklet(start_date, end_date, excel_data):
    """
    根據出貨日期區間與解析後的 PO 箱數資料，生成適用於客戶平台的「瀏覽器一鍵注入填表腳本 (Bookmarklet)」。
    
    回傳:
        raw_js: 純 JavaScript 腳本字串 (可用於 F12 開發者工具 Console 貼上執行)
        bookmarklet_url: 適用於瀏覽器書籤列的 javascript:... 協定 URL
        summary: 統計資訊字典
    """
    import json
    import urllib.parse
    
    # 格式化日期為 YYYY/MM/DD
    start_str = start_date.strftime("%Y/%m/%d") if hasattr(start_date, "strftime") else str(start_date)
    end_str = end_date.strftime("%Y/%m/%d") if hasattr(end_date, "strftime") else str(end_date)
    
    # 建立 PO -> 箱數 (需求標籤張數) 字典
    # 同時收錄純英數與含連字號格式，確保 100% 匹配
    po_map = {}
    for po, boxes in excel_data.items():
        box_count = len(boxes)
        po_clean = str(po).replace("-", "").strip()
        po_map[po_clean] = box_count
        
        # 若為標準 14 碼格式，額外建立含破折號版本
        if len(po_clean) == 14:
            dashed_po = f"{po_clean[:2]}-{po_clean[2:8]}-{po_clean[8:]}"
            po_map[dashed_po] = box_count

    po_map_json = json.dumps(po_map, ensure_ascii=False)

    # 純原生 JavaScript 注入代碼
    raw_js = f"""(function() {{
    const startDate = "{start_str}";
    const endDate = "{end_str}";
    const poMap = {po_map_json};

    function showHUD(msg, type = "info") {{
        let hud = document.getElementById("shipping-helper-hud");
        if (!hud) {{
            hud = document.createElement("div");
            hud.id = "shipping-helper-hud";
            hud.style.cssText = "position:fixed;top:24px;right:24px;z-index:999999;background:#111827;color:#f3f4f6;padding:18px 22px;border-radius:14px;box-shadow:0 12px 30px rgba(0,0,0,0.6);font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;font-size:14px;border:1.5px solid #374151;max-width:420px;line-height:1.6;transition:all 0.3s ease;";
            document.body.appendChild(hud);
        }}
        const colors = {{ info: "#3b82f6", success: "#10b981", warn: "#f59e0b", error: "#ef4444" }};
        const titleColors = {{ info: "#60a5fa", success: "#34d399", warn: "#fbbf24", error: "#f87171" }};
        hud.style.borderColor = colors[type] || colors.info;
        hud.innerHTML = `
            <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;border-bottom:1px solid rgba(255,255,255,0.1);padding-bottom:6px;">
                <strong style="color:${{titleColors[type] || '#fff'}};font-size:15px;display:flex;align-items:center;gap:6px;">
                    🏷️ 出貨標籤自動填表小幫手
                </strong>
                <span onclick="this.parentElement.parentElement.remove()" style="cursor:pointer;color:#9ca3af;font-size:18px;line-height:1;padding:2px 6px;">&times;</span>
            </div>
            <div style="color:#e5e7eb;">${{msg}}</div>
        `;
        if (type === "success") {{
            setTimeout(() => {{
                if (hud) {{
                    hud.style.opacity = '0';
                    setTimeout(() => hud.remove(), 500);
                }}
            }}, 12000);
        }}
    }}

    function triggerInputEvents(el, val) {{
        el.focus();
        el.value = val;
        el.dispatchEvent(new Event('input', {{ bubbles: true }}));
        el.dispatchEvent(new Event('change', {{ bubbles: true }}));
        el.dispatchEvent(new Event('blur', {{ bubbles: true }}));
    }}

    function sleep(ms) {{
        return new Promise(r => setTimeout(r, ms));
    }}

    async function runAutoFill() {{
        showHUD(`🚀 <b>啟動自動填表流程...</b><br>📅 出貨日篩選區間：<b>${{startDate}}</b> ~ <b>${{endDate}}</b>`);
        await sleep(400);

        // 1. 尋找「出貨」->「出貨管理」頁籤 (若不在當前頁)
        const shipmentLink = Array.from(document.querySelectorAll('a, span, div')).find(el => el.textContent.trim() === '出貨管理');
        if (shipmentLink && !window.location.href.includes('shipment')) {{
            showHUD("正在切換至『出貨管理』頁面...");
            shipmentLink.click();
            await sleep(1500);
        }}

        // 2. 尋找「供應商出貨日」起訖輸入框
        showHUD("📅 正在自動填入出貨日篩選區間...");
        const allElements = Array.from(document.querySelectorAll('span, div, td, label'));
        const dateHeader = allElements.find(el => {{
            const t = el.textContent || "";
            return t.includes("供應商") && t.includes("出貨日") && !t.includes("訂單") && !t.includes("三住");
        }});

        let dateInputs = [];
        if (dateHeader) {{
            const container = dateHeader.closest('tr') || dateHeader.closest('div') || dateHeader.parentElement;
            if (container) {{
                dateInputs = Array.from(container.querySelectorAll('input[type="text"], input:not([type])'));
            }}
        }}

        // Fallback 策略：若找不到文字標籤，找所有帶有日期相關特徵的 input
        if (dateInputs.length < 2) {{
            dateInputs = Array.from(document.querySelectorAll('input[type="text"], input:not([type])')).filter(i => {{
                return (i.placeholder && (i.placeholder.includes('年') || i.placeholder.includes('YYYY') || i.placeholder.includes('/'))) || 
                       (i.value && (i.value.includes('/') || i.value.includes('-')));
            }});
        }}

        if (dateInputs.length >= 2) {{
            triggerInputEvents(dateInputs[0], startDate);
            triggerInputEvents(dateInputs[1], endDate);
        }} else {{
            console.warn("未精確鎖定日期輸入框，嘗試直接進入列表比對");
        }}

        // 3. 點擊「搜索」按鈕
        showHUD("🔍 點擊『搜索』查詢訂單清單中...");
        const searchBtn = Array.from(document.querySelectorAll('button, input[type="button"], a')).find(el => {{
            const t = (el.textContent || el.value || "").trim();
            return t === "搜索" || t === "搜尋";
        }});
        if (searchBtn) {{
            searchBtn.click();
            await sleep(2200);
        }}

        // 4. 切換出貨訊息編輯模式 (點擊開始)
        showHUD("⚙️ 啟動『出貨訊息編輯模式』...");
        const editStartBtn = Array.from(document.querySelectorAll('button, input[type="button"]')).find(el => {{
            const t = (el.textContent || el.value || "").trim();
            return t === "開始" || t.includes("開始編輯");
        }});

        if (editStartBtn) {{
            editStartBtn.click();
            await sleep(800);
        }}

        // 5. 點選「訂單標籤列印張數」頁籤
        showHUD("📑 切換至『訂單標籤列印張數』分頁...");
        const labelTab = Array.from(document.querySelectorAll('div, span, a, button, li')).find(el => {{
            return (el.textContent || "").includes("訂單標籤列印張數");
        }});
        if (labelTab) {{
            labelTab.click();
            await sleep(700);
        }}

        // 6. 遍歷表格每一列，比對 PO 號碼並填入箱數
        showHUD("📦 正在逐列比對 PO 號碼並填寫標籤箱數...");
        let matchedRows = 0;
        const rows = Array.from(document.querySelectorAll('table tbody tr, table tr'));

        rows.forEach(row => {{
            const rowText = (row.textContent || "").replace(/[^0-9A-Za-z]/g, "");
            for (const [poKey, qty] of Object.entries(poMap)) {{
                const cleanKey = poKey.replace(/[^0-9A-Za-z]/g, "");
                if (rowText.includes(cleanKey)) {{
                    const inputs = Array.from(row.querySelectorAll('input[type="text"], input[type="number"], input:not([type])')).filter(i => !i.disabled && !i.readOnly);
                    if (inputs.length > 0) {{
                        // 取最後一個或最符合數量的輸入框
                        const targetInput = inputs[inputs.length - 1];
                        triggerInputEvents(targetInput, qty);
                        targetInput.style.backgroundColor = "#dcfce7";
                        targetInput.style.border = "2px solid #16a34a";
                        targetInput.style.fontWeight = "bold";
                        matchedRows++;
                        break;
                    }}
                }}
            }}
        }});

        // 7. 保存編輯設定：點擊「完成」
        showHUD("💾 正在儲存設定 (點擊編輯模式『完成』)...");
        const editFinishBtn = Array.from(document.querySelectorAll('button, input[type="button"]')).find(el => {{
            const t = (el.textContent || el.value || "").trim();
            return t === "完成" || t.includes("完成編輯");
        }});

        if (editFinishBtn) {{
            editFinishBtn.click();
            await sleep(1000);
        }}

        // 8. 全選訂單 Checkbox
        showHUD("☑️ 正在勾選所有訂單項目...");
        const selectAllCb = document.querySelector('table th input[type="checkbox"], thead input[type="checkbox"], th input[type="checkbox"]');
        if (selectAllCb) {{
            if (!selectAllCb.checked) {{
                selectAllCb.click();
                selectAllCb.dispatchEvent(new Event('change', {{ bubbles: true }}));
            }}
        }} else {{
            // Fallback: 逐一勾選每行的 checkbox
            const itemCheckboxes = Array.from(document.querySelectorAll('table tbody input[type="checkbox"], td input[type="checkbox"]'));
            itemCheckboxes.forEach(cb => {{
                if (!cb.checked) {{
                    cb.click();
                    cb.dispatchEvent(new Event('change', {{ bubbles: true }}));
                }}
            }});
        }}

        await sleep(500);

        // 9. 顯示最終完成通知
        showHUD(`
            🎉 <b>自動填表與勾選全部完成！</b><br>
            📅 篩選出貨日：<b>${{startDate}}</b> ~ <b>${{endDate}}</b><br>
            📦 成功填入標籤箱數：<b>${{matchedRows}}</b> 筆訂單<br>
            ☑️ 所有訂單項目已自動全選。<br><br>
            👉 <b>請確認畫面無誤後，直接點擊【印製表單】下載標籤 PDF！</b>
        `, "success");
    }}

    runAutoFill();
}})();"""

    # 將 JavaScript 進行安全壓縮與 URL 編碼，製成 Bookmarklet
    minified_js = " ".join([line.strip() for line in raw_js.splitlines() if line.strip() and not line.strip().startswith("//")])
    bookmarklet_url = f"javascript:{urllib.parse.quote(minified_js)}"

    summary = {
        "start_date": start_str,
        "end_date": end_str,
        "po_count": len(po_map) // 2 if any("-" in k for k in po_map) else len(po_map),
        "po_map": po_map
    }

    return raw_js, bookmarklet_url, summary
