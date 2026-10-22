import os
import re
import zipfile
import imaplib
import email
from email.header import decode_header
from datetime import datetime, timedelta
import tempfile
import pandas as pd
from playwright.sync_api import sync_playwright

import win32com.client
import pythoncom

# ==========================================
# 🔑 自動化參數配置 (使用者可依需求修改)
# ==========================================
# 客戶網站設定
MISUMI_URL = "https://prosugate.misumi-ec.com/"

# 網頁元素定位 Selector (若網站 HTML 更新，請在此修改)
SELECTOR_USERNAME_INPUT = '#userNameInput, input[type="email"], input[name="UserName"]'
SELECTOR_PASSWORD_INPUT = '#passwordInput, input[type="password"]'
SELECTOR_LOGIN_BUTTON = '#submitButton, span#submitButton, [id="submitButton"]'

# 左側選單 (優先匹配側邊欄菜單，防範點擊到首頁中間看板卡片)
SELECTOR_MENU_SHIPMENT = 'xpath=(//div[contains(@class, "menu") or contains(@class, "sidebar") or contains(@class, "nav")]//*[text()="出貨"] | //span[text()="出貨"] | //a[text()="出貨"])[1]'
SELECTOR_MENU_SHIPMENT_MANAGE = 'xpath=(//div[contains(@class, "menu") or contains(@class, "sidebar") or contains(@class, "nav")]//*[text()="出貨管理"] | //span[text()="出貨管理"] | //a[text()="出貨管理"])[1]'

# 篩選條件
# 尋找 "供應商出貨日" 文字旁邊的兩個日期輸入框 (限定 input 必須為 text 類型，且排除含有 "訂單" 和 "三住" 的大容器以防誤填)
SELECTOR_START_DATE_INPUT = 'xpath=(//*[(self::span or self::div or self::td or self::label) and contains(., "供應商") and contains(., "出貨日") and not(contains(., "訂單")) and not(contains(., "三住"))]/following-sibling::*//input[not(@type) or @type="text"] | //*[(self::span or self::div or self::td or self::label) and contains(., "供應商") and contains(., "出貨日") and not(contains(., "訂單")) and not(contains(., "三住"))]/..//input[not(@type) or @type="text"])[1]'
SELECTOR_END_DATE_INPUT = 'xpath=(//*[(self::span or self::div or self::td or self::label) and contains(., "供應商") and contains(., "出貨日") and not(contains(., "訂單")) and not(contains(., "三住"))]/following-sibling::*//input[not(@type) or @type="text"] | //*[(self::span or self::div or self::td or self::label) and contains(., "供應商") and contains(., "出貨日") and not(contains(., "訂單")) and not(contains(., "三住"))]/..//input[not(@type) or @type="text"])[2]'
SELECTOR_SEARCH_BUTTON = 'button:has-text("搜索"), text="搜索"'

# 編輯模式與欄位
SELECTOR_EDIT_START_BUTTON = 'xpath=//div[contains(@class, "出貨訊息編輯模式") or contains(., "出貨訊息編輯模式")]//button[contains(., "開始")]'
SELECTOR_EDIT_FINISH_BUTTON = 'xpath=//div[contains(@class, "出貨訊息編輯模式") or contains(., "出貨訊息編輯模式")]//button[contains(., "完成")]'
SELECTOR_TAB_LABEL_COUNT = 'text="訂單標籤列印張數"'
SELECTOR_BATCH_UPDATE_INPUT = 'xpath=//div[contains(., "一次更新指定範圍")]//input'
SELECTOR_BATCH_UPDATE_SET_BUTTON = 'xpath=//div[contains(., "一次更新指定範圍")]//button[contains(., "設定")]'

# 表格勾選與下載
SELECTOR_TABLE_SELECT_ALL = 'xpath=//table//th//input[@type="checkbox"]'
SELECTOR_PRINT_BUTTON = 'button:has-text("印製表單"), text="印製表單"'


def log_step(status_placeholder, text):
    """
    更新 Streamlit UI 的即時 Log 顯示。
    """
    if status_placeholder:
        status_placeholder.code(text)
    try:
        print(text)
    except UnicodeEncodeError:
        # 防止 Windows 終端機 cp950 編碼出錯，自動過濾特殊字元
        try:
            print(text.encode('cp950', errors='replace').decode('cp950'))
        except Exception:
            pass


def download_packing_list_from_email(email_user=None, email_pass=None, status_placeholder=None):
    """
    模組一：自動監聽信件與下載（Email Trigger）
    
    連接本地已登入的 Outlook 桌面版，下載主旨包含指定關鍵字的最新信件 zip 附件，並解壓縮出 Excel 檔案。
    """
    search_subject = "Packing list report for  customer: MISUMI CORPORATION "
    log_step(status_placeholder, "正在連接本地 Outlook 桌面版應用程式...")
    
    # 初始化 COM 庫，防止 Streamlit 多執行緒下的執行期錯誤
    pythoncom.CoInitialize()
    
    try:
        # 使用本地 Windows MAPI 連接 Outlook，免輸入帳密
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        # 6 代表 Inbox (收件匣)
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items
        
        # 依接收時間由新到舊排序
        messages.Sort("[ReceivedTime]", True)
        
        log_step(status_placeholder, f"正在本機收件匣中搜尋主旨包含 '{search_subject}' 的最新信件...")
        
        target_msg = None
        # 遍歷尋找最新的符合主旨信件
        for msg in messages:
            # 確保該項目是 MailItem (Class == 43)
            if msg.Class == 43:
                try:
                    subject = msg.Subject
                    if search_subject in subject:
                        target_msg = msg
                        break
                except Exception:
                    continue
                    
        if not target_msg:
            log_step(status_placeholder, "未找到符合主旨的信件。")
            return None
            
        log_step(status_placeholder, f"成功找到信件: '{target_msg.Subject}' (收到時間: {target_msg.ReceivedTime})")
        
        # 下載附件
        temp_dir = tempfile.mkdtemp()
        excel_path = None
        
        attachments = target_msg.Attachments
        if attachments.Count == 0:
            log_step(status_placeholder, "該信件沒有附件。")
            return None
            
        for i in range(1, attachments.Count + 1):
            att = attachments.Item(i)
            filename = att.FileName
            log_step(status_placeholder, f"發現附件檔: {filename}")
            
            # 如果是 zip 檔案，下載並解壓
            if filename.lower().endswith(".zip"):
                zip_path = os.path.join(temp_dir, filename)
                att.SaveAsFile(zip_path)
                
                log_step(status_placeholder, "正在解壓縮 zip 檔案...")
                with zipfile.ZipFile(zip_path, "r") as zip_ref:
                    zip_ref.extractall(temp_dir)
                    
                # 搜尋解壓後的 Excel 檔
                for file in os.listdir(temp_dir):
                    if file.lower().endswith((".xls", ".xlsx")) and not file.startswith("~$"):
                        excel_path = os.path.join(temp_dir, file)
                        log_step(status_placeholder, f"成功解壓出 Excel 檔案: {file}")
                        break
                break
            # 如果直接是 Excel 檔案
            elif filename.lower().endswith((".xls", ".xlsx")):
                excel_path = os.path.join(temp_dir, filename)
                att.SaveAsFile(excel_path)
                log_step(status_placeholder, f"成功下載 Excel 檔案: {filename}")
                break
                
        return excel_path
        
    except Exception as e:
        log_step(status_placeholder, f"本地 Outlook 讀取失敗: {e}")
        return None
    finally:
        # 釋放 COM
        pythoncom.CoUninitialize()


def run_browser_automation(username, password, start_date_str, end_date_str, box_count, download_dir, headless=True, status_placeholder=None):
    """
    模組二：瀏覽器自動化操作（Browser Automation）
    
    使用 Playwright 登入 MISUMI 系統，填入日期與箱數，並下載產出的標籤 PDF 檔。
    """
    log_step(status_placeholder, "正在啟動瀏覽器自動化引擎 (Playwright)...")
    
    try:
        with sync_playwright() as p:
            # 優先啟動系統的 Microsoft Edge 瀏覽器，否則 fallback 至預設的 Chromium
            try:
                log_step(status_placeholder, "正在嘗試開啟 Microsoft Edge 瀏覽器...")
                browser = p.chromium.launch(channel="msedge", headless=headless)
            except Exception as e:
                log_step(status_placeholder, f"Edge 瀏覽器開啟失敗: {e}。Fallback 至預設 Chromium 核心...")
                browser = p.chromium.launch(headless=headless)
            # 模擬一般螢幕尺寸，防止 RWD 排版跑掉
            context = browser.new_context(viewport={"width": 1280, "height": 800})
            page = context.new_page()
            
            log_step(status_placeholder, f"正在導航至客戶登入頁面...")
            page.goto(MISUMI_URL)
            page.wait_for_load_state("networkidle")
            
            # 1. 檢測登入狀態並進行登入
            log_step(status_placeholder, "正在檢測網頁登入狀態...")
            # 同時等待登入輸入框或首頁的選單按鈕出現，最多等待 10 秒
            try:
                page.wait_for_selector(f'{SELECTOR_USERNAME_INPUT}, {SELECTOR_MENU_SHIPMENT}', timeout=10000)
            except Exception:
                pass
                
            # 判斷是否需要登入
            if page.locator(SELECTOR_USERNAME_INPUT).is_visible():
                log_step(status_placeholder, "檢測到尚未登入，正在填入帳號與密碼進行登入...")
                page.fill(SELECTOR_USERNAME_INPUT, username)
                page.fill(SELECTOR_PASSWORD_INPUT, password)
                page.click(SELECTOR_LOGIN_BUTTON)
                # 等待登入成功並載入主頁
                page.wait_for_load_state("networkidle")
                log_step(status_placeholder, "登入成功！正在切換至『出貨 -> 出貨管理』頁面...")
            else:
                log_step(status_placeholder, "檢測到已處於登入狀態，跳過登入步驟...")
            
            # 2. 導航至出貨管理
            # 等待選單元件渲染完成
            page.wait_for_selector(SELECTOR_MENU_SHIPMENT, timeout=15000)
            # 點擊左側「出貨」展開選單，再點擊「出貨管理」
            page.click(SELECTOR_MENU_SHIPMENT)
            page.wait_for_timeout(500)  # 等待動畫展開
            page.click(SELECTOR_MENU_SHIPMENT_MANAGE)
            page.wait_for_load_state("networkidle")
            
            # 3. 填入「供應商出貨日」起迄區間
            log_step(status_placeholder, f"設定出貨日篩選區間: {start_date_str} 至 {end_date_str}...")
            
            # 等待日期輸入框完全渲染
            page.wait_for_selector(SELECTOR_START_DATE_INPUT, timeout=15000)
            
            # 定位兩個日期輸入框並填入值
            start_input = page.locator(SELECTOR_START_DATE_INPUT)
            end_input = page.locator(SELECTOR_END_DATE_INPUT)
            
            # 強制清空並填入新日期值
            start_input.focus()
            page.keyboard.press("Control+A")
            page.keyboard.press("Backspace")
            start_input.fill(start_date_str)
            
            end_input.focus()
            page.keyboard.press("Control+A")
            page.keyboard.press("Backspace")
            end_input.fill(end_date_str)
            
            # 點擊「搜索」
            log_step(status_placeholder, "點擊『搜索』按鈕中...")
            page.click(SELECTOR_SEARCH_BUTTON)
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(1000) # 等待列表載入完成
            
            # 4. 出貨訊息編輯模式 (開始)
            log_step(status_placeholder, "切換至編輯模式：點擊『開始』按鈕...")
            page.click(SELECTOR_EDIT_START_BUTTON)
            page.wait_for_timeout(500)
            
            # 5. 選擇「訂單標籤列印張數」tab
            log_step(status_placeholder, "點選『訂單標籤列印張數』分頁...")
            page.click(SELECTOR_TAB_LABEL_COUNT)
            page.wait_for_timeout(500)
            
            # 6. 一次更新指定範圍：填入標籤數量 (箱數)，點擊「設定」
            log_step(status_placeholder, f"填入訂單標籤列印張數為: {box_count} 箱，並點擊『設定』...")
            page.fill(SELECTOR_BATCH_UPDATE_INPUT, str(box_count))
            page.click(SELECTOR_BATCH_UPDATE_SET_BUTTON)
            page.wait_for_timeout(800) # 等待批次更新渲染
            
            # 7. 完成編輯模式
            log_step(status_placeholder, "保存設定：點擊編輯模式『完成』按鈕...")
            page.click(SELECTOR_EDIT_FINISH_BUTTON)
            page.wait_for_timeout(1000)
            
            # 8. 全選訂單 checkbox
            log_step(status_placeholder, "勾選所有訂單項目...")
            page.click(SELECTOR_TABLE_SELECT_ALL)
            page.wait_for_timeout(500)
            
            # 9. 點擊「印製表單」並下載 PDF
            log_step(status_placeholder, "正在準備印製表單，開始監聽 PDF 檔案下載...")
            
            with page.expect_download() as download_info:
                page.click(SELECTOR_PRINT_BUTTON)
            
            download = download_info.value
            pdf_filename = f"tacklabel_auto_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf"
            pdf_path = os.path.join(download_dir, pdf_filename)
            
            download.save_as(pdf_path)
            log_step(status_placeholder, f"標籤 PDF 下載成功！儲存路徑為: {pdf_path}")
            
            # 關閉瀏覽器
            browser.close()
            return pdf_path
            
    except Exception as e:
        log_step(status_placeholder, f"網頁自動化執行時發生錯誤: {e}")
        return None
