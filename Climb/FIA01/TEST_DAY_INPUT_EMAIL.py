
import os
import glob
import threading
import time
import tkinter as tk
from tkinter import ttk
from datetime import datetime, timedelta
import pandas as pd
import schedule
import smtplib
from dotenv import load_dotenv
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# === 載入 .env ===
load_dotenv()

# === 設定 Email 通知參數 ===
MTP_SERVER = os.environ.get("SMTP_SERVER")
SMTP_PORT = int(os.environ.get("SMTP_PORT", 587))
EMAIL_SENDER = os.environ.get("EMAIL_SENDER")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD")
EMAIL_RECIPIENTS = os.environ.get("EMAIL_RECIPIENTS", "").split(",")

def send_mail_on_error(subject, body):
    try:
        import smtplib
        from email.mime.text import MIMEText
        from email.mime.multipart import MIMEMultipart

        msg = MIMEMultipart()
        msg["From"] = EMAIL_SENDER
        msg["To"] = ", ".join(EMAIL_RECIPIENTS)
        msg["Subject"] = subject

        msg.attach(MIMEText(body, "plain"))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.send_message(msg)

        log_console("[通知] 已寄出錯誤通知 Email")

    except Exception as e:
        log_console(f"[錯誤] 無法寄出 Email：{str(e)}")

import webbrowser
from ftplib import FTP
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

# === 設定下載資料夾 ===
DOWNLOAD_FOLDER = os.path.join(os.path.expanduser("~"), "Downloads")

# === 是否啟用 Headless 模式 ===
HEADLESS_MODE = False

DAYS_BEFORE = 3 # n天前的資料
is_downloading = False

# === 找出最新檔案 ===
def get_latest_csv():
    pattern = os.path.join(DOWNLOAD_FOLDER, "FIA01_totalReport_*.csv")
    files = glob.glob(pattern)
    files = sorted(files, key=lambda f: os.path.getmtime(f), reverse=True)
    for f in files:
        try:
            with open(f, 'r', encoding='utf-8-sig') as _:
                return f
        except:
            continue
    return None

# === 檔案分析 ===
def analyze_csv(file_path):
    df = pd.read_csv(file_path, encoding="utf-8-sig")
    total_rows = len(df)
    ng_count = 0
    empty_count = 0

    for _, row in df.iterrows():
        values = row.astype(str).values
        row_lower = [v.lower() for v in values]
        is_ng = any("ng" in v for v in row_lower)
        is_empty = any(v.strip() == "" or v.lower() in ["nan", "null"] for v in values)
        if is_ng:
            ng_count += 1
        if is_empty:
            empty_count += 1
    return total_rows, ng_count, empty_count

# === 主工作流程（下載 + 分析）===
def open_folder():
    webbrowser.open(DOWNLOAD_FOLDER)

# === FTP 上傳設定（請依實際填入） ===
    FTP_HOST = "your.ftp.server"   # ← 修改為實際位址
    FTP_PORT = 21                  # ← 可改為 22（SFTP）或其他
    FTP_USER = "your_username"
    FTP_PASS = "your_password"
    FTP_TARGET_DIR = "/"          # ← FTP 目標目錄

# === 上傳檔案到 FTP === #等確定ftp設定再解封
# def upload_latest_csv_to_ftp():
#     try:
#         pattern = os.path.join(DOWNLOAD_FOLDER, "FIA01_totalReport_*.csv")
#         files = glob.glob(pattern)
#         if not files:
#             log_console("[錯誤] 找不到任何 FIA01_totalReport_ 檔案")
#             return

#         latest_file = max(files, key=os.path.getmtime)
#         filename = os.path.basename(latest_file)
#         log_console(f"[任務] 準備上傳檔案：{filename}")

#         ftp = FTP()
#         ftp.connect(FTP_HOST, FTP_PORT, timeout=10)
#         ftp.login(FTP_USER, FTP_PASS)
#         ftp.cwd(FTP_TARGET_DIR)

#         with open(latest_file, "rb") as f:
#             ftp.storbinary(f"STOR {filename}", f)
#         ftp.quit()
#         log_console("[成功] 上傳完成")

#     except Exception as e:
#         log_console(f"[錯誤] FTP 上傳失敗：{str(e)}")
#         send_mail_on_error("FTP 上傳錯誤通知", "FIA01 報表上傳失敗，請立即確認系統狀況。")
#     webbrowser.open(DOWNLOAD_FOLDER)

# === 登入並查詢資料 ===
def setup_driver():
    options = webdriver.ChromeOptions()
    if HEADLESS_MODE:
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')
        options.add_argument('--log-level=3')

    prefs = {
        "download.default_directory": DOWNLOAD_FOLDER,
        "download.prompt_for_download": False,
        "directory_upgrade": True
    }
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=options)
    return driver


def query_report(driver, wait, date_str):
    driver.get("http://192.168.28.37/login")
    wait.until(EC.presence_of_element_located((By.ID, "email"))).send_keys("demo@mail.com")
    driver.find_element(By.ID, "password").send_keys("etic@13098982")
    driver.find_element(By.TAG_NAME, "button").click()

    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="root"]/div/div[1]/ul/li[2]/span'))).click()
    select_report = Select(wait.until(EC.element_to_be_clickable((By.TAG_NAME, "select"))))
    select_report.select_by_visible_text("總報表")
    select_mode = Select(driver.find_elements(By.TAG_NAME, "select")[1])
    select_mode.select_by_visible_text("時間")

    start_input = wait.until(EC.element_to_be_clickable((By.ID, "startDate")))
    start_input.click()
    time.sleep(0.3)
    start_input.send_keys(Keys.CONTROL, "a")
    start_input.send_keys(Keys.DELETE)
    start_input.send_keys(date_str)
    
    # 空點擊
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="root"]/div/div[1]/ul/li[2]/span'))).click()
    time.sleep(0.3)

    end_input = wait.until(EC.element_to_be_clickable((By.ID, "endDate")))
    end_input.click()
    time.sleep(0.3)
    end_input.send_keys(Keys.CONTROL, "a")
    end_input.send_keys(Keys.DELETE)
    end_input.send_keys(date_str)

    # 空點擊
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="root"]/div/div[1]/ul/li[2]/span'))).click()
    time.sleep(0.3)

    # 送出
    driver.find_element(By.XPATH, '//*[@id="root"]/div/div[2]/main/div/table/tbody/tr/td/div/div[1]/div/div[4]/div[1]/button').click()
    
    # 等待「取得資料成功」出現
    WebDriverWait(driver, 10).until(
    EC.text_to_be_present_in_element(
        (By.XPATH, "//*[text()='取得資料成功']"), "取得資料成功"
        # ((By.XPATH, "//span[contains(text(), '取得資料成功')]"))
        )
    )  
    
def get_data_count(driver):
    try:
        raw_text = driver.find_element(By.XPATH, '//*[@id="root"]/div/div[2]/main/div/table/tbody/tr/td/div/div[1]/div/div[4]/div[1]/div').text
        parts = raw_text.strip().split()
        count_text = parts[1] if len(parts) >= 2 else "0"
        return int(count_text)
    except Exception as e:
        log_console(f"[錯誤] 無法讀取資料筆數：{str(e)}")
        return 0

# === 記錄 Console ===
def log_console(msg):
    timestamp = datetime.now().strftime("%m/%d %H:%M:%S")
    txt_console.configure(state='normal')
    txt_console.insert(tk.END, f"[{timestamp}] {msg}\n")
    txt_console.configure(state='disabled')
    txt_console.see(tk.END)
    
# 等待下載完成
def start_job_thread():
    if not is_downloading:
        threading.Thread(target=job, daemon=True).start()
    else:
        log_console("[警告] 上一個下載尚未完成")
        root.update_idletasks()

def job():
    global is_downloading
    is_downloading = True

    try:
        days_before = int(day_before_var.get())
    except:
        days_before = 1
    date_obj = datetime.now() - timedelta(days=days_before)
    date_str = date_obj.strftime("%#m/%#d/%Y") if os.name == 'nt' else date_obj.strftime("%-m/%-d/%Y")
    log_console(f"[任務] 準備下載 {DAYS_BEFORE} 天前資料：{date_obj.strftime('%Y/%m/%d')}")
    root.update_idletasks()

    try:
        driver = setup_driver()
        wait = WebDriverWait(driver, 10)
        query_report(driver, wait, date_str)
        count = get_data_count(driver)

        if count == 0:
            formatted_date = date_obj.strftime("%Y/%m/%d")
            log_console(f"[通知] {formatted_date} 當天無資料")
            driver.quit()
            is_downloading = False
            return

        log_console(f"[訊息] 查詢筆數：{count}，準備下載")

        # 等待 .csv 按鈕出現再點
        try:
            download_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), '.csv')]"))
            )
            driver.execute_script("arguments[0].click();", download_button)
            download_button.click()
            log_console("[任務] 點擊下載按鈕")    
            log_console("[任務] 等待下載完成...")
        except Exception as e:
            log_console(f"[錯誤] 下載過程發生錯誤：{str(e)}")
            send_mail_on_error("下載錯誤通知", "FIA01 報表下載失敗，請立即確認系統狀況。")
        
        driver.quit()
        latest_file = get_latest_csv()
        
        if latest_file:
            log_console(f"[任務] 報表下載完成：{os.path.basename(latest_file)}")
            try:
                total, ng, empty = analyze_csv(latest_file)
                log_console(f"[任務] 分析: {os.path.basename(latest_file)}")
                var_total.set(str(total))
                var_ng.set(str(ng))
                var_empty.set(str(empty))
            except Exception as e:
                log_console(f"[錯誤] 分析 CSV 時發生錯誤：{str(e)}")
        else:
            log_console("[警告] 找不到下載後的最新檔案")
    except Exception as e:
        log_console(f"[錯誤] 發生錯誤：{str(e)}")
    finally:
        is_downloading = False    
        
# === GUI 主畫面 ===
root = tk.Tk()
root.title("FIA01測報分析")
root.geometry("720x480")

txt_console = tk.Text(root, height=12, state='disabled', bg="black", fg="lime")
txt_console.pack(fill=tk.BOTH, padx=10, pady=10, expand=True)

frm_top = ttk.Frame(root)
frm_top.pack(pady=10)

lbl_day_before = ttk.Label(frm_top, text="查詢", font=("Arial", 14))
lbl_day_before.grid(row=0, column=0, padx=(0,5))
day_before_var = tk.StringVar(value="1")
day_before_entry = ttk.Entry(frm_top, textvariable=day_before_var, width=5, font=("Arial", 14))
lbl_suffix = ttk.Label(frm_top, text="天前", font=("Arial", 14))
lbl_suffix.grid(row=0, column=2, padx=(0,10))
day_before_entry.grid(row=0, column=1, padx=(0,15))
btn_download = tk.Button(frm_top, text="下載資料", command=start_job_thread)
btn_download['font'] = ("Noto Sans CJK", 14, "bold")
btn_download.grid(row=0, column=0, padx=10, ipadx=20, ipady=10)

btn_folder = tk.Button(frm_top, text="檔案資料夾", command=open_folder)
btn_folder['font'] = ("Noto Sans CJK", 14, "bold")
btn_folder.grid(row=0, column=4, padx=10, ipadx=10, ipady=10)
# btn_upload = tk.Button(frm_top, text="上傳 FTP", command=upload_latest_csv_to_ftp) # ← 取消註解以啟用上傳
btn_upload = tk.Button(frm_top, text="上傳 FTP", command=lambda: log_console("尚未設定上傳路徑"))
btn_upload['font'] = ("Noto Sans CJK", 14, "bold")
btn_upload.grid(row=0, column=5, padx=10, ipadx=10, ipady=10)

frm_status = ttk.Frame(root)
frm_status.pack(pady=10)

var_total = tk.StringVar(value="0")
var_ng = tk.StringVar(value="0")
var_empty = tk.StringVar(value="0")

ttk.Label(frm_status, text="測試總數", font=("Noto Sans CJK", 14)).grid(row=0, column=0, padx=10)
entry_total = ttk.Entry(frm_status, textvariable=var_total, width=10, state='readonly', font=("Arial", 14))
entry_total.grid(row=0, column=1)

ttk.Label(frm_status, text="NG數量", font=("Noto Sans CJK", 14)).grid(row=0, column=2, padx=10)
entry_ng = ttk.Entry(frm_status, textvariable=var_ng, width=10, state='readonly', font=("Arial", 14))
entry_ng.grid(row=0, column=3)

ttk.Label(frm_status, text="空值數量", font=("Noto Sans CJK", 14)).grid(row=0, column=4, padx=10)
entry_empty = ttk.Entry(frm_status, textvariable=var_empty, width=10, state='readonly', font=("Arial", 14))
entry_empty.grid(row=0, column=5)

root.mainloop()
