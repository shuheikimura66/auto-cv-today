import os
import json
import time
import glob
import csv
from datetime import datetime, timedelta, timezone
from urllib.parse import quote
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# --- 環境変数 ---
USER_ID = os.environ["USER_ID"]
PASSWORD = os.environ["USER_PASS"]
json_creds = json.loads(os.environ["GCP_JSON"])

# --- 設定 ---
# 参考コードのURL (Action Log)
TARGET_URL = "https://asp1.six-pack.xyz/admin/log/action/list"

# 転記先スプレッドシート設定
SPREADSHEET_ID = "1_nbkQfF-8vlQVkBBVf7cUD5k8ymTdMn-hN28VImm8j0"
SHEET_NAME = "raw_cv_当日"

def get_google_service(service_name, version):
    """Google APIサービスを取得するヘルパー関数"""
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_info(json_creds, scopes=scopes)
    return build(service_name, version, credentials=creds)

def update_google_sheet(csv_path):
    """CSVの中身を読み込んでスプレッドシートに張り付ける関数"""
    print(f"スプレッドシートへの転記を開始: {SHEET_NAME}")
    service = get_google_service('sheets', 'v4')

    # 1. CSVデータの読み込み (文字コード判定付き)
    csv_data = []
    try:
        # まずUTF-8で試行
        with open(csv_path, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            csv_data = list(reader)
    except UnicodeDecodeError:
        print("UTF-8での読み込みに失敗しました。Shift_JIS(CP932)で再試行します。")
        try:
            # 失敗したらShift_JIS(CP932)で試行 (日本のASPによくある形式)
            with open(csv_path, 'r', encoding='cp932') as f:
                reader = csv.reader(f)
                csv_data = list(reader)
        except Exception as e:
            print(f"CSV読み込みエラー: {e}")
            return

    if not csv_data:
        print("CSVデータが空のため転記をスキップします。")
        return

    # 2. シートのクリア (古いデータを消す)
    try:
        service.spreadsheets().values().clear(
            spreadsheetId=SPREADSHEET_ID,
            range=SHEET_NAME
        ).execute()
        print("既存データをクリアしました。")
    except Exception as e:
        print(f"シートクリアエラー: {e}")

    # 3. データの書き込み
    body = {
        'values': csv_data
    }
    try:
        result = service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET_NAME}!A1",
            valueInputOption='USER_ENTERED',
            body=body
        ).execute()
        print(f"スプレッドシート更新完了: {result.get('updatedCells')} セル更新")
    except Exception as e:
        print(f"書き込みエラー: {e}")

def get_today_jst():
    """日本時間の【当日】を計算して文字列(YYYY年MM月DD日)で返す"""
    JST = timezone(timedelta(hours=+9), 'JST')
    now = datetime.now(JST)
    return now.strftime("%Y年%m月%d日")

def input_date_range(driver, wait, label_text, date_str):
    """日付範囲を入力する共通関数"""
    try:
        full_date_str = f"{date_str} - {date_str}"
        print(f"「{label_text}」に日付を入力します: {full_date_str}")
        
        label_elem = wait.until(EC.presence_of_element_located((By.XPATH, f"//*[contains(text(), '{label_text}')]")))
        input_elem = label_elem.find_element(By.XPATH, "./following::input[1]")
        
        driver.execute_script("arguments[0].click();", input_elem)
        driver.execute_script("arguments[0].value = '';", input_elem)
        time.sleep(0.5)
        
        input_elem.send_keys(full_date_str)
        input_elem.send_keys(Keys.ENTER)
        time.sleep(1)
    except Exception as e:
        print(f"日付入力エラー({label_text}): {e}")

def main():
    print("=== Action Log取得処理開始(当日分) ===")
    
    download_dir = os.path.join(os.getcwd(), "downloads_action")
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)

    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--window-size=1920,1080')
    
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 20)

    try:
        # --- 1. ログイン ---
        safe_user = quote(USER_ID, safe='')
        safe_pass = quote(PASSWORD, safe='')
        url_body = TARGET_URL.replace("https://", "").replace("http://", "")
        auth_url = f"https://{safe_user}:{safe_pass}@{url_body}"
        
        print(f"アクセス中: {TARGET_URL}")
        driver.get(auth_url)
        time.sleep(3)

        # --- 2. 画面リフレッシュ ---
        print("画面を再読み込みします...")
        driver.get(auth_url)
        time.sleep(5) 

        # --- 3. 「絞り込み検索」ボタンをクリック ---
        print("検索メニューを開きます...")
        try:
            filter_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), '絞り込み検索')]")))
            filter_btn.click()
            time.sleep(3)
        except:
            pass

        # --- 4. 検索条件入力 (当日日付) ---
        today_str = get_today_jst()
        input_date_range(driver, wait, "登録日時", today_str)
        input_date_range(driver, wait, "承認日時", today_str)

        print("パートナーを入力します...")
        try:
            partner_label = driver.find_element(By.XPATH, "//div[contains(text(), 'パートナー')] | //label[contains(text(), 'パートナー')]")
            partner_target = partner_label.find_element(By.XPATH, "./following::input[contains(@placeholder, '選択')][1]")
            partner_target.click()
            time.sleep(1)
            
            active_elem = driver.switch_to.active_element
            active_elem.send_keys("株式会社フルアウト")
            time.sleep(3) 
            active_elem.send_keys(Keys.ENTER)
            print("パートナーを選択しました")
            time.sleep(2)

        except Exception as e:
            print(f"パートナー入力エラー: {e}")

        # --- 5. 詳細項目 > クリック時リファラ ---
        print("詳細項目を設定します...")
        try:
            # 「詳細項目」を開く
            detail_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), '詳細項目')]")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", detail_btn)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", detail_btn)
            print("「詳細項目」をクリックしました")
            
            time.sleep(3) # メニュー展開待ち

            target_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@value='clickReferrer']")))
            target_div = target_input.find_element(By.XPATH, "./..")
            
            current_status = target_div.get_attribute("aria-checked")
            
            if current_status == "false":
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target_div)
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", target_div)
                print("「クリック時リファラ」をONにしました")
            else:
                print("「クリック時リファラ」は既にONのため操作しません")
            
            time.sleep(1)

        except Exception as e:
            print(f"詳細項目の設定でエラー: {e}")

        # --- 6. 検索ボタン実行 ---
        print("検索ボタンを探して押します...")
        try:
            search_btns = driver.find_elements(By.XPATH, "//input[@value='検索'] | //button[contains(text(), '検索')]")
            target_search_btn = None
            for btn in search_btns:
                if btn.is_displayed():
                    target_search_btn = btn
            
            if target_search_btn:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target_search_btn)
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", target_search_btn)
                print("検索ボタンをクリックしました")
            else:
                webdriver.ActionChains(driver).send_keys(Keys.ENTER).perform()

        except Exception as e:
            print(f"検索ボタン操作エラー: {e}")
            webdriver.ActionChains(driver).send_keys(Keys.ENTER).perform()
        
        # --- 検索結果の反映待ち ---
        print("検索結果を待機中...")
        time.sleep(15)

        # --- 7. CSV生成ボタン ---
        print("CSV生成ボタンを押します...")
        try:
            csv_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@value='CSV生成' or contains(text(), 'CSV生成')]")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", csv_btn)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", csv_btn)
            print("CSV生成ボタンをクリックしました")
            
        except Exception as e:
            print(f"CSVボタンエラー: {e}")
            return
        
        # ダウンロード待ち
        print("ダウンロード待機中...")
        time.sleep(8)
        for i in range(30):
            files = glob.glob(os.path.join(download_dir, "*.csv"))
            if files:
                break
            time.sleep(3)
            
        files = glob.glob(os.path.join(download_dir, "*.csv"))
        if not files:
            print("【エラー】CSVファイルが見つかりません。")
            return
        
        csv_file_path = files[0]
        print(f"ダウンロード成功: {csv_file_path}")

        # --- 8. スプレッドシートへ転記 ---
        update_google_sheet(csv_file_path)

    except Exception as e:
        print(f"【エラー発生】: {e}")
        import traceback
        traceback.print_exc()
        
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
