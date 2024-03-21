from selenium import webdriver
import chromedriver_binary
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.select import Select
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException, TimeoutException, ElementNotInteractableException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.remote.webdriver import WebDriver

import requests
import psutil
import re
import time
from urllib.parse import quote
from bs4 import BeautifulSoup
from urllib.parse import urlparse, urljoin
import csv
from datetime import datetime
import os
import ctypes
from ctypes import wintypes
import traceback
import threading

from tkinter import filedialog
import tkinter as tk
from tkinter import Tk
import customtkinter as ctk
from tkinter import messagebox

# GUID for Desktop folder
FOLDERID_Desktop = "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"

# Define the GUID structure
class GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", wintypes.DWORD),
        ("Data2", wintypes.WORD),
        ("Data3", wintypes.WORD),
        ("Data4", wintypes.BYTE * 8)
    ]

def get_desktop_path():
    # Function prototype
    SHGetKnownFolderPath = ctypes.windll.shell32.SHGetKnownFolderPath
    SHGetKnownFolderPath.argtypes = [
        ctypes.POINTER(GUID),
        wintypes.DWORD,
        wintypes.HANDLE,
        ctypes.POINTER(wintypes.LPWSTR)
    ]
    
    path = wintypes.LPWSTR()
    folderid = GUID()
    ctypes.windll.ole32.CLSIDFromString(FOLDERID_Desktop, ctypes.byref(folderid))
    result = SHGetKnownFolderPath(ctypes.byref(folderid), 0, None, ctypes.byref(path))
    
    if result == 0:  # S_OK
        return path.value
    else:
        raise Exception("Failed to get desktop path")

# 指定したプロセス名のすべてのプロセスを終了します。
def kill_process_by_name(process_name):
    for proc in psutil.process_iter(attrs=["pid", "name"]):
        if process_name in proc.info["name"]:
            psutil.Process(proc.info["pid"]).terminate()


def is_driver_active(driver):
    try:
        # driverのsession_idとcurrent_urlにアクセスしてみる
        # WebDriverが閉じられている場合、これらの操作はエラーを引き起こす
        _ = driver.session_id
        _ = driver.current_url
        return True
    except:
        return False

driver = None
# ChromeのWebDriverのインスタンスを生成します。
# もしエラーが発生した場合、"chromedriver"プロセスを終了して再度試みます。
def create_chrome_driver():
    global driver
    if driver is not None and is_driver_active(driver):
        return driver
    try:
        # Chromeのログレベルを変更するオプションを追加
        # Chromeのオプションを設定
        chrome_options = Options()
        #chrome_options.add_argument("--headless")  # ヘッドレスモードを有効化
        chrome_options.add_argument('--log-level=3')  # 致命的なエラーのみ表示
        #chrome_options.add_argument('--disable-gpu')  # GPUハードウェアアクセラレーションを無効化（一部の環境で必要）
        #chrome_options.add_argument('--window-size=1920x1080')  # ウィンドウサイズ指定
        chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])  # DevToolsのログを無効化
        # ChromeDriverManagerを使用してChromeDriverを自動で管理
        driver = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)
        return driver
    except Exception as e:
        log_message(f"エラーが発生しました: {e}")
        # エラー処理が必要な場合はここに記述
        # kill_process_by_name("chromedriver")  # 必要に応じて実装
        # エラー発生時の処理を再度試行する場合は、再度driverのインスタンスを生成
        driver = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)
        return driver


def get_text_by_id(driver, element_id, timeout=60, scroll_amount=300):
    start_time = time.time()
    while True:
        if time.time() - start_time > timeout:
            raise Exception(f"Timeout on waiting for element with ID '{element_id}'.")
        try:
            # 指定されたIDを持つ要素を見つける
            element = driver.find_element(By.ID, element_id)
            # 要素のテキストを取得して返す
            return element.text
        except:
            # 要素が見つからない場合は指定された量だけスクロールして再試行
            driver.execute_script(f"window.scrollBy(0, {scroll_amount});")


def get_text_by_class(driver, class_name, timeout=60, scroll_amount=300):
    start_time = time.time()
    while True:
        if time.time() - start_time > timeout:
            raise Exception(f"Timeout on waiting for element with class name '{class_name}'.")
        try:
            element = driver.find_element(By.CLASS_NAME, class_name)
            return element.text
        except:
            driver.execute_script(f"window.scrollBy(0, {scroll_amount});")
            continue

def get_text_by_class_and_index(driver, class_name, index=0, timeout=60, scroll_amount=300):
    start_time = time.time()
    while True:
        if time.time() - start_time > timeout:
            raise Exception(f"Timeout on waiting for elements with class name '{class_name}'.")
        try:
            elements = driver.find_elements(By.CLASS_NAME, class_name)
            if elements and len(elements) > index:
                return elements[index].text
            # 指定されたインデックスの要素が見つからない場合は、スクロールして再試行
            driver.execute_script(f"window.scrollBy(0, {scroll_amount});")
        except Exception as e:
            raise Exception(f"Error occurred while trying to find elements by class name '{class_name}': {e}")

def click_element_with_exact_text(driver, text, timeout=60, scroll_amount=300):
    start_time = time.time()
    while True:
        if time.time() - start_time > timeout:
            raise Exception(f"Timeout on waiting for element with text '{text}'.")
        try:
            # normalize-space()を使用して空白を無視
            element = driver.find_element(By.XPATH, f"//*[normalize-space(text()) = '{text}']")
            element.click()
            break
        except NoSuchElementException:
            # 要素が見つからない場合はスクロール
            driver.execute_script(f"window.scrollBy(0, {scroll_amount});")
        except ElementClickInterceptedException:
            # 要素がクリック可能になるまで少し待つ
            time.sleep(1)
        except:
            # その他の例外が発生した場合
            continue

def extract_asin(url):
    pattern = r'/dp/([A-Z0-9]{10})'
    match = re.search(pattern, url)
    if match:
        return match.group(1)
    else:
        return None

def get_amazon_product_info(url, retry=True):
    try:
        # WebDriverの設定と初期化
        driver = create_chrome_driver()  # create_chrome_driver 関数を適切に定義する必要があります
        driver.get(url)
        time.sleep(2)
        asin = extract_asin(driver.current_url)  # extract_asin 関数を適切に定義する必要があります

        # 商品のタイトルと価格を取得
        search_query = get_text_by_id(driver, "productTitle", timeout=10, scroll_amount=0)  # get_text_by_id 関数を適切に定義する必要があります
        price = ""
        for i in range(100):
            price = get_text_by_class_and_index(driver, "a-price-whole", i, timeout=10, scroll_amount=0)  # get_text_by_class_and_index 関数を適切に定義する必要があります
            if price != "":
                break

        # ブラウザを閉じる
        #driver.quit()

        return search_query[:75], price, asin
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        if retry:
            print("再試行します。")
            return get_amazon_product_info(url, retry=False)  # 再帰呼び出し時にはretryをFalseに設定
        else:
            # 再試行後のエラーの場合はNoneを返す
            return None, None, None




def get_nodes_by_class(search_url, class_names, boot_driver=False):
    # Chrome WebDriverのインスタンスを作成
    if boot_driver:
        driver = create_chrome_driver()  # ChromeDriverのパスが必要になる場合があります
        driver.get(search_url)
        time.sleep(2)
        for i in range(3):
            driver.execute_script(f"window.scrollBy(0, 30000);")
            time.sleep(1)
        # JavaScriptが実行された後のページのHTMLを取得
        html = driver.page_source
        soup = BeautifulSoup(html, "html.parser")
        elements = soup.find_all(class_=class_names)
        #driver.quit()  # ブラウザを閉じる
    else:
        response = requests.get(search_url)

        # HTML を解析
        soup = BeautifulSoup(response.content, "html.parser")
        elements = soup.find_all(class_=class_names) 

    return elements

def modify_url(original_url, additional_path):
    parsed_url = urlparse(original_url)
    # ドメイン後の最初のパスセグメントを取得
    path_segments = parsed_url.path.split('/')
    if len(path_segments) > 1:
        first_segment = path_segments[1]  # 最初のパスセグメント
    else:
        first_segment = ''

    # 新しいベースURLを構築
    base_url = f"{parsed_url.scheme}://{parsed_url.netloc}/{first_segment}"

    # urljoinを使用して新しいURLを構築
    new_url = urljoin(base_url + '/', additional_path)

    return new_url

def extract_numbers_from_strings(strings):
    numbers = []
    for string in strings:
        # 正規表現で数字部分を抜き出す
        matches = re.findall(r'\d{1,3}(?:,\d{3})*', string)
        for match in matches:
            # カンマを除去してリストに追加
            numbers.append(match)
    return numbers


def extract_number_from_string(string):
    # 正規表現で数字の部分を抜き出す
    match = re.search(r'\d+', string)
    if match:
        # 抜き出された数字を整数として返す
        return int(match.group())
    else:
        # 数字が見つからない場合はNoneを返す
        return ""


def sanitize_filename(filename):
    # WindowsとUNIX/Linuxで禁止されている文字を置換
    return re.sub(r'[<>:"/\\|?*]', '_', filename)

def save_data_to_csv(base_path, csv_data, filename_prefix):
    # ファイル名のサニタイズ
    sanitized_prefix = sanitize_filename(filename_prefix)
    
    current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{sanitized_prefix}_{current_time}.csv"
    full_path = os.path.join(base_path, filename)

    with open(full_path, mode='w', newline='', encoding='cp932', errors='ignore') as file:
        writer = csv.writer(file)
        for row in csv_data:
            writer.writerow(row)

    log_message(f"ファイル'{full_path}'にデータを保存しました。", "info")

def get_jan_from_asin(asin, timeout=10):
    """
    指定されたASINに基づいてJANコードを取得する関数。

    :param asin: Amazonの商品識別番号（ASIN）
    :param timeout: リクエストのタイムアウト秒数（デフォルトは10秒）
    :return: 対応するJANコード、またはエラーメッセージ
    """
    url = f"https://caju.jp/{asin}"
    log_message(f"ASIN '{asin}' に基づいて URL '{url}' にアクセスしています...")

    try:
        response = requests.get(url, timeout=timeout)
        #print(f"リクエストのステータスコード: {response.status_code}")

        if response.status_code != 200:
            return f"Error: Unable to access {url}"
    except requests.Timeout:
        return "Error: Request timed out"
    except requests.RequestException as e:
        return f"Error: An error occurred while making the request - {e}"

    #print(f"ASIN '{asin}' のページの内容を解析しています...")
    soup = BeautifulSoup(response.text, 'html.parser')
    jan_elements = soup.find_all(class_="ml-12")

    if len(jan_elements) > 1 and jan_elements[1].get_text().strip():
        jan_code = jan_elements[1].get_text().strip()
        log_message(f"ASIN '{asin}' に対する JANコード '{jan_code}' を見つけました。")
        return jan_code
    else:
        log_message(f"ASIN '{asin}' に対する JANコードが見つかりませんでした。")
        return ""

def get_texts_by_data_track_type(driver, data_track_type, timeout=60, scroll_amount=300):
    start_time = time.time()
    while True:
        if time.time() - start_time > timeout:
            raise Exception(f"Timeout on waiting for elements with data-track-type='{data_track_type}'.")
        try:
            # CSSセレクタを使用して、data-track-type属性が指定された値を持つ全ての要素を取得
            elements = driver.find_elements(By.CSS_SELECTOR, f'[data-track-type="{data_track_type}"]')
            if elements:
                # 要素が見つかった場合は、それらのテキストをリストにして返す
                return [element for element in elements]
            else:
                # 要素がまだ見つからない場合は、ページをスクロールして再検索
                driver.execute_script(f"window.scrollBy(0, {scroll_amount});")
                continue
        except Exception as e:
            # 何らかのエラーが発生した場合は、ページをスクロールして再検索
            print(f"An error occurred: {e}")
            driver.execute_script(f"window.scrollBy(0, {scroll_amount});")
            continue




















def scraping_rakuten(url, store_count):
    log_message("★★★★★★★★★★★★★★★★★★★★★")
    log_message("楽天のデータを収集します。")
    log_message("★★★★★★★★★★★★★★★★★★★★★")

    title, price, asin = get_amazon_product_info(url)

    if title is None or price is None or asin is None:
        log_message(f"URL {url} の情報取得に失敗しました。", "error")
        return

    jan = get_jan_from_asin(asin, timeout=10)

    if jan == "":
        jan = title

    log_message(f"AmazonURL:\n {url}\n\nAmazonタイトル:\n {title}\n\nAmazon価格:\n {price}\n\nasin:\n {asin}\n\njan:\n {jan}")

    driver = create_chrome_driver()  # ChromeDriverのパスが必要になる場合があります
    driver.get(f"https://search.rakuten.co.jp/search/mall/{jan}/?s=11")
    time.sleep(2)
    log_message(f"楽天URL：\nhttps://search.rakuten.co.jp/search/mall/{jan}/?s=11")



    csv_data = [["商品名", "Amazon価格", "楽天価格", "送料", "価格差", "利益率", "店名", "URL（売れている順）", "商品数", "評価", "レビュー数"]]

    
    stores_value = []
    while store_count >= len(stores_value):
        try:
            elements = get_texts_by_data_track_type(driver, "item", timeout=10, scroll_amount=0)
        except:
            #csv_data.append(["なし", "なし", "なし", "なし", "なし", "なし", "なし", "なし", "なし", "なし", "なし"])
            log_message("楽天の検索条件に一致する商品が見つかりませんでした。")
            return
        
        for element in elements:
            stores_value.append([element.text.split("\n"), element.get_attribute("data-shop-id")])
            #print([element.text.split("\n"), element.get_attribute("data-shop-id")])
            #store_value = elements[i].text.split("\n")
            #data_shop_id = elements[i].get_attribute("data-shop-id")
        try:
            click_element_with_exact_text(driver, "次のページ", timeout=5, scroll_amount=100)
        except:
            break

    for i in range(len(stores_value)):

        if i == store_count:  # インデックスは0から始まるので、30回目はインデックスが29
            break

        log_message("--------------------------")

        if store_count > len(stores_value):
            log_message("{}/{}個目".format(i + 1, len(stores_value)))
        else:
            log_message("{}/{}個目".format(i + 1, store_count))
        
        try:
            store_value = stores_value[i][0]
            data_shop_id = stores_value[i][1]

            
            url = f"https://review.rakuten.co.jp/rd/0_{data_shop_id}_{data_shop_id}_0/"

            # requestsを使用してWebページを取得
            response = requests.get(url)
            #print(url)
            #response.text
            # BeautifulSoupオブジェクトの作成（HTMLパーサーとしてlxmlを使用）
            soup = BeautifulSoup(response.text, 'lxml')

            assessment_elements = soup.select('.revEvaNumber.average')
            if assessment_elements:  # 空でない場合、要素が存在する
                assessment = assessment_elements[0].text
            else:
                assessment = ""

            subject_elements = soup.select('.count')
            if subject_elements:  # 空でない場合、要素が存在する
                subject = subject_elements[0].text
            else:
                subject = ""
            
            #data_shop_id = "270689"
            products_url = f"https://search.rakuten.co.jp/search/mall/?sid={data_shop_id}"
            response = requests.get(products_url)
            soup = BeautifulSoup(response.text, 'lxml')
            # "count"と"_medium"の両方のクラスを持つ<span>要素を検索
            element = soup.find('span', class_=["count", "_medium"])

            # 要素のテキストを表示
            if element:
                item_num = element.text.strip()
                # 正規表現を使用してカンマを含む数値を抜き出す
                # 括弧内の数字（カンマ含む）と「件」を抜き出す正規表現
                match = re.search(r'\（([0-9,]+)件\）', item_num)
                if match:
                    # 括弧内の数字のみを取得（「件」を除外）
                    item_num = match.group(1)  # '2,078,980' を取得
                    #print(item_num)

            else:
                item_num = ""
            
            log_message(store_value)
            log_message(products_url)
            log_message(item_num)
            log_message(url)
            log_message(assessment)
            log_message(subject)
            
            if "(価格+送料)" in store_value[1]:# or "中古" in store_value[1]:
                store_price = store_value[2].replace("円", "").replace("(価格+送料)", "").replace("〜", "")
                store_send_price = store_value[3].replace("+送料", "").replace("円", "")
            else:
                store_price = store_value[1].replace("円", "").replace("〜", "")
                store_send_price = "0"

            if int(store_price.replace(",", "")) >= int(price.replace(",", "")):
                col1 = store_value[0]
                if "同じ商品を安い順で見る" in store_value[-1]:
                    col2 = store_value[-2]
                else:
                    col2 = store_value[-1]
                
                col3 = products_url
                col4 = item_num
                col5 = assessment
                col6 = subject
                col7 = "¥" + price
                col8 = "¥" + store_price
                col9 = "¥" + store_send_price
                
                col10 = int(store_price.replace(",", "")) - int(price.replace(",", "")) + int(col9.replace("¥", "").replace(",", ""))
                # 3桁ごとにカンマを挿入してフォーマット
                col10 = format(col10, ",")

                # 計算式を実行
                col11 = ((int(store_price.replace(",", "")) + int(store_send_price.replace(",", ""))) / int(price.replace(",", ""))) - 1

                # 計算結果をパーセンテージ表示に変換
                col11 = col11 * 100

                # 計算結果を四捨五入して整数に変換
                col11 = round(col11)

                # パーセンテージを整数として文字列に変換し、「%」を追加
                col11 = f"{col11}%"
                #print(col11)

                csv_data.append([col1, col7, col8, col9, col10, col11, col2, col3, col4, col5, col6])
            
            else:
                log_message("Amazonより値段が安いためデータは追加しません。", "info")

        except Exception as e:
            csv_data.append(["エラー", "エラー", "エラー", "エラー", "エラー", "エラー", "エラー", "エラー", "エラー", "エラー", "エラー"])
            log_message("予期せぬエラーが発生しました。")
            log_message(f"エラーメッセージ: {e}")
            # スタックトレースを出力
            traceback.print_exc()

    return title, csv_data



def scraping_yahoo(url, store_count):
    log_message("★★★★★★★★★★★★★★★★★★★★★")
    log_message("ヤフーショッピングのデータを収集します。")
    log_message("★★★★★★★★★★★★★★★★★★★★★")
    #url = input("Amazon商品ページのURLをペーストしてエンターを押してください: ")
    #url = input()

    #url = "https://www.amazon.co.jp/%E7%8E%84%E4%BA%BA%E5%BF%97%E5%90%91-%E3%82%B7%E3%83%B3%E3%82%B0%E3%83%AB%E3%83%95%E3%82%A1%E3%83%B3%E3%83%A2%E3%83%87%E3%83%AB-GF-GT1030-E2GB-LP-D5/dp/B07Q6X71JD/ref=sr_1_1?__mk_ja_JP=%E3%82%AB%E3%82%BF%E3%82%AB%E3%83%8A&crid=15FY2BD5DUDXA&dib=eyJ2IjoiMSJ9.hEDqcUBhiXI7gwUEQgak1jDUgY5o_4gsHT1CKvJLzoQnMa9ZWfEq4QCOwUMMT0_E7Tw_Cbpa7RdmoLZdobUR5XYyQkCA_0Mutto-Ac8KTX1jlQDr6jRBywEG7ikum2D7NPfJH3Jv2FcYRJxjfOwCGnweL59jEJWsZXv_JOvKQpalAgEMGRF2QIJR42BzAXnOgm6QElIFhVwRQEOlkIp6-VyukGqKbwJO1EatR2Or1pEOpSSFdOGbrPYhnN6-V5r_lLD_u2oIaSmD8h8Ee8TI72SP7KQ1pR58FK9sS6HxtWg.PX0S2ExBVAJRGaOHK58bNtgq-R2WP7IONUxfCx9j_-0&dib_tag=se&keywords=%E7%8E%84%E4%BA%BA%E5%BF%97%E5%90%91%2BNVIDIA%2BGeForce%2BGT%2B1030%2B%E6%90%AD%E8%BC%89%2B%E3%82%B0%E3%83%A9%E3%83%95%E3%82%A3%E3%83%83%E3%82%AF%E3%83%9C%E3%83%BC%E3%83%89%2B2GB%2B%E3%82%B7%E3%83%B3%E3%82%B0%E3%83%AB%E3%83%95%E3%82%A1%E3%83%B3%E3%83%A2%E3%83%87%E3%83%AB%2BGF-GT1030-E2GB%2FLP%2FD5&qid=1708770298&s=computers&sprefix=%E7%8E%84%E4%BA%BA%E5%BF%97%E5%90%91%2Bnvidia%2Bgeforce%2Bgt%2B1030%2B%E6%90%AD%E8%BC%89%2B%E3%82%B0%E3%83%A9%E3%83%95%E3%82%A3%E3%83%83%E3%82%AF%E3%83%9C%E3%83%BC%E3%83%89%2B2gb%2B%E3%82%B7%E3%83%B3%E3%82%B0%E3%83%AB%E3%83%95%E3%82%A1%E3%83%B3%E3%83%A2%E3%83%87%E3%83%AB%2Bgf-gt1030-e2gb%2Flp%2Fd5%2Ccomputers%2C280&sr=1-1&th=1"

    title, price, asin = get_amazon_product_info(url)

    if title is None or price is None or asin is None:
        log_message(f"URL {url} の情報取得に失敗しました。", "error")
        return

    jan = get_jan_from_asin(asin, timeout=10)

    log_message(f"AmazonURL:\n {url}\n\nAmazonタイトル:\n {title}\n\nAmazon価格:\n {price}\n\nasin:\n {asin}\n\njan:\n {jan}")

    # 検索クエリ
    #search_query = "Anker 735 Charger (GaNPrime 65W) (USB PD 充電器A USB-A & USB-C 3ポート) (ブラック)"

    #search_query = "玄人志向 NVIDIA GeForce GT 1030 搭載 グラフィックボード 2GB シングルファンモデル GF-GT1030-E2GB/LP/D5"

    # URLエンコード
    encoded_query = quote(title)
    # 検索URLを構築
    base_url = "https://shopping.yahoo.co.jp/search?first=1&tab_ex=commerce&fr=shp-prop&mcr=d84c87ea8a3b637c3265e6fa772d12d2&ts=1708676353&sretry=1"
    search_url = f"{base_url}&p={encoded_query}&tab_ex=commerce&prom=1&X=2&sc_i=shopping-pc-web-result-item-sort_mdl-sortitem"
    log_message(f"ヤフーショッピングURL:\n{search_url}")
    #search_url = "https://shopping.yahoo.co.jp/search?X=2&p=CORSAIR+Elite+CPU+%E3%82%AF%E3%83%BC%E3%83%A9%E3%83%BC+%E3%82%A2%E3%83%83%E3%83%97%E3%82%B0%E3%83%AC%E3%83%BC%E3%83%89%E5%B0%82%E7%94%A8+LCD+%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%AD%E3%83%83%E3%83%88+CW-9060056-WW&first=1&ss_first=1&tab_ex=commerce&sc_i=shopping-pc-web-result-item-h_srch-kwd&ts=1708733442&mcr=bd273c123c0b7357bb7e01b5c0135157&sretry=1&area=13&b=31&view=list"
    elements = get_nodes_by_class(search_url, "LoopList__item", True)
    #log_message(len(elements))
    if not elements:
        log_message("Yahooショッピングの検索条件に一致する商品が見つかりませんでした。")
        return


    csv_data = [["商品名", "Amazon価格", "ヤフー価格", "送料", "価格差", "利益率", "店名", "URL（売れている順）", "商品数", "評価", "レビュー数"]]
    # 各要素のテキストを出力
    for i, element in enumerate(elements):
        # 現在の繰り返し回数が30回に達したらループを抜ける
        if i == store_count:  # インデックスは0から始まるので、30回目はインデックスが29
            break        
        try:
            log_message("--------------------------")

            if store_count > len(elements):
                log_message("{}/{}個目".format(i + 1, len(elements)))
            else:
                log_message("{}/{}個目".format(i + 1, store_count))
                
            span_texts = [span.text.strip() for span in element.find_all('span') if span.text.strip() != '']
            span_texts = [item for item in span_texts if "お取り寄せ" not in item]
            log_message(span_texts)
            
            if re.match(r'.*円分$', span_texts[0]):
                del span_texts[0]  # 0番目が「〇〇円分」である場合、削除
            
            a_tag = element.find('a')['href']
            #log_message(a_tag)
            yahoo_element = get_nodes_by_class(a_tag, "elInfoMain")
            yahoo_element_text = [span.text.strip() for span in yahoo_element[0].find_all('span') if span.text.strip() != '']
            while len(yahoo_element_text) < 3:
                yahoo_element_text.append("")
            log_message(yahoo_element_text)
            
            new_url = modify_url(a_tag, "search.html?X=4#CentSrchFilter1")
            log_message(new_url)
            
            yahoo_product_element = get_nodes_by_class(new_url, "mdSearchHeader")
            yahoo_product_element_text = [span.text.strip() for span in yahoo_product_element[0].find_all('p') if span.text.strip() != '']
            log_message(yahoo_product_element_text)
            
            # 「円」を含む要素のインデックスを探し、その一つ前の要素を取得
            for i, text in enumerate(span_texts):
                if "円" == text:
                    if i > 0:  # 最初の要素でなければ、一つ前が存在する
                        previous_value = span_texts[i-1]
                        #log_message(previous_value)
                        break
                    else:
                        log_message("予期せぬエラーが発生しました。", "error")
                        csv_data.append(["エラー", "エラー", "エラー", "エラー", "エラー", "エラー", "エラー", "エラー", "エラー", "エラー", "エラー"])
                        continue
            
            if int(previous_value.replace(",", "")) >= int(price.replace(",", "")):
                col1 = span_texts[0]
                col2 = yahoo_element_text[0]
                col3 = new_url
                col4 = extract_numbers_from_strings(yahoo_product_element_text)[0]
                col5 = yahoo_element_text[1]
                col6 = extract_number_from_string(yahoo_element_text[2])
                col7 = "¥" + price
                #log_message("col7")
                #log_message(col7)
                col8 = "¥" + previous_value
                #log_message("col8")
                #log_message(col8)

                for phrase in span_texts:
                    # 「+送料」が含まれているかどうか確認し、その後の数字部分を抽出（カンマも含む）
                    match = re.search(r'\+送料(\d{1,3}(,\d{3})*)', phrase)
                    if match:
                        # 数字が見つかった場合、カンマを削除して出力
                        number = match.group(1)#.replace(',', '')  # group(1)で最初のキャプチャグループ（数字部分）を取得
                        col9 = "¥" + number
                        #log_message("col9")
                        #log_message(col9)
                        break
                    else:
                        number = "0"
                        col9 = f"¥{number}"
                        #log_message("col9")
                        #log_message(col9)
                
                col10 = int(previous_value.replace(",", "")) - int(price.replace(",", "")) + int(number.replace(",", ""))
                # 3桁ごとにカンマを挿入してフォーマット
                col10 = format(col10, ",")

                # 計算式を実行
                col11 = ((int(previous_value.replace(",", "")) + int(number.replace(",", ""))) / int(price.replace(",", ""))) - 1

                # 計算結果をパーセンテージ表示に変換
                col11 = col11 * 100

                # 計算結果を四捨五入して整数に変換
                col11 = round(col11)

                # パーセンテージを整数として文字列に変換し、「%」を追加
                col11 = f"{col11}%"
                #log_message(col11)

                csv_data.append([col1, col7, col8, col9, col10, col11, col2, col3, col4, col5, col6])
            else:
                log_message("Amazonより値段が安いためデータは追加しません。", "info")
        except Exception as e:
            csv_data.append(["エラー", "エラー", "エラー", "エラー", "エラー", "エラー", "エラー", "エラー", "エラー", "エラー", "エラー"])
            log_message("予期せぬエラーが発生しました。")
            log_message(f"エラーメッセージ: {e}")
            # スタックトレースを出力
            traceback.print_exc()

    return title, csv_data


# ------------------------------------------------------

def start_thread():
    # スレッドを作成し、targetにmain関数を指定
    thread = threading.Thread(target=main)
    # スレッドを開始
    thread.start()

def main():
    input_datas = validate_inputs()
    start_button.configure(state=ctk.DISABLED)

    if input_datas:
        log_message("-------------------------------------------")
        if "amazon_urls" in input_datas:
            for url in input_datas["amazon_urls"]:
                log_message(f"Amazon URL:\n{url}")
        log_message(f"Yahoo選択: {input_datas['yahoo_selected']}")
        log_message(f"楽天選択: {input_datas['rakuten_selected']}")
        log_message(f"ストア件数: {input_datas['store_count']}")
        log_message(f"CSV保存先: {input_datas['csv_save_path']}")
        log_message("-------------------------------------------")
    else:
        start_button.configure(state=ctk.NORMAL)
        return
    
    if input_datas["yahoo_selected"]:
        for i, amazon_url in enumerate(input_datas["amazon_urls"]):
            # scraping_yahoo関数から戻り値を受け取る
            result = scraping_yahoo(amazon_url, input_datas["store_count"])
            
            # resultがNoneかどうかをチェック
            if result is not None:
                title, csv_data = result
                # 結果がNoneでなければ、処理を続ける
                base_path = input_datas["csv_save_path"]  # 保存先のベースディレクトリのパスを指定
                save_data_to_csv(base_path, csv_data, f"【ヤ{i+1}】{title}")
            else:
                # 結果がNoneの場合、このURLに対する処理をスキップ
                log_message(f"{amazon_url} に対するデータ取得に失敗しました。", "error")

    if input_datas["rakuten_selected"]:
        for i, amazon_url in enumerate(input_datas["amazon_urls"]):
            # scraping_yahoo関数から戻り値を受け取る
            result = scraping_rakuten(amazon_url, input_datas["store_count"])
            
            # resultがNoneかどうかをチェック
            if result is not None:
                title, csv_data = result
                # 結果がNoneでなければ、処理を続ける
                base_path = input_datas["csv_save_path"]  # 保存先のベースディレクトリのパスを指定
                save_data_to_csv(base_path, csv_data, f"【楽{i+1}】{title}")
            else:
                # 結果がNoneの場合、このURLに対する処理をスキップ
                log_message(f"{amazon_url} に対するデータ取得に失敗しました。", "error")


    start_button.configure(state=ctk.NORMAL)
    driver.quit()
    log_message("処理が終了しました。")


def validate_inputs():
    # Amazon URLのバリデーション
    amazon_urls = amazon_url_box.get("1.0", tk.END).strip()
    if not amazon_urls:
        messagebox.showerror("バリデーションエラー", "AmazonのURLを入力してください。")
        return
    
    # URLの形式を判定する正規表現パターン
    url_pattern = re.compile(r'https?://[^\s]+')

    for url in amazon_urls.split("\n"):
        if not url_pattern.match(url):
            messagebox.showerror("バリデーションエラー", "無効なURLが含まれています。")
            return

    # チェックボックスのバリデーション
    if not (yahoo_check_var.get() or rakuten_check_var.get()):
        messagebox.showerror("バリデーションエラー", "少なくとも一つのサイトを選択してください。")
        return


    # 取得ストア件数のバリデーション
    try:
        store_count = int(store_count_entry.get())
    except ValueError:
        messagebox.showerror("バリデーションエラー", "取得ストア件数には数値を入力してください。")
        return


    # CSV保存先のバリデーション
    csv_save_path = csv_save_entry.get()
    if not csv_save_path or not os.path.exists(csv_save_path):
        messagebox.showerror("バリデーションエラー", "有効なCSV保存先パスを入力してください。")
        return


    # すべてのバリデーションを通過した場合、値を返す
    return {
        "yahoo_selected": yahoo_check_var.get(),
        "rakuten_selected": rakuten_check_var.get(),
        "amazon_urls": amazon_urls.split("\n"),
        "store_count": store_count,
        "csv_save_path": csv_save_path
    }


# ------------------------------------------------------
    
ctk.set_appearance_mode("Dark")  # 'Dark' または 'Light'
ctk.set_default_color_theme("blue")  # 'blue' (デフォルト), 'dark-blue', 'green'

root = ctk.CTk()
root.title("ライバルセラー抽出ツール_ver2.6")
root.geometry("970x500")

#root.grid_rowconfigure(1, weight=1)  # 1行目の高さを固定
#root.grid_columnconfigure(1, weight=1)  # 1列目の幅を固定

# 太文字フォントを適用する
bold_font = ctk.CTkFont(family="Arial", size=13)
bold_font2 = ctk.CTkFont(family="Arial", size=14)

# 説明テキストを表示するラベルを作成
description_label = ctk.CTkLabel(root,
                                  text="Amazonと同じ商品を出品しているストアを抽出するツールです。\n以下の項目を入力して取得実行を押下すると、商品ごとにcsvが自動で保存されます。",
                                  wraplength=680,  # ラベルのテキストの折り返し幅を指定
                                  font=bold_font,  # 既に定義した太文字フォントを使用
                                  anchor='e', 
                                  justify='left')  # テキストを左寄せにする
description_label.grid(row=0, column=0, columnspan=2, pady=10, padx=10, sticky="w")

# Amazonの商品ページのURLに関するラベルの作成と配置
#amazon_url_label = ctk.CTkLabel(root, text="Amazonの商品ページのURL", font=bold_font, anchor="w")
#amazon_url_label.grid(row=1, column=0, pady=(10, 5), padx=10, sticky="w")

# テキストボックスを含むフレームの作成
frame1 = ctk.CTkFrame(root)
frame1.grid(row=2, column=0, pady=(10, 5), padx=10, sticky="nsew")
#frame1.grid_propagate(False)  # フレームのサイズを固定

# Amazonの商品ページのURLに関するラベルの作成とフレーム1内への配置
amazon_url_label = ctk.CTkLabel(frame1, text="Amazonの商品ページのURL：", font=bold_font2)
amazon_url_label.grid(row=0, column=0, columnspan=3, padx=10, pady=5, sticky="w")

# テキストボックスの作成と配置
amazon_url_box = ctk.CTkTextbox(frame1, width=250, height=150)
amazon_url_box.grid(row=1, column=0, columnspan=4, padx=10, pady=(5, 20), sticky="nsew")
#amazon_url_box.grid_propagate(False)  

# 「取得サイト選択」のラベルとチェックボックスを含むフレーム内のウィジェットを横に並べる
select_site_label = ctk.CTkLabel(frame1, text="取得サイト選択：", font=bold_font2)
select_site_label.grid(row=2, column=0, padx=10, pady=10, sticky="w")

yahoo_check_var = ctk.BooleanVar(value=True)  # チェックボックスの状態を保持する変数、デフォルトでチェック
yahoo_checkbox = ctk.CTkCheckBox(frame1, text="ヤフーショッピング", variable=yahoo_check_var)
yahoo_checkbox.grid(row=2, column=1, padx=(10, 5), pady=10, sticky="w")

rakuten_check_var = ctk.BooleanVar(value=True)  # チェックボックスの状態を保持する変数、デフォルトでチェック
rakuten_checkbox = ctk.CTkCheckBox(frame1, text="楽天市場", variable=rakuten_check_var)
rakuten_checkbox.grid(row=2, column=2, padx=10, pady=10, sticky="w")


# 取得ストア件数のラベルとエントリ
store_count_label = ctk.CTkLabel(frame1, text="取得ストア件数：", font=bold_font2)
store_count_label.grid(row=3, column=0, padx=10, pady=(10, 5), sticky="w")

# エントリにデフォルト値を設定
store_count_entry = ctk.CTkEntry(frame1, width=80, placeholder_text="例: 30")
store_count_entry.grid(row=3, column=1, padx=(10, 5), pady=(10, 5), sticky="w")
store_count_entry.insert(0, "30")  # デフォルトのテキストを挿入

#store_count_entry = ctk.CTkEntry(frame1, width=80, placeholder_text="例: 30")
# 右のパディングを減らす
#store_count_entry.grid(row=3, column=1, padx=(10, 0), pady=(10, 5), sticky="w")

# 「件」というラベル
store_subject_label = ctk.CTkLabel(frame1, text="件", font=bold_font2)
# 左のパディングを減らして「件」をエントリに近づける
#store_subject_label.grid(row=3, column=2, padx=(0, 10), pady=(5, 5), sticky="w")

# 「件」のラベルを特定の座標に配置
store_subject_label.place(x=238, y=270)

# CSV保存先のラベル
csv_save_label = ctk.CTkLabel(frame1, text="CSV保存先：", font=bold_font2)
csv_save_label.grid(row=4, column=0, padx=10, pady=(10, 5), sticky="w")

# CSV保存先を入力するためのエントリ
csv_save_entry = ctk.CTkEntry(frame1, width=120)
csv_save_entry.grid(row=4, column=1, columnspan=2, padx=(10, 5), pady=(10, 5), sticky="ew")

# 参照ボタンの作成と配置
def browse_button_command():
    # ファイル選択ダイアログを開き、選択されたフォルダのパスをエントリに設定
    folder_selected = filedialog.askdirectory()
    csv_save_entry.delete(0, ctk.END)
    csv_save_entry.insert(0, folder_selected)

browse_button = ctk.CTkButton(frame1, text="参照", font=bold_font2, width=20, height=25, command=browse_button_command, 
                                fg_color="#A9A9A9",  # ボタンの色 (ダークグレー)
                                hover_color="#D3D3D3",  # マウスホバー時の色 (ライトグレー)
                                text_color="#000000"  # 文字色 (黒)
                               )
browse_button.grid(row=4, column=3, padx=10, pady=(10, 5), sticky="ew")

# スタートボタンの作成と配置
start_button = ctk.CTkButton(frame1, text="スタート", font=bold_font2, command=start_thread)
start_button.grid(row=5, column=0, columnspan=4, padx=10, pady=(20, 20), sticky="ew")



# ログ表示用のフレームの作成
frame2 = ctk.CTkFrame(root)
frame2.grid(row=2, column=1, pady=(10, 5), padx=10, sticky="nsew")
#frame2.grid_propagate(False)  # フレームのサイズを固定

# ログ表示用のラベル
log_label = ctk.CTkLabel(frame2, text="ログ：", font=bold_font2)
log_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")

# ログ表示用のテキストボックス
log_text_box = ctk.CTkTextbox(frame2, width=450, height=300)
log_text_box.grid(row=1, column=0, padx=10, pady=(5, 19), sticky="nsew")
# タグの設定は一度だけ行う
log_text_box.tag_config("error", foreground="red")
log_text_box.tag_config("info", foreground="#32CD32")  # 鮮やかな緑色

def save_log():
    # 現在の日付と時刻を取得してファイル名に使用
    now = datetime.now()
    formatted_date = now.strftime("%Y%m%d_%H%M%S")  # YYYYMMDD_HHMMSSの形式
    default_filename = f"【ログ】ライバルセラー抽出ツール_{formatted_date}.txt"  # デフォルトのファイル名

    # ファイル保存ダイアログを開く
    filepath = filedialog.asksaveasfilename(
        initialfile=default_filename,  # デフォルトファイル名を設定
        defaultextension=".txt",
        filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
    )
    if filepath:  # ファイルパスが選択された場合
        with open(filepath, "w", encoding="utf-8") as file:
            file.write(log_text_box.get("1.0", "end"))  # テキストボックスの内容をファイルに書き込む

log_save_button = ctk.CTkButton(frame2, text="ログ保存", font=bold_font2, command=save_log, 
                                fg_color="#A9A9A9",  # ボタンの色 (ダークグレー)
                                hover_color="#D3D3D3",  # マウスホバー時の色 (ライトグレー)
                                text_color="#000000"  # 文字色 (黒)
                               )
log_save_button.grid(row=3, column=0, padx=10, pady=(5, 10), sticky="ew")


def log_message(message, message_type="default"):
    # メッセージをテキストボックスに挿入
    log_text_box.insert(tk.END, str(message) + "\n\n")
    
    if message_type == "error":
        # エラーメッセージの場合、赤色にするためのタグを追加
        log_text_box.tag_add("error", "end-4c linestart", "end-1c")
    elif message_type == "info":
        # 情報メッセージの場合、緑色にするためのタグを追加
        log_text_box.tag_add("info", "end-4c linestart", "end-1c")
    
    # スクロールしてメッセージを表示
    log_text_box.see(tk.END)


# アプリケーションの起動時に日時ログを出力
current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
log_message(f"アプリケーションが起動しました。\n時刻: {current_time}\n")



# アプリケーションの実行
root.mainloop()

try:
    driver.quit()
except:
    pass

def kill_process_by_pid(pid):
    try:
        process = psutil.Process(pid)
        process.terminate()  # もしくは process.kill() を使用することもできます
    except psutil.NoSuchProcess:
        pass
kill_process_by_pid(os.getpid())