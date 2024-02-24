
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


# ChromeのWebDriverのインスタンスを生成します。
# もしエラーが発生した場合、"chromedriver"プロセスを終了して再度試みます。
def create_chrome_driver():
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
        print(f"エラーが発生しました: {e}")
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


def get_amazon_product_info(url):
    # WebDriverの設定と初期化
    driver = create_chrome_driver()  # 適切なパスを設定する必要があるかもしれません
    driver.get(url)

    # 商品のタイトルと価格を取得
    search_query = get_text_by_id(driver, "productTitle", timeout=10, scroll_amount=0)
    price = get_text_by_class(driver, "a-price-whole", timeout=10, scroll_amount=0)

    # ブラウザを閉じる
    driver.quit()

    return search_query[:75], price


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
        driver.quit()  # ブラウザを閉じる
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

    print(f"ファイル'{full_path}'にデータを保存しました。")



while True:
    print(" ")
    print("★★★★★★★★★★★★★★★★★★★★★")
    url = input("Amazon商品ページのURLをペーストしてエンターを押してください: ")
    #url = input()

    #url = "https://www.amazon.co.jp/%E7%8E%84%E4%BA%BA%E5%BF%97%E5%90%91-%E3%82%B7%E3%83%B3%E3%82%B0%E3%83%AB%E3%83%95%E3%82%A1%E3%83%B3%E3%83%A2%E3%83%87%E3%83%AB-GF-GT1030-E2GB-LP-D5/dp/B07Q6X71JD/ref=sr_1_1?__mk_ja_JP=%E3%82%AB%E3%82%BF%E3%82%AB%E3%83%8A&crid=15FY2BD5DUDXA&dib=eyJ2IjoiMSJ9.hEDqcUBhiXI7gwUEQgak1jDUgY5o_4gsHT1CKvJLzoQnMa9ZWfEq4QCOwUMMT0_E7Tw_Cbpa7RdmoLZdobUR5XYyQkCA_0Mutto-Ac8KTX1jlQDr6jRBywEG7ikum2D7NPfJH3Jv2FcYRJxjfOwCGnweL59jEJWsZXv_JOvKQpalAgEMGRF2QIJR42BzAXnOgm6QElIFhVwRQEOlkIp6-VyukGqKbwJO1EatR2Or1pEOpSSFdOGbrPYhnN6-V5r_lLD_u2oIaSmD8h8Ee8TI72SP7KQ1pR58FK9sS6HxtWg.PX0S2ExBVAJRGaOHK58bNtgq-R2WP7IONUxfCx9j_-0&dib_tag=se&keywords=%E7%8E%84%E4%BA%BA%E5%BF%97%E5%90%91%2BNVIDIA%2BGeForce%2BGT%2B1030%2B%E6%90%AD%E8%BC%89%2B%E3%82%B0%E3%83%A9%E3%83%95%E3%82%A3%E3%83%83%E3%82%AF%E3%83%9C%E3%83%BC%E3%83%89%2B2GB%2B%E3%82%B7%E3%83%B3%E3%82%B0%E3%83%AB%E3%83%95%E3%82%A1%E3%83%B3%E3%83%A2%E3%83%87%E3%83%AB%2BGF-GT1030-E2GB%2FLP%2FD5&qid=1708770298&s=computers&sprefix=%E7%8E%84%E4%BA%BA%E5%BF%97%E5%90%91%2Bnvidia%2Bgeforce%2Bgt%2B1030%2B%E6%90%AD%E8%BC%89%2B%E3%82%B0%E3%83%A9%E3%83%95%E3%82%A3%E3%83%83%E3%82%AF%E3%83%9C%E3%83%BC%E3%83%89%2B2gb%2B%E3%82%B7%E3%83%B3%E3%82%B0%E3%83%AB%E3%83%95%E3%82%A1%E3%83%B3%E3%83%A2%E3%83%87%E3%83%AB%2Bgf-gt1030-e2gb%2Flp%2Fd5%2Ccomputers%2C280&sr=1-1&th=1"

    title, price = get_amazon_product_info(url)
    print(f"Title: {title}, Price: {price}")

    # 検索クエリ
    #search_query = "Anker 735 Charger (GaNPrime 65W) (USB PD 充電器A USB-A & USB-C 3ポート) (ブラック)"

    #search_query = "玄人志向 NVIDIA GeForce GT 1030 搭載 グラフィックボード 2GB シングルファンモデル GF-GT1030-E2GB/LP/D5"

    # URLエンコード
    encoded_query = quote(title)
    # 検索URLを構築
    base_url = "https://shopping.yahoo.co.jp/search?first=1&tab_ex=commerce&fr=shp-prop&mcr=d84c87ea8a3b637c3265e6fa772d12d2&ts=1708676353&sretry=1"
    search_url = f"{base_url}&p={encoded_query}&tab_ex=commerce&prom=1&X=2&sc_i=shopping-pc-web-result-item-sort_mdl-sortitem"
    print(search_url)
    #search_url = "https://shopping.yahoo.co.jp/search?X=2&p=CORSAIR+Elite+CPU+%E3%82%AF%E3%83%BC%E3%83%A9%E3%83%BC+%E3%82%A2%E3%83%83%E3%83%97%E3%82%B0%E3%83%AC%E3%83%BC%E3%83%89%E5%B0%82%E7%94%A8+LCD+%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%AD%E3%83%83%E3%83%88+CW-9060056-WW&first=1&ss_first=1&tab_ex=commerce&sc_i=shopping-pc-web-result-item-h_srch-kwd&ts=1708733442&mcr=bd273c123c0b7357bb7e01b5c0135157&sretry=1&area=13&b=31&view=list"
    elements = get_nodes_by_class(search_url, "LoopList__item", True)
    print(len(elements))
    if not elements:
        print("Yahooショッピングの検索条件に一致する商品が見つかりませんでした。")
        continue


    csv_data = [["商品名", "Amazon価格", "ヤフー価格", "送料", "価格差", "利益率", "店名", "URL（売れている順）", "商品数", "評価", "レビュー数"]]
    # 各要素のテキストを出力
    for i, element in enumerate(elements):
        # 現在の繰り返し回数が30回に達したらループを抜ける
        if i == 30:  # インデックスは0から始まるので、30回目はインデックスが29
            break        
        try:
            print("--------------------------")
            span_texts = [span.text.strip() for span in element.find_all('span') if span.text.strip() != '']
            span_texts = [item for item in span_texts if "お取り寄せ" not in item]
            print(span_texts)
            """
            if re.match(r'.*円分$', span_texts[0]):
                del span_texts[0]  # 0番目が「〇〇円分」である場合、削除
            """
            a_tag = element.find('a')['href']
            #print(a_tag)
            yahoo_element = get_nodes_by_class(a_tag, "elInfoMain")
            yahoo_element_text = [span.text.strip() for span in yahoo_element[0].find_all('span') if span.text.strip() != '']
            while len(yahoo_element_text) < 3:
                yahoo_element_text.append("")
            print(yahoo_element_text)
            
            new_url = modify_url(a_tag, "search.html?X=4#CentSrchFilter1")
            print(new_url)
            
            yahoo_product_element = get_nodes_by_class(new_url, "mdSearchHeader")
            yahoo_product_element_text = [span.text.strip() for span in yahoo_product_element[0].find_all('p') if span.text.strip() != '']
            print(yahoo_product_element_text)
            
            # 「円」を含む要素のインデックスを探し、その一つ前の要素を取得
            for i, text in enumerate(span_texts):
                if "円" == text:
                    if i > 0:  # 最初の要素でなければ、一つ前が存在する
                        previous_value = span_texts[i-1]
                        #print(previous_value)
                        break
                    else:
                        print("予期せぬエラーが発生しました。")
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
                #print("col7")
                #print(col7)
                col8 = "¥" + previous_value
                #print("col8")
                #print(col8)

                for phrase in span_texts:
                    # 「+送料」が含まれているかどうか確認し、その後の数字部分を抽出（カンマも含む）
                    match = re.search(r'\+送料(\d{1,3}(,\d{3})*)', phrase)
                    if match:
                        # 数字が見つかった場合、カンマを削除して出力
                        number = match.group(1)#.replace(',', '')  # group(1)で最初のキャプチャグループ（数字部分）を取得
                        col9 = "¥" + number
                        #print("col9")
                        #print(col9)
                        break
                    else:
                        number = "0"
                        col9 = f"¥{number}"
                        #print("col9")
                        #print(col9)
                
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
                #print(col11)

                csv_data.append([col1, col7, col8, col9, col10, col11, col2, col3, col4, col5, col6])
            else:
                print("Amazonより値段が安いためデータは追加しません。")
        except Exception as e:
            csv_data.append(["エラー", "エラー", "エラー", "エラー", "エラー", "エラー", "エラー", "エラー", "エラー", "エラー", "エラー"])
            print("予期せぬエラーが発生しました。")
            print(f"エラーメッセージ: {e}")
            # スタックトレースを出力
            traceback.print_exc()


    # 使用例
    base_path = get_desktop_path()  # 保存先のベースディレクトリのパスを指定
    save_data_to_csv(base_path, csv_data, title)

#time.sleep(20)


