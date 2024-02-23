
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
        driver = webdriver.Chrome(ChromeDriverManager().install())
        return driver
    except Exception as e:
        kill_process_by_name("chromedriver")
        return webdriver.Chrome(ChromeDriverManager().install())


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

    return search_query, price


def get_nodes_by_class(search_url, class_names):
    #class_names = "LoopList__item"
    # リクエストを送信
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



def save_data_to_csv(base_path, csv_data, filename_prefix):
    """
    指定されたパスにデータをCSVファイルとして保存します。
    """
    current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{filename_prefix}_{current_time}.csv"
    full_path = os.path.join(base_path, filename)

    with open(full_path, mode='w', newline='', encoding='cp932', errors='ignore') as file:
        writer = csv.writer(file)
        for row in csv_data:
            writer.writerow(row)

    print(f"ファイル'{full_path}'にデータを保存しました。")



while True:
    url = input("Amazon商品ページのURLをペーストしてエンターを押してください: ")
    #url = input()

    #url = "https://www.amazon.co.jp/Anker-Charger-GaNPrime-USB-USB-C/dp/B09W9M89WS/ref=sr_1_1?__mk_ja_JP=%E3%82%AB%E3%82%BF%E3%82%AB%E3%83%8A&crid=30U0BSTB27SXP&dib=eyJ2IjoiMSJ9.1bzi2pdOe3j3A1gNHbL1DAihuuQmjZIKXF9uY7qsh7b0jenbUeuJvZN4nyaZZAkrmrtKwIxfjN7XSxBzqO2idbZ2EBCgjASST_MO3HCdu5xZuFlgfsF6b7DIA_I5nHe3lcVJJ7AEsMobx1lBx68kCozZwUHeMsSMWJPpY0cNgzksbqGh9EVVrAyvrXz23YNTcbNBYRcjvdzC_YCto0Imau6RxPeOQh81Ik7-YxAlNU1LOJ0t9SiGA8e5rHSlWu1xbJP87OQXOUAqGFQymwDG0MsxE_kUl7ZS5rIOslEcdMk.sXrY8ho8apy3Q0pOjf3AcQHdeYU82hkeGRRMjh01GjY&dib_tag=se&keywords=Anker%2B735%2BCharger%2B(GaNPrime%2B65W)%2B(USB%2BPD%2B%E5%85%85%E9%9B%BB%E5%99%A8A%2BUSB-A%2B%26%2BUSB-C%2B3%E3%83%9D%E3%83%BC%E3%83%88)%2B(%E3%83%96%E3%83%A9%E3%83%83%E3%82%AF)&qid=1708731297&sprefix=anker%2B735%2Bcharger%2Bganprime%2B65w%2Busb%2Bpd%2B%E5%85%85%E9%9B%BB%E5%99%A8a%2Busb-a%2B%26%2Busb-c%2B3%E3%83%9D%E3%83%BC%E3%83%88%2B%E3%83%96%E3%83%A9%E3%83%83%E3%82%AF%2B%2Caps%2C289&sr=8-1&th=1"

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

    elements = get_nodes_by_class(search_url, "LoopList__item")

    if not elements:
        print("Yahooショッピングの検索条件に一致する商品が見つかりませんでした。")
        continue


    csv_data = [["商品名", "店名", "URL（売れている順）", "商品数", "評価", "レビュー数"]]
    # 各要素のテキストを出力
    for element in elements:
        print("--------------------------")
        span_texts = [span.text.strip() for span in element.find_all('span') if span.text.strip() != '']
        print(span_texts)
        a_tag = element.find('a')['href']
        #print(a_tag)
        yahoo_element = get_nodes_by_class(a_tag, "elInfoMain")
        yahoo_element_text = [span.text.strip() for span in yahoo_element[0].find_all('span') if span.text.strip() != '']
        print(yahoo_element_text)
        new_url = modify_url(a_tag, "search.html?X=4#CentSrchFilter1")
        print(new_url)
        yahoo_product_element = get_nodes_by_class(new_url, "mdSearchHeader")
        yahoo_product_element_text = [span.text.strip() for span in yahoo_product_element[0].find_all('p') if span.text.strip() != '']
        print(yahoo_product_element_text)
        if int(span_texts[1].replace(",", "")) >= int(price.replace(",", "")):
            col1 = span_texts[0]
            col2 = yahoo_element_text[0]
            col3 = new_url
            col4 = extract_numbers_from_strings(yahoo_product_element_text)[0]
            col5 = yahoo_element_text[1]
            col6 = extract_number_from_string(yahoo_element_text[2])

            csv_data.append([col1, col2, col3, col4, col5, col6])
        else:
            print("Amazonより値段が安いためデータは追加しません。")


    # 使用例
    base_path = get_desktop_path()  # 保存先のベースディレクトリのパスを指定
    save_data_to_csv(base_path, csv_data, title)

#time.sleep(20)


