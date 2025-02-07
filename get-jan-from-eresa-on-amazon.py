import os
import configparser
import requests
from bs4 import BeautifulSoup
from google.oauth2 import service_account
from googleapiclient.discovery import build
import re

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
from selenium.webdriver.common.keys import Keys

# 設定ファイルのパス
CONFIG_FILE = 'config.ini'


def load_config():
    """設定ファイルを読み込みます。"""
    config = configparser.ConfigParser()
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            config.read_file(f)
    except FileNotFoundError:
        print(f"エラー: 設定ファイル '{CONFIG_FILE}' が見つかりません。")
        return None
    except configparser.Error as e:
        print(f"エラー: 設定ファイルの読み込みに失敗しました: {e}")
        return None
    return config['DEFAULT']


def authenticate_sheets_api(credentials_file):
    """Google Sheets APIを認証します。"""
    print("認証処理を開始します")
    print(f"credのファイルパス: {credentials_file}を確認中")
    try:
        if os.path.exists(credentials_file):
            print("サービスアカウントキーファイルが存在します")
            creds = service_account.Credentials.from_service_account_file(
                credentials_file, scopes=['https://www.googleapis.com/auth/spreadsheets']
            )
            print("認証情報を読み込みました")

            service = build('sheets', 'v4', credentials=creds)
            print("APIクライアントを作成しました")
            return service
        else:
            print("サービスアカウントキーファイルが存在しません")
            return None

    except Exception as e:
        print(f"認証エラー: {type(e)}, {e}")
        return None


def get_ebay_links_from_spreadsheet(service, spreadsheet_id, sheet_name, ebay_link_column, start_row):
    """スプレッドシートから指定範囲のeBayリンクを取得します。"""
    try:
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=spreadsheet_id,
                                    range=f'{sheet_name}!{ebay_link_column}{start_row}:{ebay_link_column}').execute()
        ebay_links = result.get('values', [])
        return [link[0] for link in ebay_links] if ebay_links else []  # リスト内包表記でURLのみを抽出
    except Exception as e:
        print(f"スプレッドシートからのリンク取得エラー: {e}")
        return None


def get_ebay_image_url(ebay_url):
    """eBayの商品ページから画像URLを取得します。リダイレクトに対応し、指定されたクラスのdivタグから画像URLを検出し、active imageを優先します。"""
    try:
        try:
            response = requests.get(ebay_url, timeout=20)  # タイムアウトを設定
            response.raise_for_status()
        except requests.exceptions.RequestException as e:
            print(f"eBayからのリクエストエラー: {e}")
            return None

        try:
            soup = BeautifulSoup(response.content, 'html.parser')
        except Exception as e:
            print(f"BeautifulSoup解析エラー: {e}")
            return None

        # リダイレクトページを検出
        if "Redirecting you to" in soup.get_text():
            print("リダイレクトページを検出しました。")
            # リダイレクト先のURLを取得
            redirect_url = soup.find('meta', attrs={'http-equiv': 'refresh'})
            if redirect_url and 'url=' in redirect_url.get('content', ''):
                redirect_url_value = redirect_url.get('content').split('url=')[1]

                try:
                    print(f"リダイレクト先URLへ再度アクセスします: {redirect_url_value}")
                    response = requests.get(redirect_url_value, timeout=20)  # リダイレクト先へアクセス
                    response.raise_for_status()
                    soup = BeautifulSoup(response.content, 'html.parser')

                except requests.exceptions.RequestException as e:
                    print(f"リダイレクト先へのリクエストエラー: {e}")
                    return None
                except Exception as e:
                    print(f"リダイレクト後のBeautifulSoup解析エラー: {e}")
                    return None
            else:
                print("リダイレクトURLを抽出できませんでした")
                return None

        active_image_url = None
        image_url = None

        # image-carousel-item image-treatment active image の検索と属性の検出
        div_tag_active = soup.find('div', class_='ux-image-carousel-item image-treatment active image')
        if div_tag_active:
            img_tag = div_tag_active.find('img')
            if img_tag:
                if 'src' in img_tag.attrs:
                    active_image_url = img_tag['src']
                elif 'srcset' in img_tag.attrs:
                    srcset = img_tag['srcset']
                    active_image_url = srcset.split(',')[0].split(' ')[0]
                elif 'data-zoom-src' in img_tag.attrs:
                    active_image_url = img_tag['data-zoom-src']

        # image-carousel-item image-treatment image の検索と属性の検出
        div_tag = soup.find('div', class_='ux-image-carousel-item image-treatment image')
        if div_tag:
            img_tag = div_tag.find('img')
            if img_tag:
                if 'src' in img_tag.attrs:
                    image_url = img_tag['src']
                elif 'srcset' in img_tag.attrs:
                    srcset = img_tag['srcset']
                    image_url = srcset.split(',')[0].split(' ')[0]
                elif 'data-zoom-src' in img_tag.attrs:
                    image_url = img_tag['data-zoom-src']

        # URLの優先順位の決定
        if active_image_url:
            return active_image_url
        elif image_url:
            return image_url
        else:
            print("指定されたdivタグのいずれからも画像URLが見つかりませんでした。")
            return None

    except Exception as e:
        print(f"get_ebay_image_url関数全体でのエラー: {e}")
        return None


class ChromeBrowser:
    def __init__(self, crx_path=None, eresa_username=None, eresa_password=None):
        self.driver = None
        self.crx_path = crx_path
        self.logged_in_eresa = False # ERESAのログイン状態を追跡
        self.eresa_username = eresa_username
        self.eresa_password = eresa_password
        self.first_amazon_access = True  # 初回Amazonアクセスかどうかを追跡

    def initialize_driver(self):
        """ChromeDriverを初期化します。"""
        if not self.driver:
            print("ChromeDriverを起動します...")
            options = Options()
            # options.add_argument("--headless")  # ヘッドレスモードを有効にする

            if self.crx_path:
                print(f"拡張機能（CRXファイル）を追加します: {self.crx_path}")
                options.add_extension(self.crx_path)

            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=options)
            print("ChromeDriverを起動しました。")

    def find_first_amazon_url(self):
        """Google画像検索の結果ページから、一番最初に現れるAmazonのURLを特定します。"""
        search_results = self.driver.find_elements(By.CSS_SELECTOR, "a[href*='amazon.co.jp/']")
        if search_results:
            return search_results[0].get_attribute("href")
        else:
            print("Amazonの商品URLが見つかりませんでした。")
            return None

    def search_amazon_by_image_google(self, image_url):
        """Google画像検索でAmazonの商品を探し、最も関連性の高い商品ページのURLを取得します。"""
        self.initialize_driver()  # ドライバーが未初期化なら初期化
        try:
            self.driver.get("https://images.google.com/")
            time.sleep(1)
            print("Google画像検索ページにアクセスしました。")

            # 画像検索ボタンをクリック
            search_button = WebDriverWait(self.driver, 30).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '[aria-label="画像で検索"]'))
            )
            if search_button.is_displayed():
                print("画像検索ボタンの要素を検出しました。")
                self.driver.execute_script("arguments[0].click();", search_button)
                time.sleep(1)
                print("画像検索ボタンをクリックしました。")
            else:
                print("画像検索ボタンが非表示のためクリックできませんでした")
                return None

            # 画像URLを入力
            image_input = WebDriverWait(self.driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "//input[@class='cB9M7' and @placeholder='画像リンクを貼り付ける']"))
                # @type='text'を削除
            )
            if image_input.is_displayed():
                self.driver.execute_script("arguments[0].focus();", image_input)
                image_input.send_keys(image_url)
                time.sleep(0.5)
                print("画像URLを入力欄に入力しました。")
                image_input.send_keys(Keys.ENTER)  # エンターキーを送信
                time.sleep(2)  # エンターキー送信後の待機
            else:
                print("画像URLの入力欄が非表示のため、入力できませんでした。")
                return None

            # 検索結果から商品ボタンを特定
            product_button = WebDriverWait(self.driver, 40).until(
                EC.presence_of_element_located((By.XPATH, "//div[@role='listitem']/a/div[text()='商品']"))
            )
            if product_button.is_displayed():
                time.sleep(1)
                print("商品ボタンの要素を検出しました。")
                search_button = WebDriverWait(self.driver, 30).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[@role='listitem']/a/div[text()='商品']"))
                )
                search_button.click()
                time.sleep(1)
                print("商品ボタンをクリックしました")
            else:
                print("商品ボタンを検出できませんでした。")
                return None

            time.sleep(1)

            # 検索結果からAmazonのリンクを探す
            amazon_url = self.find_first_amazon_url()
            if amazon_url:
                return amazon_url
            else:
                print("Amazonの商品URLを取得できませんでした。")
                return None

        except Exception as e:
            print(f"Google画像検索エラー: {e}")
            return None

    def login_to_eresa_in_iframe(self):
         """Amazonページ内のiframeでERESAにログインします。"""
         if self.logged_in_eresa:
             print("既にERESAにログイン済みです。")
             return True
         
         try:
             
             # iframeが表示されるまで待機
             print("iframeの表示を待機します...")
             iframe = WebDriverWait(self.driver, 20).until(
             EC.presence_of_element_located((By.XPATH, "//iframe[@data-added-by-eresa='true' and @id='eresa_chart']"))
             )
             print("iframeが表示されました。")

             # iframeに切り替え
             print("iframeに切り替えます...")
             self.driver.switch_to.frame(iframe)
             print("iframeに切り替えました。")

             print("ERESAのログインページにアクセスします...")
             
             # ページが完全にロードされるまで待機
             print("ログインページのロードを待機します...")
             WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
             print("ログインページのロードが完了しました。")
             
             
             # 2. ログイン情報の入力 (placeholderで要素を特定)
             print("ユーザー名入力欄を特定します...")
             username_input = WebDriverWait(self.driver, 20).until(
             EC.presence_of_element_located((By.XPATH, "//input[@placeholder='メールアドレスを入力してください']"))
             )
             print("パスワード入力欄を特定します...")
             password_input = WebDriverWait(self.driver, 20).until(
             EC.presence_of_element_located((By.XPATH, "//input[@placeholder='パスワードを入力してください']"))
             )
             print("ユーザー名とパスワードを入力します...")
             username_input.send_keys(self.eresa_username)
             password_input.send_keys(self.eresa_password)
             print("ユーザー名とパスワードを入力しました。")
             
             # 3. ログインボタンのクリック (クラス名で要素を特定)
             print("ログインボタンを特定します...")
             login_button = WebDriverWait(self.driver, 20).until(
             EC.element_to_be_clickable((By.CLASS_NAME, "login_button"))
             )
             print("ログインボタンをクリックします...")
             login_button.click()
             print("ログインボタンをクリックしました。")
             
             # 4. ログイン後の状態の確認（例：ヘッダーの要素が表示されるまで待機）
             print("ログイン後のヘッダー要素の表示を待機します...")
             WebDriverWait(self.driver, 20).until(
             EC.presence_of_element_located((By.CSS_SELECTOR, "header.header"))
             )
             print("ログイン後のヘッダー要素が表示されました。")
             
             self.logged_in_eresa = True  # ログイン状態を更新
             
             # デフォルトフレームに戻る
             self.driver.switch_to.default_content()
             print("デフォルトフレームに戻りました。")
             
             
             return True
         
         except Exception as e:
             print(f"iframe内のERESAログインエラー: {e}")
             if self.driver:
                try:
                    print(f"現在のURL: {self.driver.current_url}")
                    print(f"現在のページソース: {self.driver.page_source}")
                except:
                   print("ページソースの取得に失敗しました")
             return False
         
    def extract_jan_code_from_amazon(self, amazon_url):
        """Amazonの商品ページからJANコードを抽出します。"""
        self.initialize_driver()  # ドライバーが未初期化なら初期化

        try:
            print(f"Amazonページにアクセスします: {amazon_url}")
            self.driver.get(amazon_url)

            if self.first_amazon_access and self.eresa_username and self.eresa_password:
                print("初回アクセス時のERESAログイン処理を実行します。")
                if not self.login_to_eresa_in_iframe():
                   print("初回アクセス時のERESAログインに失敗しました")
                   return None
                self.first_amazon_access = False #初回アクセスフラグをFalseにする
                self.driver.refresh() # ログイン後、ページをリフレッシュ
                print("ページをリフレッシュしました")

                # iframeが表示されるまで待機
            print("iframeの表示を待機します...")
            iframe = WebDriverWait(self.driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//iframe[@data-added-by-eresa='true' and @id='eresa_chart']"))
            )
            print("iframeが表示されました。")

            # iframeに切り替え
            print("iframeに切り替えます...")
            self.driver.switch_to.frame(iframe)
            print("iframeに切り替えました。")


            # JANコードラベルが表示されるまで待機
            print("JANコードラベルの表示を待機します...")
            WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'JAN') and @class='font-weight-bold border-bottom']"))
            )
            print("JANコードラベルが表示されました。")

            # JANコードの<span>要素と、そのテキストが特定のパターンに一致するまで待機
            print("JANコードの要素が表示されるまで待機します...")
            jan_code_element = WebDriverWait(self.driver, 20).until(
                 lambda driver: self._check_jan_code_span_presence_and_pattern(driver)
            )
            print("JANコードの要素が表示されました。")

            # JANコードを取得
            jan_code = jan_code_element.text.strip()
            print(f"取得したJANコード: {jan_code}")

            # デフォルトフレームに戻る
            self.driver.switch_to.default_content()
            print("デフォルトフレームに戻りました。")

            return jan_code

        except Exception as e:
            print(f"Amazonページエラー: {e}")
            if self.driver:
                try:
                    print(f"現在のURL: {self.driver.current_url}")
                    print(f"現在のページソース: {self.driver.page_source}")
                except:
                   print("ページソースの取得に失敗しました")
            return None
    
    def _check_jan_code_span_presence_and_pattern(self, driver):
        """JANコードの<span>要素が存在し、テキストがパターンに一致するかをチェックします。"""
        try:
            jan_code_container = driver.find_element(By.XPATH, "//div[contains(text(), 'JAN') and @class='font-weight-bold border-bottom']/following-sibling::div")
            jan_code_element = jan_code_container.find_element(By.TAG_NAME, "span")
            
            jan_code_text = jan_code_element.text.strip()
            # 数字のみ、またはハイフンを含む数字列のパターン
            if re.match(r'^([\d-]+)$', jan_code_text):
                return jan_code_element
            else:
                return False # パターンにマッチしない場合
        except:
            return False # 要素が見つからなかった場合


    def close(self):
        if self.driver:
            self.driver.quit()
            self.driver = None


def extract_asin_from_amazon_url(amazon_url):
    """Amazonの商品ページURLからASINを抽出します。"""
    print("Amazon URLからASINを抽出します...")
    match = re.search(r'/dp/([A-Z0-9]+)', amazon_url, re.IGNORECASE)
    if match:
        print("ASINを抽出しました")
        return match.group(1)
    match = re.search(r'/product/([A-Z0-9]+)', amazon_url, re.IGNORECASE)
    if match:
        print("ASINを抽出しました")
        return match.group(1)
    print("ASINを抽出できませんでした")
    return None


def update_spreadsheet_with_jan_code(service, spreadsheet_id, sheet_name, jan_code_column, image_url_column, asin_column, amazon_url_column, row_number, jan_code, image_url, asin, amazon_url):
    """スプレッドシートの指定された行にJANコードと関連情報を入力します。"""
    try:
        sheet = service.spreadsheets()
        print(f"{row_number}行目にJANコードと関連情報を書き込みます...")
        values = [[jan_code, asin, image_url, amazon_url]]
        
        # 列の設定に基づいて範囲を設定
        start_column = jan_code_column
        end_column = amazon_url_column if amazon_url_column else start_column
        
        if end_column < start_column:
            end_column = jan_code_column  # もし設定がおかしかったらJANコードの列で止める
        
        cell_range = f'{sheet_name}!{start_column}{row_number}:{end_column}{row_number}'
            
        body = {'values': values}
        sheet.values().update(spreadsheetId=spreadsheet_id, range=cell_range, valueInputOption='USER_ENTERED', body=body).execute()
        print(f"{row_number}行目にJANコードと関連情報を書き込みました。")
    except Exception as e:
        print(f"スプレッドシート書き込みエラー: {e}")


if __name__ == '__main__':
    config = load_config()
    if not config:
        print("設定ファイルの読み込みに失敗しました。")
        exit()
    credentials_file = config['CREDENTIALS_FILE']
    spreadsheet_id = config['SPREADSHEET_ID']
    sheet_name = config['SHEET_NAME']
    ebay_link_column = config['EBAY_LINK_COLUMN']
    jan_code_column = config.get('JAN_CODE_COLUMN')
    start_row = int(config.get('START_ROW'))
    end_row = int(config.get('END_ROW'))
    image_url_column = config.get('IMAGE_URL_COLUMN')
    asin_column = config.get('ASIN_COLUMN')
    amazon_url_column = config.get('AMAZON_URL_COLUMN')
    crx_path = config.get('CRX_PATH')  # CRXファイルのパスを追加
    eresa_username = config.get('ERESA_USERNAME')
    eresa_password = config.get('ERESA_PASSWORD')

    sheets_service = authenticate_sheets_api(credentials_file)
    if sheets_service:
        ebay_links = get_ebay_links_from_spreadsheet(sheets_service, spreadsheet_id, sheet_name, ebay_link_column, start_row)
        if ebay_links:
            browser = ChromeBrowser(crx_path, eresa_username, eresa_password)  # クラスのインスタンスを作成、crx_path, ユーザー名、パスワードを渡す
            row_number = start_row  # 書き込み開始行を初期化
            try:
                for ebay_url in ebay_links:
                    if row_number > end_row:
                        print(f"指定された終了行 {end_row} に到達したため、処理を終了します。")
                        break  # 終了行に達したらループを中断
                    image_url = get_ebay_image_url(ebay_url)
                    if image_url:
                        print(f"eBayの画像URL: {image_url}")
                        amazon_url = browser.search_amazon_by_image_google(image_url)
                        if amazon_url:
                            print(f"Amazonの商品URL: {amazon_url}")
                            asin = extract_asin_from_amazon_url(amazon_url)
                            if asin:
                                print(f"ASIN: {asin}")
                                jan_code = browser.extract_jan_code_from_amazon(amazon_url)
                                if jan_code:
                                    print(f"JANコード: {jan_code}")
                                    update_spreadsheet_with_jan_code(sheets_service, spreadsheet_id, sheet_name, jan_code_column, image_url_column, asin_column, amazon_url_column, row_number, jan_code, image_url, asin, amazon_url)
                                else:
                                    print("AmazonページでJANコードが見つかりませんでした。")
                                    update_spreadsheet_with_jan_code(sheets_service, spreadsheet_id, sheet_name, jan_code_column, image_url_column, asin_column, amazon_url_column, row_number, "", image_url, asin, amazon_url)
                            else:
                                print("Amazon URLからASINを抽出できませんでした。")
                                update_spreadsheet_with_jan_code(sheets_service, spreadsheet_id, sheet_name, jan_code_column, image_url_column, asin_column, amazon_url_column, row_number, "", image_url, asin, amazon_url)
                        else:
                            print("Amazonの商品URL取得に失敗しました。")
                            update_spreadsheet_with_jan_code(sheets_service, spreadsheet_id, sheet_name, jan_code_column, image_url_column, asin_column, amazon_url_column, row_number, "", image_url, "", "")
                    else:
                        print(f"eBayの画像URL取得に失敗しました。")
                        update_spreadsheet_with_jan_code(sheets_service, spreadsheet_id, sheet_name, jan_code_column, image_url_column, asin_column, amazon_url_column, row_number, "", "", "", "",)
                    row_number += 1  # 次の行へ
            finally:
                browser.close()
        else:
            print("eBayリンクの取得に失敗しました。")
    else:
        print("API認証に失敗しました")