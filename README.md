# eBay商品画像からJANコード取得ツール

## 概要

このツールは、eBayの商品ページの画像URLを基に、Amazonの商品ページを特定し、拡張機能によってAmazonの商品ページに表示されたERESA（イーリサ）の情報からJANコードを取得するまでの一連の手順を自動化します。手作業によるJANコードの収集を効率化できます。

## 必要な環境

このツールを使用するためには、以下の環境が必要です。

- **Python 3.6以上**: プログラムを実行するためのPython環境が必要です。
- **Googleアカウント**: Google Sheets APIを使用するためにGoogleアカウントが必要です。
- **ChromeDriver**: Google Chromeを自動操作するためのChromeDriverが必要です。
- **設定ファイル (`config.ini`)**: APIキーやスプレッドシートIDなどの設定を保存するファイルが必要です。
- **サービスアカウントキーファイル (`JSON`)**: Google Sheets APIを認証するためのサービスアカウントキーファイルが必要です。

## インストール手順

以下のライブラリをインストールする必要があります。ターミナルまたはコマンドプロンプトを開き、以下のコマンドを実行してください。
```bash
pip install requests beautifulsoup4 google-api-python-client google-auth-httplib2 google-auth-oauthlib selenium webdriver-manager configparser
```
設定ファイル (config.ini) の作成
このツールを正しく動作させるためには、config.ini という設定ファイルが必要です。以下の内容を参考に、config.ini ファイルを作成してください。

```
[DEFAULT]
CREDENTIALS_FILE = C:\Users\                                                        # サービスアカウントキーファイルのパス
SPREADSHEET_ID = xxxxxxxxxxxxxxxxxxxxx                                              # GoogleスプレッドシートのID
SHEET_NAME = SHEET NAME                                                             # スプレッドシートのシート名
EBAY_LINK_COLUMN = C                                                                # eBayのURLが記載された列
JAN_CODE_COLUMN = D                                                                 # JANコードを書き込む列
IMAGE_URL_COLUMN = E                                                                # 画像URLを書き込む列
ASIN_COLUMN = F                                                                     # ASINを書き込む列
AMAZON_URL_COLUMN = G                                                               # amazonのURLを書き込む列
START_ROW = 2                                                                       # 処理を開始する行
END_ROW = 10                                                                        # 処理を修了する行
CRX_PATH = C:\Users\                                                                #crxファイルのパス
ERESA_USERNAME = USERNAME                                                           # ERESAのログインユーザー名
ERESA_PASSWORD = PASSWORD                                                           # ERESAのログインパスワード
```
### 各設定項目の説明:

-   `CREDENTIALS_FILE`: Google Cloud Platformで作成したサービスアカウントキーファイルのパスを指定します。
-   `SPREADSHEET_ID`: 処理対象のGoogleスプレッドシートのIDを指定します。スプレッドシートのURLから確認できます。
-   `SHEET_NAME`: 処理対象のスプレッドシート内のシート名を指定します。
-   `EBAY_LINK_COLUMN`: eBayのURLが記載されている列のアルファベットを指定します。（例: `A`, `B`, `C`）
-   `JAN_CODE_COLUMN`: 取得したJANコードを書き込む列のアルファベットを指定します。
-   `IMAGE_URL_COLUMN`: eBayの画像URLを書き込む列のアルファベットを指定します。
-   `ASIN_COLUMN`: AmazonのASINを書き込む列のアルファベットを指定します。
-   `AMAZON_URL_COLUMN`: AmazonのURLを書き込む列のアルファベットを指定します。
-   `START_ROW`: 動作を開始する行を指定します。
-   `END_ROW`: 動作を終了する行を指定します。
-   `CRX_PATH`: CRXファイルのパスを指定します。
-   `ERESA_USERNAME`: ERESAにログインするためのユーザー名を指定します。
-   `ERESA_PASSWORD`: ERESAにログインするためのパスワードを指定します。

## サービスアカウントキーファイル (`JSON`) の準備

Google Sheets APIを使用するには、サービスアカウントキーファイルが必要です。以下の手順で準備してください。

-   **Google Cloud Platform (GCP) にアクセス**: GCPのコンソールにアクセスし、プロジェクトを作成または選択します。
-   **サービスアカウントを作成**: 「IAMと管理」 > 「サービスアカウント」からサービスアカウントを作成します。
-   **キーを作成**: 作成したサービスアカウントの「キー」タブでJSON形式のキーを作成し、ダウンロードします。
-   **設定ファイルにパスを指定**: ダウンロードしたJSONファイルのパスを `config.ini` の `CREDENTIALS_FILE` に指定します。

## ChromeDriver の準備

Seleniumを使ってブラウザを操作するために、ChromeDriverが必要です。

-   **ChromeDriverをダウンロード**: Google Chromeのバージョンに対応したChromeDriverをダウンロードします。
    [https://chromedriver.chromium.org/downloads](https://chromedriver.chromium.org/downloads)
-   **ChromeDriverを配置**: ダウンロードしたChromeDriverを、システムのPATHが通っているディレクトリに配置するか、プログラムと同じディレクトリに配置します。

    *注意*: このプログラムでは、`webdriver-manager`というライブラリを使用するため、自動でChromeDriverをインストールするため、ダウンロードや配置の手順は不要です。

## ツールの実行手順

-   **設定ファイルを準備**: `config.ini` ファイルをプログラムと同じ場所に配置します。
-   **サービスアカウントキーファイルを準備**: ダウンロードしたJSONファイルを適切な場所に配置し、`config.ini` の `CREDENTIALS_FILE` にパスを指定します。
-   **Pythonスクリプトを実行**: ターミナルまたはコマンドプロンプトで、Pythonスクリプト (`your_script_name.py`) を実行します。

    ```bash
    python your_script_name.py
    ```

## 処理の流れ

-   **設定ファイルの読み込み**: `config.ini` ファイルから設定情報を読み込みます。
-   **Google Sheets APIの認証**: サービスアカウントキーファイルを使ってGoogle Sheets APIを認証します。
-   **eBayリンクの取得**: 指定されたスプレッドシートと列からeBayのURLを読み込みます。
-   **各eBay URLに対して以下の処理を実行**:
    -   **eBay画像URLの取得**: eBayの商品ページから画像URLを取得します。
    -   **Amazon商品URLの検索**: Google画像検索を使用して、eBay画像のAmazon商品URLを検索します。
    -   **Amazon ASINの抽出**: 取得したAmazonの商品URLからASINを抽出します。
    -   **ERESAへのログイン**: `config.ini` に ERESA のユーザー名とパスワードが設定されている場合、ERESAにログインします。
    -   **JANコードの抽出**: Amazonの商品ページに埋め込まれたERESAの商品ページからJANコードを抽出します。
    -   **スプレッドシートへの書き込み**: 取得したJANコード、画像URL、ASIN、Amazon URLを指定したスプレッドシートの行に書き込みます。
-   **処理の終了**: 全てのeBay URLの処理が完了したら、プログラムを終了します。

## 注意事項

-   画像認識の精度上、実際に取得できるJANコードが必ずしも正確であるとは限りません。
-   このツールは、Google Sheets API、Google画像検索、ERESAのウェブサイトの構造に依存しています。これらの構造が変更された場合、ツールの動作が正常でなくなる可能性があります。
-   各ウェブサイトへのアクセス頻度が高すぎる場合、アクセス制限を受ける可能性があります。
-   設定ファイルの `START_ROW` は、データの開始行を指定します。
-   **`ERESA_USERNAME`** と **`ERESA_PASSWORD`** は、ERESAにログインするためのユーザー名とパスワードを設定します。これらの情報が設定されていない場合、ERESAからのJANコード取得はスキップされます。

## エラー対応

-   エラーが発生した場合、コンソールにエラーメッセージが表示されます。
-   エラーメッセージを参考に、設定ファイルやスクリプトに問題がないかを確認してください。
-   エラーが解決しない場合は、開発者にお問い合わせください。
