# selenium_sncs

# 功能:

1. 自動化 ChromeDriver 管理：根據已安裝的 Chrome 瀏覽器版本，自動下載並配置正確版本的 ChromeDriver。
2. SNCS 網頁自動化：自動登入 SNCS 平台，選擇特定批號並下載所需的 CSV 檔案。
3. Excel 報告生成：將下載的 CSV 檔案處理為一份整合的 Excel 報告，內含格式化表格、圖表及統計資料（平均值、標準差等）。
4. 即時進度顯示：使用進度條來顯示下載和處理流程中的當前步驟。
5. 進階 Excel 格式化：包括自訂儲存格格式、邊框及條件格式，以提升報告的可讀性。
# 使用套件:
- selenium
- pandas
- openpyxl
- tqdm
- tkinter

使用以下指令安裝所有依賴項：
```bash
pip install -r requirements.txt
```
# 使用方法:
## I. sncs.py
步驟 1: ChromeDriver 設置
   
   確保已安裝 Google Chrome 瀏覽器，腳本會自動根據你的瀏覽器版本處理 ChromeDriver 的下載和配置。
   
步驟 2: 設定環境參數
   
   請在相同層級資料夾下，創建一個`.env`的檔案，裡面包含[XQC](https://sncs-web.com/quality/login)登入帳號及密碼
   `.env`變數名稱如下:
   
```
SNCS_ACCOUNT='ur_account'
SNCS_PASSWORD='ur_password'
```
   
步驟 3: 執行腳本

   要開始 SNCS 資料處理，請執行以下命令：
```bash
python sncs.py
```
此腳本會執行以下操作：

1. 打開一個 GUI 讓你選擇下載資料夾。
2. 使用預設帳戶登入 SNCS 系統。
3. 下載選定的 CSV 檔案。
4. 根據下載的資料生成 Excel 檔案（merge.xlsx），內含詳細的格式及統計資料。
   
步驟 3: Excel 報告輸出
當流程完成後，會將 Excel 報告保存在你指定的資料夾中。
## II. sncs_lot.py
步驟 1: ChromeDriver 設置

確保已安裝 Google Chrome 瀏覽器，腳本會自動根據你的瀏覽器版本處理 ChromeDriver 的下載和配置。

步驟 2: 設定環境參數

請在相同層級資料夾下，創建一個`.env`的檔案，裡面包含[sysmex academy](https://academy.sysmex.com.tw/user/login)登入帳號及密碼
   `.env`變數名稱如下:
   
```
LOT_SNCS_ACCOUNT='ur_account'
LOT_SNCS_PASSWORD='ur_password'
```
   
   要開始 SNCS_lot 資料下載，請執行以下命令：
```bash
python sncs_lot.py
```
此腳本會執行以下操作：

1. 打開一個 GUI 讓你選擇下載資料夾。
2. 在終端(terminal)中請使用者輸入批號的前四碼。
3. 使用預設帳戶登入 academy 系統。
4. 利用使用者輸入的前四碼進行搜尋。
5. 將搜尋結果顯示在終端中，並請使用者選擇要下載哪一個文件。
6. 根據使用者輸入的文件編號，下載相對應的xqn文件，並且進行解壓縮。
   
# 概述
sncs.py：主要腳本，負責 SNCS 自動化、CSV 檔案處理及 Excel 報告生成。

sncs_lot.py：負責 SNCS 新批號設定檔qxn下載。

drivertester.py：包含自動下載、解壓縮並管理 ChromeDriver 版本的函數。

# 授權
本專案基於 MIT 授權條款。
