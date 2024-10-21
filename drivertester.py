##autherise by Henry Tsai
import os
import platform
import subprocess
import zipfile
import re
import shutil
from urllib import request
from selenium import webdriver
from selenium.webdriver.chrome.service import Service

# 定義 ChromeDriver 的下載 URL 結構
CHROME_DRIVER_BASE_URL = "https://storage.googleapis.com/chrome-for-testing-public"
DRIVER_PATH = os.path.join(os.getcwd(), 'chromedriver.exe')

# 檢查 ChromeDriver 版本
def get_chromedriver_version():
    try:
        result = subprocess.run([DRIVER_PATH, '--version'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        version_output = result.stdout.strip()
        version_match = re.search(r'(\d+\.\d+\.\d+\.\d+)', version_output)
        if version_match:
            return version_match.group(1)
        else:
            return None
    except Exception as e:
        print("無法檢查 ChromeDriver 版本:", e)
        return None

# 獲取本地 Chrome 瀏覽器版本
def get_local_chrome_version():
    try:
        if platform.system() == "Windows":
            command = r'reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version'
            version = subprocess.check_output(command, shell=True).decode('utf-8')
            chrome_version = version.strip().split()[-1]
        elif platform.system() == "Darwin":  # macOS
            process = subprocess.Popen(['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', '--version'],
                                       stdout=subprocess.PIPE)
            chrome_version, _ = process.communicate()
            chrome_version = chrome_version.decode('utf-8').strip().split(' ')[-1]
        elif platform.system() == "Linux":
            process = subprocess.Popen(['google-chrome', '--version'], stdout=subprocess.PIPE)
            chrome_version, _ = process.communicate()
            chrome_version = chrome_version.decode('utf-8').strip().split(' ')[-1]
        return chrome_version
    except Exception as e:
        print("無法獲取 Chrome 版本:", e)
        return None

# 根據操作系統獲取對應的 OS 標識符
def get_os_identifier():
    if platform.system() == "Windows":
        return "win64"
    elif platform.system() == "Darwin":  # macOS
        return "mac-x64"
    elif platform.system() == "Linux":
        return "linux64"
    else:
        raise Exception("無法識別操作系統")

# 下載對應版本的 ChromeDriver
def download_chromedriver(version):
    try:
        os_identifier = get_os_identifier()
        driver_file = f"chromedriver-{os_identifier}.zip"
        url = f"{CHROME_DRIVER_BASE_URL}/{version}/{os_identifier}/{driver_file}"
        
        print(f"下載 ChromeDriver 中，版本：{version}，網址：{url}")
        driver_zip_path = os.path.join(os.getcwd(), driver_file)

        # 使用 urllib 下載檔案
        request.urlretrieve(url, driver_zip_path)

        # 解壓縮文件
        with zipfile.ZipFile(driver_zip_path, 'r') as zip_ref:
            zip_ref.extractall(os.getcwd())
        os.remove(driver_zip_path)
        print("ChromeDriver 下載完成並解壓縮。")

        # 移動 chromedriver.exe 並刪除資料夾
        move_and_cleanup(os_identifier)

    except Exception as e:
        print("下載 ChromeDriver 時出現錯誤:", e)

# 移動 chromedriver.exe 並刪除解壓縮後的資料夾
def move_and_cleanup(os_identifier):
    folder_name = f"chromedriver-{os_identifier}"
    extracted_driver_path = os.path.join(os.getcwd(), folder_name, 'chromedriver.exe')
    
    if os.path.exists(extracted_driver_path):
        # 將 chromedriver.exe 移動到上一層資料夾
        shutil.move(extracted_driver_path, DRIVER_PATH)
        print(f"已將 chromedriver.exe 移動到 {DRIVER_PATH}")

        # 刪除解壓縮後的資料夾
        shutil.rmtree(os.path.join(os.getcwd(), folder_name))
        print(f"已刪除資料夾 {folder_name}")
    else:
        print(f"未找到解壓縮的 chromedriver.exe 檔案在 {folder_name} 資料夾中")

# 刪除舊的 ChromeDriver
def delete_old_chromedriver():
    if os.path.exists(DRIVER_PATH):
        try:
            os.remove(DRIVER_PATH)
            print("舊版本的 ChromeDriver 已刪除。")
        except Exception as e:
            print(f"刪除舊 ChromeDriver 時出現錯誤: {e}")

# 檢查並設置 ChromeDriver
def check_and_setup_driver():
    chrome_version = get_local_chrome_version()
    if chrome_version:
        print(f"檢測到本地 Chrome 版本: {chrome_version}")
        
        if os.path.exists(DRIVER_PATH):
            chromedriver_version = get_chromedriver_version()
            if chromedriver_version:
                print(f"檢測到現有的 ChromeDriver 版本: {chromedriver_version}")
                # 如果 ChromeDriver 與 Chrome 版本不匹配，刪除舊版本並下載新版本
                if not chromedriver_version.startswith(chrome_version.split('.')[0]):
                    print(f"ChromeDriver 版本 ({chromedriver_version}) 與 Chrome 版本 ({chrome_version}) 不匹配，正在更新...")
                    delete_old_chromedriver()
                    download_chromedriver(chrome_version)
            else:
                print("無法檢測到現有 ChromeDriver 版本，重新下載。")
                delete_old_chromedriver()
                download_chromedriver(chrome_version)
        else:
            print("未檢測到 ChromeDriver，正在下載...")
            download_chromedriver(chrome_version)

# 獲取 Selenium 的 WebDriver
def get_driver(download_dir=None):
    # 如果沒有指定下載路徑，默認為當前工作目錄的 downloads 子目錄
    if download_dir is None:
        download_dir = os.path.join(os.getcwd(), "downloads")

    # 如果下載目錄不存在，創建它
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)

    # 設置 ChromeOptions，指定下載路徑
    options = webdriver.ChromeOptions()
    prefs = {"download.default_directory": download_dir}
    options.add_experimental_option("prefs", prefs)

    # 啟動 ChromeDriver
    service = Service(DRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=options)
    
    return driver