#autherise by Henry Tsai
import sys,os, glob,zipfile,py7zr,re,csv,threading, drivertester
from dotenv import load_dotenv

from tkinter import filedialog,messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from tqdm import tqdm
from time import time,sleep


load_dotenv()
ac = os.getenv('LOT_SNCS_ACCOUNT')
pw = os.getenv('LOT_SNCS_PASSWORD')


class DOWNLOAD:
    def __init__(self):
        self.folder_path = filedialog.askdirectory()
        self.pbarper = ["輸入新批號","從SNCS爬取所需批號","處理解壓縮檔案","完成"]


    def wait_for_download_complete(self,download_path, timeout=300):
        """
        等待直到下載完成或超時為止。
        :param download_path: 下載的資料夾路徑
        :param timeout: 最長等待秒數，預設 10 秒
        """
        start_time = time()

        while True:
            # 檢查有無以 .tmp 結尾的檔案（Chrome）
            downloading = any(filename.endswith(".crdownload") or filename.endswith(".tmp") for filename in os.listdir(download_path))

            if not downloading:
                # 沒有 .tmp 檔案，表示下載完成
                break

            if time() - start_time > timeout:
                raise TimeoutError("等待下載超過時限，可能下載失敗或卡住。")

            sleep(1)  # 每秒檢查一次

        return True

    def download_newlot(self,step):
        # """使用 match-case 模式匹配來執行每個步驟的具體邏輯"""
        match step:
            case "輸入新批號":
                self.input_lotno = input("\n請刷條碼/輸入新批號:")
                self.current_index+=1   #改變狀態
                # print(self.input_lotno)
            case "從SNCS爬取所需批號":
                print("正在從SNCS爬取所需批號...")
                if not os.path.exists(self.download_dir):
                    os.makedirs(self.download_dir)


                # # 檢查並下載（如果需要） ChromeDriver
                drivertester.check_and_setup_driver()

                driver = drivertester.get_driver(download_dir=self.download_dir)
                #get web address
                driver.get("https://academy.sysmex.com.tw/user/login")
                
                #setting waiting index
                wait = WebDriverWait(driver, 5)
                #key in account & pw
                input_id = wait.until(EC.element_to_be_clickable((By.ID, "edit-name")))
                input_id.send_keys(ac)
                input_pw = wait.until(EC.element_to_be_clickable((By.ID, "edit-pass")))
                input_pw.send_keys(pw)
                #click login btn
                loginbtn = driver.find_element(By.ID, "edit-submit")
                loginbtn.click()
                
                sleep(2)
            #get to the csv website
                driver.get("https://academy.sysmex.com.tw/resource-library/search?s=%s&folder=3536/3538/instructions-for-use"%(self.input_lotno))
                # driver.get("https://academy.sysmex.com.tw/resource-library/search?s=%s&folder=3536/3538/instructions-for-use"%('5013'))
                self.result_title =[]
                self.result_link=[]
                search_h4_result = driver.find_elements(By.TAG_NAME, "h4")
                saerch_div_result = driver.find_element(By.CSS_SELECTOR, "div.mtx-list")
                # print(saerch_a_downloadlink)
                links = saerch_div_result.find_elements(By.CSS_SELECTOR,"a.kb-download")
                # print(links)
                for link in links:
                    self.result_link.append(link.get_attribute("href"))
                # for text in link:
                for title in search_h4_result:
                    self.result_title.append(title.text)
                
                driver.minimize_window()
                if len(self.result_title) == 0:
                    print("根據'%s'查無任何結果!請重新搜索!"%(self.input_lotno))
                    return "input_not_found"

                else:
                    print("目前根據'%s'搜索到的結果如下:"%(self.input_lotno))
                    for i in range(1,len(search_h4_result)+1):
                        print(str(i) + ". " + search_h4_result[i-1].text)
                    self.download_No = input("請選擇要下載的編號:")
                    # sleep(15)
                    driver.get(self.result_link[int(self.download_No) - 1])
                    # 執行下載後
                    sleep(1)
                    self.wait_for_download_complete(self.folder_path)
                    
                    self.current_index+=1   #改變狀態

            
            case "處理解壓縮檔案":
                print("正在解壓縮檔案...")
                # 2. 指定要解壓縮到哪個資料夾
                list_zipfilename =self.result_title[int(self.download_No) - 1].split(".")
                extract_folder = os.path.join(self.folder_path, list_zipfilename[0])
                zipfile_subname = list_zipfilename[1] 
                if zipfile_subname == "zip":
                    #利用壓縮檔名作為資料夾檔名儲存
                    zipfilepath = os.path.join(self.folder_path, self.result_title[int(self.download_No) - 1])
                    with zipfile.ZipFile(zipfilepath, 'r') as zip_ref:
                        zip_ref.extractall(path=self.folder_path)
                elif zipfile_subname == "7z":
                    #利用壓縮檔名作為資料夾檔名儲存
                    zipfilepath = os.path.join(self.folder_path, self.result_title[int(self.download_No) - 1])
                    with py7zr.SevenZipFile(zipfilepath, 'r') as zip_ref:
                        zip_ref.extractall(path=self.folder_path)
                else:
                    return "unzip_error"
                #刪除zip檔
                os.remove(zipfilepath)
                
                self.current_index+=1   #改變狀態

            case "完成":
                messagebox.showinfo('土城長庚醫院檢驗科', message='下載整理成功!檔案在%s'%(self.folder_path))

                self.current_index+=1   #改變狀態
                return "done"


            
    def run(self):
        """控制流程並顯示進度條"""
        if self.folder_path:
            self.current_index = 0  # 以 steps[0] 代表步驟1，依此類推
            self.download_dir = self.folder_path.replace("/","\\")
            with tqdm(total=len(self.pbarper), desc="處理進度", ncols=100) as pbar:
                while self.current_index < len(self.pbarper):
                    result = self.download_newlot(self.pbarper[self.current_index])  # 執行每一個步驟

                    if result =="input_not_found":
                        self.current_index = 0
                        pbar.update(-1)
                        continue
                    elif result =="unzip_error":
                        print("解壓縮失敗，檔案在%s，請自行解壓縮"%(self.folder_path))
                        break
                    elif result =="done":
                        break
                    pbar.set_postfix_str(f"正在處理: {self.pbarper[self.current_index]}")  # 顯示當前步驟
                    pbar.update(1)  # 更新進度條
                # for step in self.pbarper:
                #     self.download_newlot(step)  # 執行每一個步驟
                #     pbar.set_postfix_str(f"正在處理: {step}")  # 顯示當前步驟
                #     pbar.update(1)  # 更新進度條
            print("下載成功!")
            return
        else:
            print("byebye!")
            return



def main():
    D = DOWNLOAD()
    D.run()         # 執行類別中的處理流程


if __name__ == '__main__':  
    main()
