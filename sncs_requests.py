import requests
import os

def download_global_qc_files(save_path,cookie_session):
    """
    先登入網站(帶帳密), 再用已登入的 session 去呼叫需要權限的 API 以下載檔案
    """

    # 如果指定的資料夾不存在，則自動建立
    if not os.path.exists(save_path):
        os.makedirs(save_path)
    
    # 1. 建立一個 Session，用來維持 cookies
    session = cookie_session


    # 3. 帶著已登入的 session 呼叫需要授權的 API
    get_url = "https://sncs-web.com/quality/api/csvFileControlLot"
    try:
        response = session.get(get_url)
        response.raise_for_status()
        data = response.json()
    except Exception as e:
        print("取得 controlLotNo 列表失敗:", e)
        return

    try:
        control_lot_list = data["controlList"][0]["controlLotList"]
    except (KeyError, IndexError) as e:
        print("解析 JSON 時發生錯誤，無法正確取得 controlLotNo 列表:", e)
        return

    first_three_lots = control_lot_list[:3]

    post_url = "https://sncs-web.com/quality/api/controlLotCsvFileDownload"
    instrument_code_list = ["8000031576", "8000031586"]
    threepath=[]
    for item in first_three_lots:
        control_lot_no = item.get("controlLotNo")
        control_lot_disp = item.get("controlLotDisp")
        if not control_lot_no:
            continue

        payload_download = {
            "controlLotNo": control_lot_no,
            "instrumentCodeList": instrument_code_list
        }

        try:
            post_resp = session.post(post_url, json=payload_download)
            post_resp.raise_for_status()
            post_data = post_resp.json()
            download_url = post_data.get("urlToDownloadFile")
            if not download_url:
                print(f"[{control_lot_no}] 無法取得下載連結，跳過。")
                continue
        except Exception as e:
            print(f"[{control_lot_no}] 無法取得下載連結:", e)
            continue

        custom_filename = f"{control_lot_disp}_global_QC.zip"
        save_full_path = os.path.join(save_path, custom_filename)
        threepath.append(save_full_path)
        try:
            # 4. 使用同一個 session 去拿下載檔
            #    (若該下載連結也需要驗證Cookie才能讀取，就必須 session.get(download_url))
            file_resp = session.get(download_url)
            file_resp.raise_for_status()
            with open(save_full_path, "wb") as f:
                f.write(file_resp.content)
            print(f"已下載: {save_full_path}")
        except Exception as e:
            print(f"[{control_lot_no}] 下載失敗:", e)

    print("全部下載流程結束。")
    return threepath

if __name__ == "__main__":

    # 要將檔案下載到哪個資料夾
    folder_path = "E:\土城長庚醫院\新下載的\equests"
    session = "ur_cookies"
    download_global_qc_files(folder_path,session)
