import streamlit as st  # streamlit=1.47.1
import pandas as pd     # pandas=2.3.1
import os, time, json, random

from selenium import webdriver  # selenium=4.34.2
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

import shutil
from openpyxl import Workbook, load_workbook    # openpyxl=3.1.5
from openpyxl.styles import PatternFill
from io import BytesIO
import xlsxwriter   # xlsxwriter=3.2.5
import tempfile
from datetime import datetime
import matplotlib.pyplot as plt
import numpy as np


import socket

tam='''
    def wait_for_download(download_dir, timeout=60):
        """Chờ đến khi có file trong thư mục download"""
        start_time = time.time()
        while True:
            files = os.listdir(download_dir)
            if files:  # có ít nhất 1 file
                files = [os.path.join(download_dir, f) for f in files]
                # nếu nhiều file thì chọn file mới chỉnh sửa gần nhất
                return max(files, key=os.path.getmtime)
            if time.time() - start_time > timeout:
                raise TimeoutError("Không thấy file tải về trong thời gian chờ")
            time.sleep(1)


    def download_file_with_selenium(url, new_name):
        """Mở url, click tải file, đổi tên file thành new_name, trả về path"""
        # 1. Tạo thư mục tạm trên server
        download_dir = tempfile.mkdtemp()

        chrome_options = webdriver.ChromeOptions()
        prefs = {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        chrome_options.add_experimental_option("prefs", prefs)
        chrome_options.add_argument("--headless=new")  # chạy không cần giao diện
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")

        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=chrome_options
        )

        # 2. Mở trang web
        driver.get(url)

        # TODO: chèn bước click nút download tại đây
        # driver.find_element("xpath", "//button[text()='Download']").click()

        # 3. Đợi file tải xong
        latest_file = wait_for_download(download_dir)

        # 4. Đặt tên file mới
        new_path = os.path.join(download_dir, new_name)
        os.rename(latest_file, new_path)

        driver.quit()
        return new_path


    # ================== STREAMLIT APP ==================
    st.title("📥 Demo tải file và đổi tên bằng Selenium")

    if st.button("Tải dữ liệu"):
        try:
            # Tải file và đổi tên
            url = "https://example.com"  # thay link thật
            new_name = f"data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

            file_path = download_file_with_selenium(url, new_name)

            # Cho phép tải về local
            with open(file_path, "rb") as f:
                st.download_button(
                    label="⬇️ Tải file đã đổi tên",
                    data=f,
                    file_name=new_name,
                    mime="text/plain"
                )

        except Exception as e:
            st.error(f"Lỗi: {e}")
    '''
# cac vung de chon-----------------------------------------------------
#@st.cache_data
def download_data_smarts(regions):
    #xoa thu muc downloads va tao lai de chi chua 2 file du lieu
    #folder_path_cu = 'downloads'
    # Xóa thư mục nếu tồn tại
    #if os.path.exists(folder_path_cu):
    #    shutil.rmtree(folder_path_cu)  # Xóa toàn bộ thư mục và nội dung bên trong

    #download_dir = os.path.abspath("downloads")
    #os.makedirs(download_dir, exist_ok=True)
    download_dir = tempfile.mkdtemp()

    # ✅ CẤU HÌNH CHROME:
    options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_setting_values.automatic_downloads": 1   # DÒNG QUAN TRỌNG DE TAT THONG BAO
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--headless")  # chạy ẩn trình duyệt

    # ✅ KHỞI TẠO TRÌNH DUYỆT
    driver = webdriver.Chrome(options=options)

    driver.get("https://smarts.waterboards.ca.gov/smarts/SwPublicUserMenu.xhtml")
    print("✅ Vào trang chính")

    WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.LINK_TEXT, "Download NOI Data By Regional Board"))
    ).click()

    WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 1)
    driver.switch_to.window(driver.window_handles[-1])
    print("✅ Đã chuyển sang tab mới")

    links = [
        "Industrial Application Specific Data",
        "Industrial Ad Hoc Reports - Parameter Data",
        "Industrial Annual Reports"
    ]

    def wait_for_download_and_get_new_file(before_files, timeout=40):
        for _ in range(timeout * 2):
            time.sleep(0.5)
            after_files = set(os.listdir(download_dir))
            new_files = after_files - before_files
            txt_files = [f for f in new_files if f.endswith(".txt")]
            if txt_files:
                return txt_files[0]
        return None
    #---------------------
    region = regions
    print(f"\n🔹 Chọn Region: {region}")
    dropdown = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.NAME, "intDataFileDowloaddataFileForm:intDataDumpSelectOne"))
    )
    Select(dropdown).select_by_visible_text(region)
    time.sleep(3)  # Đợi dropdown load lại
    
    lfile_datai = []

    for j, name in enumerate(links):
        try:
            print(f"📥 Đang click tải: {name}")
            before = set(os.listdir(download_dir))

            link_elem = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.LINK_TEXT, name))
            )
            driver.execute_script("arguments[0].click();", link_elem)

            fname = wait_for_download_and_get_new_file(before)
            if fname:
                # Tạo tên file chuẩn theo Region + tên file
                src = os.path.join(download_dir, fname)
                dst_name = f"{region} - {name}.txt"
                dst_name = dst_name.replace(" ", "_")  # Nếu muốn
                dst = os.path.join(download_dir, dst_name)
                os.rename(src, dst)
                print(f"File đã lưu: {dst}")
                lfile_datai.append(f"{dst}")
                # Hiển thị nút tải file về local
                with open(dst, "rb") as f:
                    st.download_button(f"⬇️ Tải file {j+1}", 
                        f, 
                        file_name=dst
                    )

            else:
                print("❌ Không tìm thấy file mới sau khi tải")
        except Exception as e:
            print(f"❌ Lỗi khi tải {name} ở Region {region}: {e}")

    driver.quit()
    print("\n🎉 Hoàn tất tải file cho "+region)
    return lfile_datai


#------------------------------------------
regions = st.selectbox("Select a Region:", 
            ("Region 1 - North Coast",
            "Region 2 - San Francisco Bay",
            "Region 3 - Central Coast",
            "Region 4 - Los Angeles",
            "Region 5F - Fresno",
            "Region 5R - Redding",
            "Region 5S - Sacramento",
            "Region 6A - South Lake Tahoe",
            "Region 6B - Victorville",
            "Region 7 - Colorado River Basin",
            "Region 8 - Santa Ana",
            "Region 9 - San Diego"),
            index=None,
            placeholder="No selected Region",
            )
#neu mot vung duoc chon thi lam
if regions:
    placeholder_1 = st.empty()
    placeholder_1.write('Wait for downloading 3 files of ' + regions)
    try :
        #thuc thi ham download_data_smarts(regions) va tra ve list cac file da tai 
        lfile_datai = download_data_smarts(regions)
        #placeholder_1.write('Downloaded files:')
        #st.write(lfile_datai)
    except:
        placeholder_1.write('Tải file không đạt!')

