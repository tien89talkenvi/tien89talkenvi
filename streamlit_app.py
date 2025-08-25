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



# ---------------------------
# Bảng ngưỡng tham chiếu (ví dụ)
# Bạn có thể cập nhật lại theo tài liệu chuẩn của bạn
# ---------------------------
thresholds = {
    "Zinc": 0.26,    # mg/L
    "TSS": 100,      # mg/L
    "COD": 120,      # mg/L
    "BOD": 30,       # mg/L
    "pH_low": 6.0,   # std units
    "pH_high": 9.0
}

#-------------------------
def ThucThiPhan_I():
    LOI='OK'
    # Ham tai file txt du lieu dang cvs cua cac mien thuoc bang Cali
    def download_data_smarts(regions):
        #xoa thu muc downloads va tao lai de chi chua 2 file du lieu
        folder_path_cu = 'downloads'
        # Xóa thư mục nếu tồn tại
        if os.path.exists(folder_path_cu):
            shutil.rmtree(folder_path_cu)  # Xóa toàn bộ thư mục và nội dung bên trong

        download_dir = os.path.abspath("downloads")
        os.makedirs(download_dir, exist_ok=True)

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
                else:
                    print("❌ Không tìm thấy file mới sau khi tải")
            except Exception as e:
                print(f"❌ Lỗi khi tải {name} ở Region {region}: {e}")

        driver.quit()
        print("\n🎉 Hoàn tất tải file cho "+region)
        return lfile_datai
        # CHU Y rang neu ten file dat trung voi file da co thi that bai.

    #-phan chinh cchon------------------------
    # cac vung de chon
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
    LOI='OK'
    if regions:
        placeholder_1 = st.empty()
        placeholder_1.write('Wait for downloading 2 files of ' + regions)
        #thuc thi ham download_data_smarts(regions) va tra ve list cac file da tai 
        try :
            lfile_datai = download_data_smarts(regions)
            placeholder_1.write('Downloaded files:')
            st.write(*lfile_datai,sep='\n')
        except:
            LOI='LOI'
            placeholder_1.write('Tai file không đạt!')
    if LOI == 'LOI':
        st.write('Nếu không đạt, '+ ':red[ mở trực tiếp trang sau làm theo các bước để tải:]')
        st.markdown("1-[Open Page SMARTS](https://smarts.waterboards.ca.gov/smarts/SwPublicUserMenu.xhtml)", unsafe_allow_html=True)
        st.write('2-Click on “Download NOI Data By Regional Board”')
        st.write('3-Select your region from the dropdown menu')
        st.write('4-Click on both “Industrial Application Specific Data” and “Industrial Ad Hoc Reports - Parameter Data”')
        st.write('5-Data will be downloaded to two separate .txt files, each titled “file”')
        st.write('6-Nên đổi tên 2 file thành Industrial_Application_Specific_Data và Industrial_Ad_Hoc_Reports_-_Parameter_Data rồi chép vào thư mục riêng của bạn để dễ làm việc ở các bước sau.')

#-------------------
#@st.cache_data
def ThucThiPhan_II(sheet2,sheet1,sheet3):
    if sheet2 and sheet1:
        # Đọc sheet2.txt (Application Specific Data)
        df1 = pd.read_csv(sheet2, sep="\t", dtype=str, encoding='cp1252')
        #st.write(df1)

        # Đọc sheet1.txt (Ad Hoc Reports Data)
        df2 = pd.read_csv(sheet1, sep="\t", dtype=str)
        #st.write(df2)


        # Đảm bảo APP_ID cùng kiểu dữ liệu
        df1["APP_ID"] = df1["APP_ID"].astype(str)
        df2["APP_ID"] = df2["APP_ID"].astype(str)


        # Merge dữ liệu
        df = df2.merge(df1, on="APP_ID", how="left")

        # Giữ lại STATUS=Active
        df = df[df["STATUS"] == "Active"].copy()

        # Chuyển RESULT sang số
        df["RESULT"] = pd.to_numeric(df["RESULT"], errors="coerce")

        # Chuyển đổi đơn vị µg/L → mg/L nếu cần
        df["UNITS"] = df["UNITS"].str.strip()
        mask = df["UNITS"] == "µg/L"
        df.loc[mask, "RESULT"] = df.loc[mask, "RESULT"] / 1000
        df.loc[mask, "UNITS"] = "mg/L"

        # Thêm cột OLD/NEW (mặc định = "New")
        df["OLD/NEW"] = "New"


        # Thêm cột EXCEED và NOTES
        df["EXCEED"] = False
        df["NOTES"] = ""

        # --- Check pH ---
        mask_ph = df["PARAMETER"] == "pH"
        df.loc[mask_ph, "EXCEED"] = ~df.loc[mask_ph, "RESULT"].between(
            thresholds["pH_low"], thresholds["pH_high"], inclusive="both"
        )
        df.loc[mask_ph & df["EXCEED"], "NOTES"] = (
            f"pH out of range ({thresholds['pH_low']}–{thresholds['pH_high']})"
        )

        # --- Check các parameter còn lại ---
        for param, limit in thresholds.items():
            if param in ["pH_low", "pH_high"]:
                continue
            mask_param = df["PARAMETER"] == param
            df.loc[mask_param, "EXCEED"] = df.loc[mask_param, "RESULT"] > limit
            df.loc[mask_param & df["EXCEED"], "NOTES"] = "> NAL=" + str(limit)


        # Đảm bảo đúng thứ tự 19 cột
        final_cols = [
            "WDID", "APP_ID", "STATUS", "FACILITY_NAME", "OPERATOR_NAME",
            "ADDRESS", "CITY", "STATE", "ZIP",
            "PRIMARY_SIC", "SECONDARY_SIC", "TERTIARY_SIC",
            "PARAMETER", "RESULT", "UNITS", "REPORTING_YEAR",
            "OLD/NEW", "EXCEED", "NOTES"
        ]

        # Thêm cột bị thiếu
        for col in final_cols:
            if col not in df.columns:
                df[col] = ""

        # Lấy đúng thứ tự 19 cột
        df_final = df[final_cols].copy()


        # Xuất ra Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df_final.to_excel(writer, sheet_name="Data", index=False)

            # Lấy workbook và worksheet
            workbook  = writer.book
            worksheet = writer.sheets["Data"]

            # Freeze Panes: cố định dòng đầu và 4 cột đầu
            worksheet.freeze_panes(1, 4)  
            # (1,4) nghĩa là khóa hàng trên dòng 2 và cột trước cột E

        tbaodong1.success("✅ Dữ liệu đã được xử lý xong!")

        # Nút tải về
        st.download_button(
            label="📥 Tải Excel kết quả",
            data=output.getvalue(),
            file_name="SMARTS_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


        # Hiển thị dataframe
        st.write(df_final.shape)
        st.dataframe(df_final)
    else:
         tbaodong1.write('Xử lí không thành công!')

#----------------------
def ThucThiPhan_III():
    pass


# -----------------------------------
# Streamlit App
# -----------------------------------
st.header("📊 SMARTS Data Processing")
st.markdown('''
    This program will analyze the wastewater data sets of reported 
    industrial establishments to extract the necessary information 
    to illustrate them graphically to help monitor, inspect and 
    evaluate the environmental management in this field in a more 
    legal and effective manner.''')

# I. TAI FILES DU LIEU TXT TU SMARTS -------------------------
st.subheader('✅ Download the data', divider=True)
st.write("3 tập dữ liệu TXT dạng CVS cần có để làm việc cho mỗi Region_N nên đặt lại tên dạng như dưới đây sau khi tải về (lúc đầu tên chúng là file.txt hoặc file(1).txt, file(2).txt...)) : ")
st.write("Region_N_Industrial_Ad_Hoc_Reports_Parameter_Data.txt")
st.write("Region_N_Industrial_Application_Specific_Data.txt")
st.write("Region_N_Industrial_Annual_Reports.txt")

checkboxI = st.checkbox("📌Chọn tải về 3 files với Region sẽ chọn , :red[nếu chưa có] : ", key='P1', value=False)
if checkboxI:
    ThucThiPhan_I()

# II XU LI VA TAO DATA MOI ------------------------------------
st.subheader('✅ Analyze the new data and creat Sheet', divider=True)
sheet1=None
sheet2=None
sheet3=None
checkboxII = st.checkbox("📌:blue[Lấy 3 files TXT (Application Specific Data, Ad Hoc Reports Data, Industrial Annual Reports) để xử lí và tạo dữ liệu mới]", key='P2', value=False)
if checkboxII:
    laydatafrom = st.radio(
        "WHERE GET DATA ", 
        [":blue[From Local]",":green[From Datatest]", ":red[Empty]"],
        index=2,horizontal=True , label_visibility="visible"
    ) 

    if laydatafrom==":red[Empty]":
        pass  

    elif laydatafrom==":green[From Datatest]":
        sheet2 = "Datatest/sheet2.txt"
        sheet1 = "Datatest/sheet1.txt"
        if not (os.path.exists(sheet2) and os.path.exists(sheet1)):
            tbaodong1 = st.empty()
            tbaodong1.write(f"Chưa có : {sheet2}, {sheet1}")
        else:
            tbaodong1 = st.empty()
            tbaodong1.success(f"Dữ liệu đã lấy là : :blue[{sheet2}], :green[{sheet1}]. :red[⏳ Đang xử lý...]")
            ThucThiPhan_II(sheet2, sheet1, sheet3)

    elif laydatafrom==":blue[From Local]":
        sheet2=None
        sheet1=None
        sheet3=None

        sheet2_ok=False
        sheet1_ok=False
        sheet3_ok=False
        ## Tai len file Industrial Application Specific Data dong vai tro shhet2
        sheet2 = st.file_uploader("**1**. Upload file: :red[Industrial Application Specific Data]", type=["txt"], key='UF1')
        # File mẫu có sẵn trên server ung voi sheet2
        headers_mau = pd.read_csv("Headers/sheet2_headers.txt", nrows=0, encoding='cp1252').columns.tolist()
        if sheet2 is not None:
            # Đọc trực tiếp file upload sheet2
            df_upload = pd.read_csv(sheet2, nrows=0, encoding='cp1252')
            # lay headers cua no
            headers_upload = df_upload.columns.tolist()
            # so sanh de lay lai hoac di tiep
            if headers_upload == headers_mau:
                st.success("👍File Industrial Application Specific Data was uploaded successfully.")
                sheet2_ok = True
            else:
                sheet2_ok = False
                st.warning("😩Invalid file please select another file...")

        if sheet2_ok == True:
            ## Tai len file Industrial Ad Hoc Reports Data dong vai tro shhet1
            sheet1_ok=False
            sheet1 = st.file_uploader("**2**. Upload file: :red[Industrial Ad Hoc Reports Data]", type=["txt"], key='UF2')
            # File mẫu có sẵn trên server ung voi sheet2
            headers_mau1 = pd.read_csv("Headers/sheet1_headers.txt", nrows=0, encoding='cp1252').columns.tolist()
            if sheet1 is not None:
                # Đọc trực tiếp file upload sheet2
                df_upload1 = pd.read_csv(sheet1, nrows=0, encoding='cp1252')
                # lay headers cua no
                headers_upload1 = df_upload1.columns.tolist()
                # so sanh de lay lai hoac di tiep
                if headers_upload1 == headers_mau1:
                    st.success("👍File Industrial Ad Hoc Reports Data was uploaded successfully.")
                    sheet1_ok = True
                else:
                    sheet1_ok = False
                    st.warning("😩Invalid file please select another file...")

        if sheet1_ok == True:
            ## Tai len file Industrial Annual Reports dong vai tro shhet3
            sheet3_ok=False
            sheet3 = st.file_uploader("**3**. Upload file: :red[Industrial Annual Reports]", type=["txt"], key='UF3')
            # File mẫu có sẵn trên server ung voi sheet2
            headers_mau3 = pd.read_csv("Headers/sheet3_headers.txt", nrows=0, encoding='cp1252').columns.tolist()
            if sheet3 is not None:
                # Đọc trực tiếp file upload sheet2
                df_upload3 = pd.read_csv(sheet3, nrows=0, encoding='cp1252')
                # lay headers cua no
                headers_upload3 = df_upload3.columns.tolist()
                # so sanh de lay lai hoac di tiep
                if headers_upload3 == headers_mau3:
                    st.success("👍File Industrial Annual Reports was uploaded successfully.")
                    sheet3_ok = True
                else:
                    sheet3_ok = False
                    st.warning("😩Invalid file please select another file...")

        if sheet2_ok == True and sheet1_ok == True and sheet3_ok == True:
            tbaodong1 = st.empty()
            tbaodong1.success("✅3 file txt were uploaded succesheader. :red[⏳ Đang xử lý...]")
            # phai cho 3 upload_file ve dau vi trong buoc lay header no da xuong vai dong roi
            sheet2.seek(0)
            sheet1.seek(0)
            sheet3.seek(0)
            ThucThiPhan_II(sheet2, sheet1, sheet3)


# IV DO THI HOA DU LIEU -------------------------
st.subheader('✅ Visualize the data', divider=True)
ThucThiPhan_III()
#⏱️

#============================================================================
def update_checkbox_sidebar(tep_mo):
    with open(tep_mo, "r", encoding="utf-8") as f:
        data = json.load(f)   
    # Ghép key và value thành chuỗi "key - value"
    if '1' in tep_mo:
        options1 = [f"{k} - {v}" for k, v in data.items()]
        for op in options1:
            if '#' not in op: 
                st.write(op)
    elif '2' in tep_mo:
        options2 = [f"{k} - {v}" for k, v in data.items()]
        for op in options2:
            if '#' not in op: 
                st.write(op)
    elif '3' in tep_mo:
        options3 = [f"{k} - {v}" for k, v in data.items()]
        for op in options3:
            if '#' not in op: 
                st.write(op)
    else:
        options4 = [f"{k} - {v}" for k, v in data.items()]
        for op in options4:
            if '#' not in op: 
                st.write(op)

def Xem_do_thi_1():
    # Example data
    data = {
        "Month": ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
        "Sales": [150, 200, 250, 220, 300, 400],
        "Expenses": [100, 120, 180, 160, 210, 280]
    }
    # Create DataFrame
    df = pd.DataFrame(data)
    # Set index for better x-axis labels
    df.set_index("Month", inplace=True)
    st.title("📈 Company Sales and Expenses Over Months")
    # Line chart
    st.line_chart(df)


def Xem_do_thi_2():
    # Example data
    data = {
        "Product": ["A", "B", "C", "D"],
        "Sales": [300, 450, 150, 500],
        "Profit": [80, 120, 300, 20]
    }
    # Create DataFrame
    df = pd.DataFrame(data)
    # Set 'Product' as index (so it appears on x-axis)
    df.set_index("Product", inplace=True)
    st.title("📊 Sales and Profit by Product")
    # Streamlit bar chart
    st.bar_chart(df)

def Xem_do_thi_3():
    data = {
        "Product": ["A", "B", "C", "D"],
        "Sales": [300, 450, 150, 500],
        "Profit": [80, 120, 50, 200]
    }

    # Create DataFrame
    df = pd.DataFrame(data)
    df.set_index("Product", inplace=True)

    st.title("📊 Horizontal Bar Plot Example")

    # Create matplotlib horizontal bar plot
    fig, ax = plt.subplots()
    df.plot(kind="barh", ax=ax)

    # Customize
    ax.set_title("Sales and Profit by Product")
    ax.set_xlabel("Amount ($)")
    ax.set_ylabel("Product")

    # Show in Streamlit
    st.pyplot(fig)


def Xem_do_thi_4():
    # Giả lập dữ liệu: thu nhập của 1000 nhân viên (triệu VND)
    np.random.seed(42)
    incomes = np.random.normal(loc=15, scale=5, size=1000)  # trung bình 15, lệch chuẩn 5
    incomes = np.clip(incomes, 5, 50)  # Giới hạn từ 5 đến 50 triệu

    df = pd.DataFrame({"Income": incomes})

    # Vẽ histogram
    fig, ax = plt.subplots()
    df["Income"].hist(bins=20, ax=ax, edgecolor="black")
    ax.set_title("Biểu đồ Histogram về Phân bố thu nhập nhân viên")
    ax.set_xlabel("Thu nhập (triệu VND)")
    ax.set_ylabel("Số lượng nhân viên")
    st.pyplot(fig)

def Xem_do_thi_5():
    st.title("📦 Phân bố điểm thi của học sinh 3 lóp A,B,C bằng Box Plot")
    # Tạo dữ liệu giả lập
    np.random.seed(42)
    data = {
        "Class": (["A"] * 30) + (["B"] * 30) + (["C"] * 30),
        "Score": list(np.random.normal(75, 10, 30)) +   # Lớp A: trung bình 75, lệch chuẩn 10
                list(np.random.normal(65, 15, 30)) +   # Lớp B: trung bình 65, lệch chuẩn 15
                list(np.random.normal(80, 50, 30))      # Lớp C: trung bình 80, lệch chuẩn 5
    }

    df = pd.DataFrame(data)
    # Vẽ box plot
    fig, ax = plt.subplots()
    df.boxplot(column="Score", by="Class", ax=ax)
    # Tùy chỉnh
    ax.set_title("So sánh phân bố điểm thi giữa các lớp")
    ax.set_xlabel("Lớp học")
    ax.set_ylabel("Điểm số")
    plt.suptitle("")  # Xóa tiêu đề mặc định của pandas

    # Hiển thị trong Streamlit
    st.pyplot(fig)
    
    st.markdown("""
    ✅ Ở ví dụ này:
    Lớp A có điểm khá ổn định quanh 75.
    Lớp B phân bố rộng, nhiều học sinh chênh lệch.
    Lớp C tập trung quanh 80, ít biến động.
    👉 Đây chính là tình huống điển hình mà chỉ box plot mới diễn tả được, 
    còn bar chart chỉ cho bạn thấy trung bình, mất hết thông tin về phân bố.
    """)
    st.markdown("""
    ### 📌 Vì sao dùng Box Plot?
    - Hiển thị **median (trung vị)**: mức điển hình của lớp.
    - Cho thấy **khoảng tứ phân vị (IQR)**: độ phân tán.
    - Thấy ngay **outliers (điểm bất thường)**, ví dụ học sinh điểm quá thấp hoặc quá cao.
    - Các biểu đồ khác như **bar chart, line chart** không thể hiện được những thông tin này. 
    [Read more here](http://sociologyhue.edu.vn/blog/post/22288)
    """)

def Xem_do_thi_6():
    st.title("🌈 Area Plot Example")
    # Example dataset
    data = {
        "Month": ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
        "Sales": [150, 200, 250, 220, 300, 400],
        "Expenses": [100, 120, 180, 160, 210, 280]
    }
    # Create DataFrame
    df = pd.DataFrame(data)
    df.set_index("Month", inplace=True)

    # Create matplotlib figure
    fig, ax = plt.subplots()
    df.plot(kind="area", alpha=0.5, ax=ax)  # alpha để trong suốt nhìn rõ chồng lên nhau

    # Customize
    ax.set_title("Sales vs Expenses (Area Plot)")
    ax.set_xlabel("Month")
    ax.set_ylabel("Amount ($)")
    ax.grid(True)

    # Show in Streamlit
    st.pyplot(fig)

def Xem_do_thi_7():
    st.title("🥧 Pie Chart Example")

    # Example dataset
    data = {
        "Product": ["A", "B", "C", "D"],
        "Sales": [300, 450, 150, 500]
    }

    # Create DataFrame
    df = pd.DataFrame(data)

    # Create matplotlib figure
    fig, ax = plt.subplots()
    ax.pie(df["Sales"], labels=df["Product"], autopct="%1.1f%%", startangle=90)
    ax.set_title("Sales Distribution by Product")

    # Show in Streamlit
    st.pyplot(fig)

def Xem_do_thi_8():
    st.title("🔹 Scatter Plot Example")
    # Example dataset
    np.random.seed(42)
    data = {
        "Advertising": np.random.randint(50, 200, 20),  # Chi phí quảng cáo
        "Sales": np.random.randint(100, 500, 20)        # Doanh số
    }

    # Create DataFrame
    df = pd.DataFrame(data)

    # Create matplotlib figure
    fig, ax = plt.subplots()
    ax.scatter(df["Advertising"], df["Sales"], color="blue", s=100, alpha=0.7, edgecolors="k")

    # Customize
    ax.set_title("Sales vs Advertising")
    ax.set_xlabel("Advertising ($)")
    ax.set_ylabel("Sales ($)")
    ax.grid(True)

    # Show in Streamlit
    st.pyplot(fig)

def Xem_do_thi_9():
    st.title("🔷 Hexbin Plot Example")
    # Tạo dữ liệu ví dụ
    np.random.seed(42)
    x = np.random.randn(1000) * 50 + 200   # Dữ liệu Advertising
    y = np.random.randn(1000) * 80 + 300   # Dữ liệu Sales

    df = pd.DataFrame({"Advertising": x, "Sales": y})

    # Tạo figure
    fig, ax = plt.subplots(figsize=(7,5))

    # Vẽ hexbin plot
    hb = ax.hexbin(df["Advertising"], df["Sales"], gridsize=30, cmap="Blues", mincnt=1)

    # Thêm colorbar
    cb = fig.colorbar(hb, ax=ax)
    cb.set_label("Number of points")

    # Tùy chỉnh
    ax.set_title("Sales vs Advertising (Hexbin Plot)")
    ax.set_xlabel("Advertising ($)")
    ax.set_ylabel("Sales ($)")
    ax.grid(True)

    # Hiển thị trong Streamlit
    st.pyplot(fig)


# PHU LUC SIDEBAR -------------------------
with st.sidebar:
    st.header('🏷️ :red[LOOK UP DOCUMENT]')
    # Xem tai lieu SMARTS
    st.write("---")
    checkbox_sidebar_0 = st.checkbox(":blue[📌 SMARTS documents used as a basis for writing this program]", key='PL0', value=False)
    if checkbox_sidebar_0:
        # Đọc nội dung file Markdown
        with open("hd-lam-app-cho-thong.md", "r", encoding="utf-8") as f:
            md_content = f.read()
        st.markdown(md_content, unsafe_allow_html=True)
    
    st.write("---")
    checkbox_sidebar_1 = st.checkbox(":green[📌 Headers in Sheet1 (Industrial_Ad_Hoc_Reports)]", key='PL1', value=False)
    if checkbox_sidebar_1:
        tep_1 = "Headers/dict_sheet1.json"
        update_checkbox_sidebar(tep_1)

    st.write("---")
    # Xem Header Sheet2 
    checkbox_sidebar_2 = st.checkbox(":green[📌 Headers in Sheet2 (Industrial_Application_Specific_Data)]", key='PL2', value=False)
    if checkbox_sidebar_2:
        tep_2 = "Headers/dict_sheet2.json"
        update_checkbox_sidebar(tep_2)


    st.write("---")
    # Xem Header Sheet3 
    checkbox_sidebar_3 = st.checkbox(":green[📌 Headers in Sheet3 (Industrial_Annual_Reports)]", key='PL3', value=False)
    if checkbox_sidebar_3:
        tep_3 = "Headers/dict_sheet3.json"
        update_checkbox_sidebar(tep_3)

    st.write("---")
    # Xem Header Data
    checkbox_sidebar_4 = st.checkbox(":green[📌 Headers in Data]", key='PL4', value=False)
    if checkbox_sidebar_4:
        tep_4 = "Headers/dict_data.json"
        update_checkbox_sidebar(tep_4)


    st.write("---")
    # Minh hoa vai loai do thi
    checkbox_sidebar_5 = st.checkbox(":red[📌Pandas graph reference]", key='PL5', value=False)
    if checkbox_sidebar_5:
        loai_do_thi = st.radio(
            "Chon loai do thi",
            ["Line plot",
            "Bar plot", 
            "Barh plot", 
            "Histogram", 
            "Box plot", 
            "Area plot", 
            "Pie chart",
            "Scatter plot",
            "Hexbin plot"
            ],
            index=None,
        )
        if loai_do_thi == "Line plot":
            Xem_do_thi_1()
        elif loai_do_thi == "Bar plot":
            Xem_do_thi_2()
        elif loai_do_thi == "Barh plot":
            Xem_do_thi_3()
        elif loai_do_thi == "Histogram":
            Xem_do_thi_4()
        elif loai_do_thi == "Box plot":
            Xem_do_thi_5()
        elif loai_do_thi == "Area plot":
            Xem_do_thi_6()
        elif loai_do_thi == "Pie chart":
            Xem_do_thi_7()
        elif loai_do_thi == "Scatter plot":
            Xem_do_thi_8()
        elif loai_do_thi == "Hexbin plot":
            Xem_do_thi_9()
