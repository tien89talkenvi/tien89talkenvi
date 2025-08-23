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
#-------------------
def ThucThiPhan_I():
    pass

#-------------------
#@st.cache_data
def ThucThiPhan_II(sheet2,sheet1):
    if sheet2 and sheet1:
        # Đọc sheet2.txt (Application Specific Data)
        df1 = pd.read_csv(sheet2, sep="\t", dtype=str)
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

# I. TAI FILES DU LIEU TXT TU SMARTS -------------------------
st.subheader('✅ Download the data', divider=True)
ThucThiPhan_I()

# II XU LI VA TAO DATA MOI ------------------------------------
st.subheader('✅ Analyze the new data and creat Sheet', divider=True)
sheet1=None
sheet2=None
checkboxII = st.checkbox("📌:blue[Lấy 2 files TXT (Application Specific Data và Ad Hoc Reports Data) để xử lí và tạo dữ liệu mới]", key='P2', value=False)
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
            ThucThiPhan_II(sheet2, sheet1)

    elif laydatafrom==":blue[From Local]":
        sheet2 = st.file_uploader("Upload sheet2.txt :red[(Application Specific Data)]", type=["txt"])
        sheet1 = st.file_uploader("Upload sheet1.txt :red[(Ad Hoc Reports Data)]", type=["txt"])
        if not (sheet2 and sheet1):
            tbaodong1 = st.empty()
            tbaodong1.write('Chưa có dữ liệu')
        else:
            tbaodong1 = st.empty()
            tbaodong1.success(f"Đã lấy dữ liệu cần : {sheet2.name}, {sheet1.name}.Chờ xử lí.")
            ThucThiPhan_II(sheet2, sheet1)
 
# IV DO THI HOA DU LIEU -------------------------
st.subheader('✅ Visualize the data', divider=True)
ThucThiPhan_III()
#⏱️
