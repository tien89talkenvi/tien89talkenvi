import streamlit as st
import pandas as pd
import io, time
from datetime import datetime


start_time = time.time()  # Bắt đầu tính giờ

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

# ---------------------------
# Streamlit App
# ---------------------------
st.title("📊 Xử lý dữ liệu SMARTS")

# Upload sheet2.txt & sheet1.txt
sheet2 = st.file_uploader("Upload sheet2.txt :red[(Application Specific Data)]", type=["txt"])
sheet1 = st.file_uploader("Upload sheet1.txt :red[(Ad Hoc Reports Data)]", type=["txt"])

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
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_final.to_excel(writer, sheet_name="Data", index=False)

        # Lấy workbook và worksheet
        workbook  = writer.book
        worksheet = writer.sheets["Data"]

        # Freeze Panes: cố định dòng đầu và 4 cột đầu
        worksheet.freeze_panes(1, 4)  
        # (1,4) nghĩa là khóa hàng trên dòng 2 và cột trước cột E

    st.success("✅ Dữ liệu đã xử lý xong")

    # Nút tải về
    st.download_button(
        label="📥 Tải Excel kết quả",
        data=output.getvalue(),
        file_name="SMARTS_Data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


    # Hiển thị dataframe
    st.dataframe(df_final)

    end_time = time.time()
    elapsed_time = (end_time - start_time)/60

    st.write(f"⏳ It took : {elapsed_time:.2f} minutes")