import streamlit as st
import pandas as pd
import os
from pathlib import Path
from datetime import datetime
import zipfile
import io

# ===================== Core Function =====================
def split_excel_by_columns(df, col_input):
    df = df.fillna("")

    # Nếu nhập bằng chữ cái (A,B,C) thì convert sang tên cột
    col_input = [c.strip() for c in col_input.split(",")]
    excel_cols = list(df.columns)

    col_names = []
    for col in col_input:
        if len(col) == 1 and col.isalpha():  # A,B,C
            idx = ord(col.upper()) - 65
            if idx < len(excel_cols):
                col_names.append(excel_cols[idx])
            else:
                raise ValueError(f"❌ Column {col} không tồn tại")
        else:  # Nếu nhập trực tiếp tên cột
            if col in excel_cols:
                col_names.append(col)
            else:
                raise ValueError(f"❌ Column {col} không tồn tại trong sheet")

    bad_words = ["(All)", "Sum of", "Supplier", "Invoice", "Shipmode"]

    # Tạo buffer zip
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        for keys, group in df.groupby(col_names):
            if isinstance(keys, tuple):
                split_key = "-".join([str(k).strip() for k in keys])
            else:
                split_key = str(keys).strip()

            if not split_key or any(bw.lower() in split_key.lower() for bw in bad_words):
                continue

            for ch in r'\/:*?"<>|':
                split_key = split_key.replace(ch, "_")

            file_name = f"{split_key}-{datetime.today().strftime('%Y%m%d')}.xlsx"

            # Save group vào memory
            output = io.BytesIO()
            group.to_excel(output, index=False)
            zf.writestr(file_name, output.getvalue())

    return zip_buffer


# ===================== Streamlit UI =====================
st.title("📊 Split Excel by Multi Columns")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file:
    sheet_name = st.text_input("Nhập tên Sheet (ví dụ: PivotSheet)", "")
    col_input = st.text_input("Nhập các cột cần tách (vd: A,B hoặc Supplier,Invoice)", "")

    if st.button("🚀 Split Now"):
        try:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name if sheet_name else 0, dtype=str)
            zip_buffer = split_excel_by_columns(df, col_input)

            st.success("✅ Đã tách file thành công!")

            st.download_button(
                label="📥 Download ZIP",
                data=zip_buffer.getvalue(),
                file_name=f"SplitResult-{datetime.today().strftime('%Y%m%d')}.zip",
                mime="application/zip"
            )
        except Exception as e:
            st.error(f"❌ Lỗi: {e}")
