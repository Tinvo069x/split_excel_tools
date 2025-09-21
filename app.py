import streamlit as st
import pandas as pd
import io, zipfile
from datetime import datetime

# ===================== Core Function =====================
def split_excel_by_columns(df, selected_cols):
    df = df.fillna("")

    bad_words = ["(All)", "Sum of", "Supplier", "Invoice", "Shipmode"]

    # Tạo buffer zip
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        for keys, group in df.groupby(selected_cols):
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
    # Lấy danh sách sheet
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = st.selectbox("Chọn sheet:", xls.sheet_names)

    # Nhập dòng header
    header_row = st.number_input("Chọn dòng header (ví dụ: 1,2,3...)", min_value=1, value=1, step=1)

    if st.button("🔍 Xem trước dữ liệu"):
        try:
            df_preview = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row-1, nrows=10)
            st.dataframe(df_preview)
        except Exception as e:
            st.error(f"Lỗi khi đọc file: {e}")

    try:
        # Đọc full data theo header row
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row-1, dtype=str)

        # Multi-select để chọn cột split
        selected_cols = st.multiselect("Chọn các cột để tách file:", df.columns.tolist())

        if st.button("🚀 Split Now"):
            if not selected_cols:
                st.warning("⚠️ Vui lòng chọn ít nhất 1 cột để split")
            else:
                zip_buffer = split_excel_by_columns(df, selected_cols)
                st.success("✅ Đã tách file thành công!")

                st.download_button(
                    label="📥 Download ZIP",
                    data=zip_buffer.getvalue(),
                    file_name=f"SplitResult-{datetime.today().strftime('%Y%m%d')}.zip",
                    mime="application/zip"
                )
    except Exception as e:
        st.error(f"❌ Lỗi khi xử lý: {e}")
