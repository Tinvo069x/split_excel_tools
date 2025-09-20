import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO

st.title("📊 Excel Splitter App (Streamlit Pro)")

uploaded_file = st.file_uploader("📂 Upload file Excel", type=["xlsx", "xls"])

if uploaded_file:
    # Lấy danh sách sheet
    excel_file = pd.ExcelFile(uploaded_file)
    sheet_name = st.selectbox("📑 Chọn sheet", excel_file.sheet_names)

    # Nhập header row
    header_row = st.number_input("Header row (ví dụ: 1 = dòng đầu)", min_value=1, value=1, step=1)

    # Preview vài dòng
    preview_df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row-1, nrows=5)
    st.write("👀 Preview dữ liệu:", preview_df)

    # Chọn cột để split
    split_col = st.selectbox("🔑 Chọn cột để split", preview_df.columns)

    if st.button("🚀 Run Split"):
        try:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row-1)

            if split_col not in df.columns:
                st.error(f"❌ Không tìm thấy cột '{split_col}'")
            else:
                # Nén vào zip
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zipf:
                    for val in df[split_col].dropna().unique():
                        df_split = df[df[split_col] == val]
                        safe_val = str(val).replace("/", "_").replace("\\", "_").replace(":", "_")
                        output_filename = f"{split_col}_{safe_val}.xlsx"

                        excel_buffer = BytesIO()
                        df_split.to_excel(excel_buffer, index=False)
                        zipf.writestr(output_filename, excel_buffer.getvalue())

                zip_buffer.seek(0)
                st.success("✅ Đã tách file xong!")

                st.download_button(
                    label="⬇️ Download ZIP",
                    data=zip_buffer,
                    file_name="output_split.zip",
                    mime="application/zip"
                )
        except Exception as e:
            st.error(f"⚠️ Lỗi: {str(e)}")
