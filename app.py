import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Smart Excel Reporting Agent",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Smart Excel Reporting Agent")
st.write("Upload your Excel file and preview the data.")

uploaded_file = st.file_uploader(
    "📂 Upload Excel File",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:

    try:
        excel_file = pd.ExcelFile(uploaded_file)

        st.success("✅ Excel file uploaded successfully!")

        sheet_name = st.selectbox(
            "📄 Select Sheet",
            excel_file.sheet_names
        )

        df = pd.read_excel(
            uploaded_file,
            sheet_name=sheet_name
        )

        st.subheader("📋 Data Preview")

        st.dataframe(
            df,
            use_container_width=True
        )

        col1, col2, col3 = st.columns(3)

        with col1:
            st.metric("Total Rows", len(df))

        with col2:
            st.metric("Total Columns", len(df.columns))

        with col3:
            st.metric("Total Sheets", len(excel_file.sheet_names))

    except Exception as e:
        st.error(f"❌ Error reading Excel file: {e}")
