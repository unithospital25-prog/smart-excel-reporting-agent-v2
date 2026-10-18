import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Smart Excel Reporting Agent",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Smart Excel Reporting Agent")
st.write("Upload your Excel file and analyze the data.")

# Upload Excel
uploaded_file = st.file_uploader(
    "📂 Upload Excel File",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:

    try:
        # Read Excel
        excel_file = pd.ExcelFile(uploaded_file)

        st.success("✅ Excel file uploaded successfully!")

        # Sheet selection
        sheet_name = st.selectbox(
            "📄 Select Sheet",
            excel_file.sheet_names
        )

        df = pd.read_excel(
            uploaded_file,
            sheet_name=sheet_name
        )

        # Clean column names
        df.columns = df.columns.astype(str).str.strip()

        st.subheader("🔎 Search & Filter")

        # Search box
        search_text = st.text_input(
            "🔍 Search anything in the Excel data"
        )

        filtered_df = df.copy()

        if search_text:
            mask = filtered_df.astype(str).apply(
                lambda row: row.str.contains(
                    search_text,
                    case=False,
                    na=False
                ).any(),
                axis=1
            )

            filtered_df = filtered_df[mask]

        # Column filter
        filter_column = st.selectbox(
            "📌 Filter by Column",
            ["No Filter"] + list(df.columns)
        )

        if filter_column != "No Filter":

            values = (
                filtered_df[filter_column]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )

            selected_value = st.selectbox(
                f"Select {filter_column}",
                ["All"] + sorted(values)
            )

            if selected_value != "All":
                filtered_df = filtered_df[
                    filtered_df[filter_column]
                    .astype(str)
                    == selected_value
                ]

        # Summary
        st.subheader("📊 Automatic Summary")

        col1, col2, col3 = st.columns(3)

        with col1:
            st.metric(
                "Total Rows",
                len(df)
            )

        with col2:
            st.metric(
                "Filtered Rows",
                len(filtered_df)
            )

        with col3:
            st.metric(
                "Total Columns",
                len(df.columns)
            )

        # Data preview
        st.subheader("📋 Data Preview")

        st.dataframe(
            filtered_df,
            use_container_width=True,
            height=500
        )

        # Download filtered data
        csv_data = filtered_df.to_csv(
            index=False
        ).encode("utf-8")

        st.download_button(
            "⬇️ Download Filtered Data",
            data=csv_data,
            file_name="filtered_report.csv",
            mime="text/csv"
        )

    except Exception as e:

        st.error(
            f"❌ Error reading Excel file: {e}"
        )