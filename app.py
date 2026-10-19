import streamlit as st
import pandas as pd
import re
from io import BytesIO

# =========================================================
# PAGE CONFIGURATION
# =========================================================

st.set_page_config(
    page_title="Smart Excel Reporting Agent",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Smart Excel Reporting Agent")
st.write("Upload and analyze your Excel data.")

# =========================================================
# HELPER FUNCTION - EXCEL DOWNLOAD
# =========================================================

def create_excel_download(dataframe, sheet_name="Report"):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        dataframe.to_excel(
            writer,
            index=False,
            sheet_name=sheet_name
        )

    output.seek(0)

    return output.getvalue()


# =========================================================
# UPLOAD EXCEL
# =========================================================

uploaded_file = st.file_uploader(
    "📂 Upload Excel File",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:

    try:

        # =================================================
        # READ EXCEL
        # =================================================

        excel_file = pd.ExcelFile(uploaded_file)

        st.success("✅ Excel file uploaded successfully!")

        # =================================================
        # SELECT SHEET
        # =================================================

        sheet_name = st.selectbox(
            "📄 Select Sheet",
            excel_file.sheet_names
        )

        df = pd.read_excel(
            uploaded_file,
            sheet_name=sheet_name
        )

        # Clean column names

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

        # =================================================
        # SEARCH & FILTER
        # =================================================

        st.subheader("🔎 Search & Filter")

        search_text = st.text_input(
            "🔍 Search UIN, Employee Name, BN, PRAN, Mobile, Status etc."
        )

        filtered_df = df.copy()

        # =================================================
        # SEARCH
        # =================================================

        if search_text.strip():

            search_value = search_text.strip().casefold()

            search_data = (
                filtered_df
                .fillna("")
                .astype(str)
            )

            mask = search_data.apply(
                lambda column:
                column.str.casefold()
                .str.contains(
                    re.escape(search_value),
                    regex=True,
                    na=False
                )
            ).any(axis=1)

            filtered_df = filtered_df.loc[mask].copy()

            st.info(
                f"🔍 Search result: {len(filtered_df)} record(s) found"
            )

        # =================================================
        # FIND BN COLUMN
        # =================================================

        bn_col = None

        for col in df.columns:

            if col.lower() in [
                "bn",
                "battalion",
                "unit"
            ]:

                bn_col = col
                break

        # =================================================
        # BN FILTER
        # =================================================

        if bn_col:

            bn_values = sorted(
                df[bn_col]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )

            selected_bn = st.selectbox(
                "🏢 Select BN",
                ["All"] + bn_values
            )

            if selected_bn != "All":

                filtered_df = filtered_df[
                    filtered_df[bn_col]
                    .astype(str)
                    == selected_bn
                ]

        # =================================================
        # FIND STATUS COLUMN
        # =================================================

        status_col = None

        for col in df.columns:

            if col.lower() in [
                "status",
                "case status",
                "case_status"
            ]:

                status_col = col
                break

        # =================================================
        # STATUS FILTER
        # =================================================

        if status_col:

            status_values = sorted(
                df[status_col]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )

            selected_status = st.selectbox(
                "📌 Case Status",
                ["All"] + status_values
            )

            if selected_status != "All":

                filtered_df = filtered_df[
                    filtered_df[status_col]
                    .astype(str)
                    == selected_status
                ]

        # =================================================
        # FIND CONTRIBUTION COLUMN
        # =================================================

        amount_col = None

        for col in df.columns:

            if col.lower() in [
                "amount",
                "contribution",
                "total amount"
            ]:

                amount_col = col
                break

        # =================================================
        # AUTOMATIC DASHBOARD
        # =================================================

        st.subheader("📊 Automatic Dashboard")

        total_records = len(filtered_df)

        pending_count = 0
        completed_count = 0

        if status_col:

            status_data = (
                filtered_df[status_col]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.casefold()
            )

            pending_count = (
                status_data
                .str.contains(
                    "pending",
                    na=False
                )
                .sum()
            )

            completed_count = (
                status_data
                .str.contains(
                    "complete|completed|done",
                    regex=True,
                    na=False
                )
                .sum()
            )

        # =================================================
        # TOTAL CONTRIBUTION
        # =================================================

        total_amount = 0

        if amount_col:

            total_amount = pd.to_numeric(
                filtered_df[amount_col],
                errors="coerce"
            ).fillna(0).sum()

        # =================================================
        # KPI
        # =================================================

        c1, c2, c3, c4 = st.columns(4)

        with c1:

            st.metric(
                "📋 Total Records",
                total_records
            )

        with c2:

            st.metric(
                "⏳ Pending",
                pending_count
            )

        with c3:

            st.metric(
                "✅ Completed",
                completed_count
            )

        with c4:

            st.metric(
                "💰 Contribution",
                f"₹{total_amount:,.0f}"
            )

        # =================================================
        # BN-WISE SUMMARY
        # =================================================

        if bn_col:

            st.subheader("🏢 BN-wise Summary")

            bn_summary = (
                filtered_df
                .groupby(bn_col)
                .size()
                .reset_index(
                    name="Total Records"
                )
            )

            if status_col:

                temp = filtered_df.copy()

                temp["_status"] = (
                    temp[status_col]
                    .fillna("")
                    .astype(str)
                    .str.casefold()
                )

                pending = (
                    temp[
                        temp["_status"]
                        .str.contains(
                            "pending",
                            na