import streamlit as st
import pandas as pd
import re

# ==============================
# PAGE SETTINGS
# ==============================

st.set_page_config(
    page_title="Smart Excel Reporting Agent",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Smart Excel Reporting Agent")
st.write("Upload, filter and analyze your Excel data.")

# ==============================
# EXCEL UPLOAD
# ==============================

uploaded_file = st.file_uploader(
    "📂 Upload Excel File",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:

    try:

        excel_file = pd.ExcelFile(uploaded_file)

        st.success("✅ Excel file uploaded successfully!")

        # ==============================
        # SHEET SELECTION
        # ==============================

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

        # ==============================
# SMART SEARCH
# ==============================

st.subheader("🔎 Search & Filter")

search_text = st.text_input(
    "🔍 Search UIN, Employee Name, BN, PRAN, Mobile, Status etc."
)

filtered_df = df.copy()

if search_text and search_text.strip():

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
    ) ==============================
        # COLUMN FILTER
        # ==============================

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

        # ==============================
        # COLUMN DETECTION
        # ==============================

        bn_col = None
        name_col = None
        status_col = None
        amount_col = None

        # Battalion
        for col in df.columns:
            if col.lower() in [
                "bn",
                "battalion",
                "unit"
            ]:
                bn_col = col
                break

        # Name
        for col in df.columns:
            if col.lower() in [
                "name",
                "employee name",
                "employee_name"
            ]:
                name_col = col
                break

        # Status
        for col in df.columns:
            if col.lower() in [
                "status",
                "case status",
                "case_status"
            ]:
                status_col = col
                break

        # Amount / Contribution
        for col in df.columns:
            if col.lower() in [
                "amount",
                "contribution",
                "total amount"
            ]:
                amount_col = col
                break

        # ==============================
        # DASHBOARD
        # ==============================

        st.subheader("📊 Automatic Dashboard")

        total_records = len(filtered_df)

        # Amount
        total_amount = 0

        if amount_col:

            total_amount = pd.to_numeric(
                filtered_df[amount_col],
                errors="coerce"
            ).fillna(0).sum()

        # Status
        pending_count = 0
        completed_count = 0

        if status_col:

            status_values = (
                filtered_df[status_col]
                .astype(str)
                .str.strip()
                .str.lower()
            )

            pending_count = (
                status_values
                .str.contains(
                    "pending",
                    na=False
                )
                .sum()
            )

            completed_count = (
                status_values
                .str.contains(
                    "complete|completed|done",
                    regex=True,
                    na=False
                )
                .sum()
            )

        # ==============================
        # KPI CARDS
        # ==============================

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

        # ==============================
        # BN-WISE SUMMARY
        # ==============================

        if bn_col:

            st.subheader("🏢 Battalion-wise Summary")

            bn_summary = (
                filtered_df
                .groupby(
                    bn_col,
                    dropna=False
                )
                .size()
                .reset_index(
                    name="Total Records"
                )
            )

            # Status counts
            if status_col:

                temp = filtered_df.copy()

                temp["_status"] = (
                    temp[status_col]
                    .astype(str)
                    .str.strip()
                    .str.lower()
                )

                pending = (
                    temp[
                        temp["_status"]
                        .str.contains(
                            "pending",
                            na=False
                        )
                    ]
                    .groupby(bn_col)
                    .size()
                    .rename("Pending")
                )

                completed = (
                    temp[
                        temp["_status"]
                        .str.contains(
                            "complete|completed|done",
                            regex=True,
                            na=False
                        )
                    ]
                    .groupby(bn_col)
                    .size()
                    .rename("Completed")
                )

                bn_summary = bn_summary.merge(
                    pending,
                    on=bn_col,
                    how="left"
                )

                bn_summary = bn_summary.merge(
                    completed,
                    on=bn_col,
                    how="left"
                )

                bn_summary["Pending"] = (
                    bn_summary["Pending"]
                    .fillna(0)
                    .astype(int)
                )

                bn_summary["Completed"] = (
                    bn_summary["Completed"]
                    .fillna(0)
                    .astype(int)
                )

            # Contribution
            if amount_col:

                contribution = (
                    filtered_df
                    .assign(
                        _amount=pd.to_numeric(
                            filtered_df[amount_col],
                            errors="coerce"
                        ).fillna(0)
                    )
                    .groupby(bn_col)["_amount"]
                    .sum()
                    .rename("Contribution")
                )

                bn_summary = bn_summary.merge(
                    contribution,
                    on=bn_col,
                    how="left"
                )

            st.dataframe(
                bn_summary,
                use_container_width=True,
                hide_index=True
            )

        # ==============================
        # CASE STATUS SUMMARY
        # ==============================

        if status_col:

            st.subheader("📌 Case Status Summary")

            status_summary = (
                filtered_df[status_col]
                .astype(str)
                .str.strip()
                .value_counts()
                .reset_index()
            )

            status_summary.columns = [
                "Case Status",
                "Total"
            ]

            st.dataframe(
                status_summary,
                use_container_width=True,
                hide_index=True
            )

        # ==============================
        # NOMINEE SUMMARY
        # ==============================

        nominee_col = None

        for col in df.columns:
            if col.lower() in [
                "nominee",
                "nomination status"
            ]:
                nominee_col = col
                break

        if nominee_col:

            st.subheader("👨‍👩‍👧 Nominee Summary")

            nominee_summary = (
                filtered_df[nominee_col]
                .astype(str)
                .str.strip()
                .replace(
                    "",
                    "Blank"
                )
                .value_counts()
                .reset_index()
            )

            nominee_summary.columns = [
                "Nominee",
                "Total"
            ]

            st.dataframe(
                nominee_summary,
                use_container_width=True,
                hide_index=True
            )

        # ==============================
        # EMPLOYEE SUMMARY
        # ==============================

        if name_col and bn_col:

            st.subheader("👤 Employee-wise Summary")

            employee_summary = (
                filtered_df[
                    [name_col, bn_col]
                ]
                .copy()
            )

            employee_summary = (
                employee_summary
                .groupby(
                    [bn_col, name_col],
                    dropna=False
                )
                .size()
                .reset_index(
                    name="Records"
                )
            )

            st.dataframe(
                employee_summary,
                use_container_width=True,
                hide_index=True
            )

        # ==============================
        # DATA PREVIEW
        # ==============================

        st.subheader("📋 Data Preview")

        st.dataframe(
            filtered_df,
            use_container_width=True,
            height=450
        )

        # ==============================
        # DOWNLOAD
        # ==============================

        csv_data = filtered_df.to_csv(
            index=False
        ).encode("utf-8")

        st.download_button(
            "⬇️ Download Filtered Report",
            data=csv_data,
            file_name="filtered_report.csv",
            mime="text/csv"
        )

    except Exception as e:

        st.error(
            f"❌ Error reading Excel file: {e}"
        )