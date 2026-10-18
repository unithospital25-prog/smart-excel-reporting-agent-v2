import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Smart Excel Reporting Agent",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Smart Excel Reporting Agent")
st.write("Upload, filter and analyze your Excel data.")

# -----------------------------
# Excel Upload
# -----------------------------
uploaded_file = st.file_uploader(
    "📂 Upload Excel File",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:

    try:
        excel_file = pd.ExcelFile(uploaded_file)

        st.success("✅ Excel file uploaded successfully!")

        # -----------------------------
        # Sheet Selection
        # -----------------------------
        sheet_name = st.selectbox(
            "📄 Select Sheet",
            excel_file.sheet_names
        )

        df = pd.read_excel(
            uploaded_file,
            sheet_name=sheet_name
        )

        df.columns = df.columns.astype(str).str.strip()

        # -----------------------------
        # Search
        # -----------------------------
        st.subheader("🔎 Search & Filter")

        search_text = st.text_input(
            "🔍 Search anything"
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

        # -----------------------------
        # Column Filter
        # -----------------------------
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

        # =====================================================
        # DASHBOARD
        # =====================================================

        st.subheader("📊 Automatic Dashboard")

        total_records = len(filtered_df)

        total_amount = 0
        if "Amount" in filtered_df.columns:
            total_amount = pd.to_numeric(
                filtered_df["Amount"],
                errors="coerce"
            ).fillna(0).sum()

        pending_count = 0
        completed_count = 0

        if "Status" in filtered_df.columns:

            status_data = (
                filtered_df["Status"]
                .astype(str)
                .str.strip()
                .str.lower()
            )

            pending_count = (status_data == "pending").sum()
            completed_count = (status_data == "completed").sum()

        # -----------------------------
        # KPI Cards
        # -----------------------------

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric(
                "📋 Total Records",
                total_records
            )

        with col2:
            st.metric(
                "⏳ Pending",
                pending_count
            )

        with col3:
            st.metric(
                "✅ Completed",
                completed_count
            )

        with col4:
            st.metric(
                "💰 Total Amount",
                f"₹{total_amount:,.0f}"
            )

        # =====================================================
        # BN-WISE SUMMARY
        # =====================================================

        if "BN" in filtered_df.columns:

            st.subheader("🏢 BN-wise Summary")

            bn_summary = (
                filtered_df
                .groupby("BN", dropna=False)
                .agg(
                    Total_Records=("BN", "size")
                )
                .reset_index()
            )

            if "Status" in filtered_df.columns:

                pending_bn = (
                    filtered_df[
                        filtered_df["Status"]
                        .astype(str)
                        .str.lower()
                        .eq("pending")
                    ]
                    .groupby("BN")
                    .size()
                    .rename("Pending")
                )

                completed_bn = (
                    filtered_df[
                        filtered_df["Status"]
                        .astype(str)
                        .str.lower()
                        .eq("completed")
                    ]
                    .groupby("BN")
                    .size()
                    .rename("Completed")
                )

                bn_summary = bn_summary.merge(
                    pending_bn,
                    on="BN",
                    how="left"
                )

                bn_summary = bn_summary.merge(
                    completed_bn,
                    on="BN",
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

            if "Amount" in filtered_df.columns:

                amount_summary = (
                    filtered_df
                    .assign(
                        Amount=pd.to_numeric(
                            filtered_df["Amount"],
                            errors="coerce"
                        ).fillna(0)
                    )
                    .groupby("BN")["Amount"]
                    .sum()
                    .rename("Total Amount")
                )

                bn_summary = bn_summary.merge(
                    amount_summary,
                    on="BN",
                    how="left"
                )

            st.dataframe(
                bn_summary,
                use_container_width=True,
                hide_index=True
            )

        # =====================================================
        # STATUS SUMMARY
        # =====================================================

        if "Status" in filtered_df.columns:

            st.subheader("📌 Status-wise Summary")

            status_summary = (
                filtered_df["Status"]
                .astype(str)
                .str.strip()
                .value_counts()
                .reset_index()
            )

            status_summary.columns = [
                "Status",
                "Count"
            ]

            st.dataframe(
                status_summary,
                use_container_width=True,
                hide_index=True
            )

        # =====================================================
        # EMPLOYEE-WISE SUMMARY
        # =====================================================

        if "Employee Name" in filtered_df.columns:

            st.subheader("👤 Employee-wise Summary")

            employee_summary = (
                filtered_df
                .groupby(
                    ["Employee Name", "BN"],
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

        # =====================================================
        # DATA PREVIEW
        # =====================================================

        st.subheader("📋 Data Preview")

        st.dataframe(
            filtered_df,
            use_container_width=True,
            height=450
        )

        # =====================================================
        # DOWNLOAD
        # =====================================================

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