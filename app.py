import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(
    page_title="Smart Excel Reporting Agent",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Smart Excel Reporting Agent")
st.write("Upload and analyze your Excel data.")


# =========================
# EXCEL DOWNLOAD FUNCTION
# =========================

def create_excel(dataframe, sheet_name):
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


# =========================
# UPLOAD EXCEL
# =========================

uploaded_file = st.file_uploader(
    "📂 Upload Excel File",
    type=["xlsx", "xls"]
)


if uploaded_file is not None:

    try:

        # =========================
        # READ EXCEL
        # =========================

        excel_file = pd.ExcelFile(
            uploaded_file
        )

        st.success(
            "✅ Excel file uploaded successfully!"
        )


        # =========================
        # SHEET SELECTION
        # =========================

        sheet_name = st.selectbox(
            "📄 Select Sheet",
            excel_file.sheet_names
        )


        df = pd.read_excel(
            uploaded_file,
            sheet_name=sheet_name
        )


        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )


        # =========================
        # SEARCH
        # =========================

        st.subheader(
            "🔎 Search & Filter"
        )


        search_text = st.text_input(
            "🔍 Search UIN, Employee Name, BN, PRAN, Mobile, Status etc."
        )


        filtered_df = df.copy()


        if search_text.strip():

            search_value = (
                search_text
                .strip()
                .casefold()
            )


            search_data = (
                filtered_df
                .fillna("")
                .astype(str)
            )


            mask = search_data.apply(
                lambda column:
                column
                .str
                .casefold()
                .str.contains(
                    re.escape(search_value),
                    regex=True,
                    na=False
                )
            ).any(axis=1)


            filtered_df = (
                filtered_df
                .loc[mask]
                .copy()
            )


            st.info(
                f"🔍 Search result: {len(filtered_df)} record(s) found"
            )


        # =========================
        # FIND BN COLUMN
        # =========================

        bn_col = None


        for col in df.columns:

            if col.lower() in [
                "bn",
                "battalion",
                "unit"
            ]:

                bn_col = col
                break


        # =========================
        # BN FILTER
        # =========================

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

                filtered_df = (
                    filtered_df[
                        filtered_df[bn_col]
                        .astype(str)
                        == selected_bn
                    ]
                )


        # =========================
        # FIND STATUS COLUMN
        # =========================

        status_col = None


        for col in df.columns:

            if col.lower() in [
                "status",
                "case status",
                "case_status"
            ]:

                status_col = col
                break


        # =========================
        # STATUS FILTER
        # =========================

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

                filtered_df = (
                    filtered_df[
                        filtered_df[status_col]
                        .astype(str)
                        == selected_status
                    ]
                )


        # =========================
        # FIND AMOUNT COLUMN
        # =========================

        amount_col = None


        for col in df.columns:

            if col.lower() in [
                "amount",
                "contribution",
                "total amount"
            ]:

                amount_col = col
                break


        # =========================
        # DASHBOARD
        # =========================

        st.subheader(
            "📊 Automatic Dashboard"
        )


        total_records = len(
            filtered_df
        )


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


        # =========================
        # CONTRIBUTION
        # =========================

        total_amount = 0


        if amount_col:

            total_amount = (
                pd.to_numeric(
                    filtered_df[amount_col],
                    errors="coerce"
                )
                .fillna(0)
                .sum()
            )


        # =========================
        # KPI
        # =========================

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


        # =========================
        # BN SUMMARY
        # =========================

        if bn_col:

            st.subheader(
                "🏢 BN-wise Summary"
            )


            bn_summary = (
                filtered_df
                .groupby(bn_col)
                .size()
                .reset_index(
                    name="Total Records"
                )
            )


            if status_col:

                temp = (
                    filtered_df
                    .copy()
                )


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


                bn_summary = (
                    bn_summary
                    .merge(
                        pending,
                        on=bn_col,
                        how="left"
                    )
                    .merge(
                        completed,
                        on=bn_col,
                        how="left"
                    )
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


            st.dataframe(
                bn_summary,
                use_container_width=True,
                hide_index=True
            )


        # =========================
        # CASE STATUS SUMMARY
        # =========================

        if status_col:

            st.subheader(
                "📌 Case Status Summary"
            )


            status_summary = (
                filtered_df[status_col]
                .fillna("Blank")
                .astype(str)
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


        # =========================
        # DATA PREVIEW
        # =========================

        st.subheader(
            "📋 Data Preview"
        )


        st.dataframe(
            filtered_df,
            use_container_width=True,
            height=450
        )


        # =========================
        # BN REPORT
        # =========================

        if bn_col:

            st.subheader(
                "📊 BN Report"
            )


            if st.button(
                "📊 Generate BN Report"
            ):

                report_df = (
                    filtered_df
                    .copy()
                )


                bn_report = (
                    report_df
                    .groupby(bn_col)
                    .size()
                    .reset_index(
                        name="Total Records"
                    )
                )


                if status_col:

                    temp_report = (
                        report_df
                        .copy()
                    )


                    temp_report["_status"] = (
                        temp_report[status_col]
                        .fillna("")
                        .astype(str)
                        .str.strip()
                        .str.casefold()
                    )


                    pending_report = (
                        temp_report[
                            temp_report["_status"]
                            .str.contains(
                                "pending",
                                na=False
                            )
                        ]
                        .groupby(bn_col)
                        .size()
                        .rename("Pending")
                    )


                    completed_report = (
                        temp_report[
                            temp_report["_status"]
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


                    bn_report = (
                        bn_report
                        .merge(
                            pending_report,
                            on=bn_col,
                            how="left"
                        )
                        .merge(
                            completed_report,
                            on=bn_col,
                            how="left"
                        )
                    )


                else:

                    bn_report["Pending"] = 0

                    bn_report["Completed"] = 0


                bn_report["Pending"] = (
                    bn_report["Pending"]
                    .fillna(0)
                    .astype(int)
                )


                bn_report["Completed"] = (
                    bn_report["Completed"]
                    .fillna(0)
                    .astype(int)
                )


                # =========================
                # BN CONTRIBUTION
                # =========================

                if amount_col:

                    contribution_report = (
                        report_df
                        .assign(
                            _amount=pd.to_numeric(
                                report_df[amount_col],
                                errors="coerce"
                            ).fillna(0)
                        )
                        .groupby(bn_col)["_amount"]
                        .sum()
                        .rename(
                            "Total Contribution"
                        )
                    )


                    bn_report = (
                        bn_report
                        .merge(
                            contribution_report,
                            on=bn_col,
                            how="left"
                        )
                    )


                    bn_report[
                        "Total Contribution"
                    ] = (
                        bn_report[
                            "Total Contribution"
                        ]
                        .fillna(0)
                    )


                else:

                    bn_report[
                        "Total Contribution"
                    ] = 0


                st.dataframe(
                    bn_report,
                    use_container_width=True,
                    hide_index=True
                )


                # BN EXCEL

                bn_excel_data = create_excel(
                    bn_report,
                    "BN Report"
                )


                st.download_button(
                    "⬇️ Download BN Report",
                    data=bn_excel_data,
                    file_name="BN_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="bn_report_download"
                )


        # =========================
        # EMPLOYEE-WISE REPORT
        # =========================

        st.subheader(
            "👤 Employee-wise Report"
        )


        employee_report = (
            filtered_df
            .copy()
        )


        if len(employee_report) > 0:

            st.write(
                f"👥 Total Employees: {len(employee_report)}"
            )


            st.dataframe(
                employee_report,
                use_container_width=True,
                hide_index=True
            )


            # EMPLOYEE EXCEL

            employee_excel_data = create_excel(
                employee_report,
                "Employee Report"
            )


            st.download_button(
                "⬇️ Download Employee Excel Report",
                data=employee_excel_data,
                file_name="Employee_Wise_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="employee_report_download"
            )


        else:

            st.info(
                "ℹ️ No employee records found."
            )


        # =========================
        # FILTERED REPORT
        # =========================

        csv_data = (
            filtered_df
            .to_csv(
                index=False
            )
            .encode("utf-8")
        )


        st.download_button(
            "⬇️ Download Filtered Report",
            data=csv_data,
            file_name="filtered_report.csv",
            mime="text/csv",
            key="filtered_report_download"
        )


    except Exception as e:

        st.error(
            f"❌ Error: {e}"
        )