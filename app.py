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


def make_excel(dataframe, sheet_name="Report"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output.getvalue()


def find_column(columns, possible_names):
    for col in columns:
        if str(col).strip().lower() in possible_names:
            return col
    return None


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

        df.columns = df.columns.astype(str).str.strip()

        # Column detection
        bn_col = find_column(
            df.columns,
            {"bn", "battalion", "unit"}
        )

        status_col = find_column(
            df.columns,
            {"status", "case status", "case_status"}
        )

        amount_col = find_column(
            df.columns,
            {"amount", "contribution", "total amount"}
        )

        uin_col = find_column(
            df.columns,
            {
                "uin",
                "employee number",
                "employee no",
                "employee id",
                "employee_id"
            }
        )

        employee_name_col = find_column(
            df.columns,
            {"employee name", "name", "employee_name"}
        )

        pran_col = find_column(
            df.columns,
            {"pran", "pran no", "pran number", "pran_number"}
        )

        nomination_col = find_column(
            df.columns,
            {
                "nomination",
                "nomination status",
                "nominee",
                "nominee status"
            }
        )

        # Search
        st.subheader("🔎 Search & Filter")

        search_text = st.text_input(
            "🔍 Search UIN, Employee Name, BN, PRAN, Mobile, Status etc."
        )

        filtered_df = df.copy()

        if search_text.strip():
            search_value = search_text.strip().casefold()

            search_data = (
                filtered_df
                .fillna("")
                .astype(str)
            )

            mask = search_data.apply(
                lambda col: col.str.casefold().str.contains(
                    re.escape(search_value),
                    regex=True,
                    na=False
                )
            ).any(axis=1)

            filtered_df = filtered_df.loc[mask].copy()

            st.info(
                f"🔍 Search result: {len(filtered_df)} record(s) found"
            )

        # BN filter
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
                    filtered_df[bn_col].astype(str) == selected_bn
                ]

        # Status filter
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
                    filtered_df[status_col].astype(str) == selected_status
                ]

        # Dashboard
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

            pending_count = status_data.str.contains(
                "pending",
                regex=False,
                na=False
            ).sum()

            completed_count = status_data.str.contains(
                "complete|completed|done",
                regex=True,
                na=False
            ).sum()

        total_amount = 0

        if amount_col:
            total_amount = pd.to_numeric(
                filtered_df[amount_col],
                errors="coerce"
            ).fillna(0).sum()

        c1, c2, c3, c4 = st.columns(4)

        with c1:
            st.metric("📋 Total Records", total_records)

        with c2:
            st.metric("⏳ Pending", pending_count)

        with c3:
            st.metric("✅ Completed", completed_count)

        with c4:
            st.metric("💰 Contribution", f"₹{total_amount:,.0f}")

        # BN summary
        if bn_col:
            st.subheader("🏢 BN-wise Summary")

            bn_summary = (
                filtered_df
                .groupby(bn_col)
                .size()
                .reset_index(name="Total Records")
            )

            if status_col:
                temp = filtered_df.copy()

                temp["_status"] = (
                    temp[status_col]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .str.casefold()
                )

                pending = (
                    temp[
                        temp["_status"].str.contains(
                            "pending",
                            regex=False,
                            na=False
                        )
                    ]
                    .groupby(bn_col)
                    .size()
                    .rename("Pending")
                )

                completed = (
                    temp[
                        temp["_status"].str.contains(
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
            else:
                bn_summary["Pending"] = 0
                bn_summary["Completed"] = 0

            bn_summary["Pending"] = (
                bn_summary["Pending"].fillna(0).astype(int)
            )

            bn_summary["Completed"] = (
                bn_summary["Completed"].fillna(0).astype(int)
            )

            if amount_col:
                bn_amount = (
                    filtered_df.assign(
                        _amount=pd.to_numeric(
                            filtered_df[amount_col],
                            errors="coerce"
                        ).fillna(0)
                    )
                    .groupby(bn_col)["_amount"]
                    .sum()
                    .rename("Total Contribution")
                )

                bn_summary = bn_summary.merge(
                    bn_amount,
                    on=bn_col,
                    how="left"
                )
            else:
                bn_summary["Total Contribution"] = 0

            bn_summary["Total Contribution"] = (
                bn_summary["Total Contribution"].fillna(0)
            )

            st.dataframe(
                bn_summary,
                use_container_width=True,
                hide_index=True
            )

        # Case status summary
        if status_col:
            st.subheader("📌 Case Status Summary")

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

        # BN Report
        if bn_col:
            st.subheader("📊 BN Report")

            if st.button("📊 Generate BN Report"):
                report_df = filtered_df.copy()

                bn_report = (
                    report_df
                    .groupby(bn_col)
                    .size()
                    .reset_index(name="Total Records")
                )

                if status_col:
                    temp_report = report_df.copy()

                    temp_report["_status"] = (
                        temp_report[status_col]
                        .fillna("")
                        .astype(str)
                        .str.strip()
                        .str.casefold()
                    )

                    pending_report = (
                        temp_report[
                            temp_report["_status"].str.contains(
                                "pending",
                                regex=False,
                                na=False
                            )
                        ]
                        .groupby(bn_col)
                        .size()
                        .rename("Pending")
                    )

                    completed_report = (
                        temp_report[
                            temp_report["_status"].str.contains(
                                "complete|completed|done",
                                regex=True,
                                na=False
                            )
                        ]
                        .groupby(bn_col)
                        .size()
                        .rename("Completed")
                    )

                    bn_report = bn_report.merge(
                        pending_report,
                        on=bn_col,
                        how="left"
                    )

                    bn_report = bn_report.merge(
                        completed_report,
                        on=bn_col,
                        how="left"
                    )
                else:
                    bn_report["Pending"] = 0
                    bn_report["Completed"] = 0

                bn_report["Pending"] = (
                    bn_report["Pending"].fillna(0).astype(int)
                )

                bn_report["Completed"] = (
                    bn_report["Completed"].fillna(0).astype(int)
                )

                if amount_col:
                    contribution_report = (
                        report_df.assign(
                            _amount=pd.to_numeric(
                                report_df[amount_col],
                                errors="coerce"
                            ).fillna(0)
                        )
                        .groupby(bn_col)["_amount"]
                        .sum()
                        .rename("Total Contribution")
                    )

                    bn_report = bn_report.merge(
                        contribution_report,
                        on=bn_col,
                        how="left"
                    )
                else:
                    bn_report["Total Contribution"] = 0

                bn_report["Total Contribution"] = (
                    bn_report["Total Contribution"].fillna(0)
                )

                st.dataframe(
                    bn_report,
                    use_container_width=True,
                    hide_index=True
                )

                bn_excel = make_excel(
                    bn_report,
                    "BN Report"
                )

                st.download_button(
                    "⬇️ Download BN Report",
                    data=bn_excel,
                    file_name="BN_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_bn_report"
                )

        # Smart employee report
        st.subheader("👤 Smart Employee-wise Report")

        employee_report = filtered_df.copy()

        if uin_col:
            unique_employee_count = (
                employee_report[uin_col]
                .dropna()
                .astype(str)
                .str.strip()
                .replace("", pd.NA)
                .dropna()
                .nunique()
            )
        elif employee_name_col:
            unique_employee_count = (
                employee_report[employee_name_col]
                .dropna()
                .astype(str)
                .str.strip()
                .replace("", pd.NA)
                .dropna()
                .nunique()
            )
        else:
            unique_employee_count = len(employee_report)

        st.write(
            f"👥 Total Unique Employees: **{unique_employee_count}**"
        )

        if len(employee_report) > 0:
            st.dataframe(
                employee_report,
                use_container_width=True,
                hide_index=True,
                height=400
            )

            employee_excel = make_excel(
                employee_report,
                "Employee Report"
            )

            st.download_button(
                "⬇️ Download Employee Excel Report",
                data=employee_excel,
                file_name="Employee_Wise_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_employee_report"
            )
        else:
            st.info("ℹ️ No employee records found.")

        # Duplicate UIN
        if uin_col:
            st.subheader("⚠️ Duplicate UIN Report")

            uin_data = (
                employee_report[uin_col]
                .fillna("")
                .astype(str)
                .str.strip()
            )

            duplicate_mask = (
                uin_data.ne("")
                & uin_data.duplicated(keep=False)
            )

            duplicate_uin_report = (
                employee_report[duplicate_mask].copy()
            )

            if len(duplicate_uin_report) > 0:
                st.warning(
                    f"⚠️ Duplicate UIN Records: "
                    f"{len(duplicate_uin_report)}"
                )

                st.dataframe(
                    duplicate_uin_report,
                    use_container_width=True,
                    hide_index=True
                )

                duplicate_excel = make_excel(
                    duplicate_uin_report,
                    "Duplicate UIN"
                )

                st.download_button(
                    "⬇️ Download Duplicate UIN Report",
                    data=duplicate_excel,
                    file_name="Duplicate_UIN_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_duplicate_uin"
                )
            else:
                st.success("✅ No Duplicate UIN found.")

        # Missing PRAN
        if pran_col:
            st.subheader("❌ Missing PRAN Report")

            pran_data = (
                employee_report[pran_col]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.casefold()
            )

            missing_pran_mask = pran_data.isin(
                [
                    "",
                    "none",
                    "nan",
                    "null",
                    "na",
                    "n/a",
                    "-"
                ]
            )

            missing_pran_report = (
                employee_report[missing_pran_mask].copy()
            )

            if len(missing_pran_report) > 0:
                st.warning(
                    f"❌ Employees without PRAN: "
                    f"{len(missing_pran_report)}"
                )

                st.dataframe(
                    missing_pran_report,
                    use_container_width=True,
                    hide_index=True
                )

                missing_pran_excel = make_excel(
                    missing_pran_report,
                    "Missing PRAN"
                )

                st.download_button(
                    "⬇️ Download Missing PRAN Report",
                    data=missing_pran_excel,
                    file_name="Missing_PRAN_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_missing_pran"
                )
            else:
                st.success("✅ All employees have PRAN.")

        # Missing nomination
        if nomination_col:
            st.subheader("❌ Missing Nomination Report")

            nomination_data = (
                employee_report[nomination_col]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.casefold()
            )

            missing_nomination_mask = nomination_data.isin(
                [
                    "",
                    "none",
                    "nan",
                    "null",
                    "na",
                    "n/a",
                    "no",
                    "pending",
                    "not available",
                    "-"
                ]
            )

            missing_nomination_report = (
                employee_report[missing_nomination_mask].copy()
            )

            if len(missing_nomination_report) > 0:
                st.warning(
                    f"❌ Employees with missing nomination: "
                    f"{len(missing_nomination_report)}"
                )

                st.dataframe(
                    missing_nomination_report,
                    use_container_width=True,
                    hide_index=True
                )

                nomination_excel = make_excel(
                    missing_nomination_report,
                    "Missing Nomination"
                )

                st.download_button(
                    "⬇️ Download Missing Nomination Report",
                    data=nomination_excel,
                    file_name="Missing_Nomination_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_missing_nomination"
                )
            else:
                st.success("✅ All employees have nomination.")

        # =====================================================
        # COMPLETE MASTER REPORT
        # =====================================================

        st.subheader("📦 Complete Master Report")
        st.write(
            "एक button से सभी महत्वपूर्ण reports एक ही Excel file में तैयार करें।"
        )

        if st.button("📦 Generate Complete Master Report", key="generate_master_report"):

            master_output = BytesIO()

            with pd.ExcelWriter(master_output, engine="openpyxl") as writer:

                # -------------------------
                # 1. Summary Sheet
                # -------------------------
                summary_data = pd.DataFrame({
                    "Particular": [
                        "Total Records",
                        "Pending",
                        "Completed",
                        "Total Contribution",
                        "Unique Employees",
                        "Missing PRAN",
                        "Missing Nomination",
                        "Duplicate UIN Records"
                    ],
                    "Value": [
                        total_records,
                        pending_count,
                        completed_count,
                        total_amount,
                        unique_employee_count,
                        len(missing_pran_report) if pran_col else 0,
                        len(missing_nomination_report) if nomination_col else 0,
                        len(duplicate_uin_report) if uin_col else 0
                    ]
                })
                summary_data.to_excel(
                    writer,
                    index=False,
                    sheet_name="Summary"
                )

                # -------------------------
                # 2. BN Summary
                # -------------------------
                if bn_col:
                    bn_master = filtered_df.groupby(bn_col).size().reset_index(
                        name="Total Records"
                    )

                    if status_col:
                        master_temp = filtered_df.copy()
                        master_temp["_status"] = (
                            master_temp[status_col]
                            .fillna("")
                            .astype(str)
                            .str.strip()
                            .str.casefold()
                        )

                        master_pending = (
                            master_temp[
                                master_temp["_status"].str.contains(
                                    "pending",
                                    regex=False,
                                    na=False
                                )
                            ]
                            .groupby(bn_col)
                            .size()
                            .rename("Pending")
                        )

                        master_completed = (
                            master_temp[
                                master_temp["_status"].str.contains(
                                    "complete|completed|done",
                                    regex=True,
                                    na=False
                                )
                            ]
                            .groupby(bn_col)
                            .size()
                            .rename("Completed")
                        )

                        bn_master = bn_master.merge(
                            master_pending,
                            on=bn_col,
                            how="left"
                        ).merge(
                            master_completed,
                            on=bn_col,
                            how="left"
                        )
                    else:
                        bn_master["Pending"] = 0
                        bn_master["Completed"] = 0

                    bn_master["Pending"] = bn_master["Pending"].fillna(0).astype(int)
                    bn_master["Completed"] = bn_master["Completed"].fillna(0).astype(int)

                    if amount_col:
                        bn_amount_master = (
                            filtered_df.assign(
                                _amount=pd.to_numeric(
                                    filtered_df[amount_col],
                                    errors="coerce"
                                ).fillna(0)
                            )
                            .groupby(bn_col)["_amount"]
                            .sum()
                            .rename("Total Contribution")
                        )

                        bn_master = bn_master.merge(
                            bn_amount_master,
                            on=bn_col,
                            how="left"
                        )
                    else:
                        bn_master["Total Contribution"] = 0

                    bn_master["Total Contribution"] = (
                        bn_master["Total Contribution"].fillna(0)
                    )

                    bn_master.to_excel(
                        writer,
                        index=False,
                        sheet_name="BN Summary"
                    )

                # -------------------------
                # 3. Employee Report
                # -------------------------
                employee_report.to_excel(
                    writer,
                    index=False,
                    sheet_name="Employee Report"
                )

                # -------------------------
                # 4. Missing PRAN
                # -------------------------
                if pran_col:
                    missing_pran_report.to_excel(
                        writer,
                        index=False,
                        sheet_name="Missing PRAN"
                    )
                else:
                    pd.DataFrame({"Message": ["PRAN column not found"]}).to_excel(
                        writer,
                        index=False,
                        sheet_name="Missing PRAN"
                    )

                # -------------------------
                # 5. Missing Nomination
                # -------------------------
                if nomination_col:
                    missing_nomination_report.to_excel(
                        writer,
                        index=False,
                        sheet_name="Missing Nomination"
                    )
                else:
                    pd.DataFrame({"Message": ["Nomination column not found"]}).to_excel(
                        writer,
                        index=False,
                        sheet_name="Missing Nomination"
                    )

                # -------------------------
                # 6. Duplicate UIN
                # -------------------------
                if uin_col:
                    duplicate_uin_report.to_excel(
                        writer,
                        index=False,
                        sheet_name="Duplicate UIN"
                    )
                else:
                    pd.DataFrame({"Message": ["UIN column not found"]}).to_excel(
                        writer,
                        index=False,
                        sheet_name="Duplicate UIN"
                    )

                # -------------------------
                # 7. Case Status
                # -------------------------
                if status_col:
                    status_summary.to_excel(
                        writer,
                        index=False,
                        sheet_name="Case Status"
                    )
                else:
                    pd.DataFrame({"Message": ["Status column not found"]}).to_excel(
                        writer,
                        index=False,
                        sheet_name="Case Status"
                    )

                # -------------------------
                # 8. Filtered Data
                # -------------------------
                filtered_df.to_excel(
                    writer,
                    index=False,
                    sheet_name="Filtered Data"
                )

            master_output.seek(0)

            st.success(
                "✅ Complete Master Report तैयार है — सभी reports एक Excel file में हैं।"
            )

            st.download_button(
                "⬇️ Download Complete Master Report",
                data=master_output.getvalue(),
                file_name="Complete_Master_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_complete_master_report"
            )

        # Data preview
        st.subheader("📋 Data Preview")

        st.dataframe(
            filtered_df,
            use_container_width=True,
            height=450
        )

        # Filtered report
        csv_data = (
            filtered_df
            .to_csv(index=False)
            .encode("utf-8")
        )

        st.download_button(
            "⬇️ Download Filtered Report",
            data=csv_data,
            file_name="filtered_report.csv",
            mime="text/csv",
            key="download_filtered_report"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
