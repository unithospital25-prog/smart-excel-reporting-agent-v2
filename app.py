import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Smart Excel Reporting Agent", page_icon="📊", layout="wide")
st.title("📊 Smart Excel Reporting Agent")
st.write("Upload and analyze your Excel data.")


def make_excel(df, sheet="Report"):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=str(sheet)[:31])
    out.seek(0)
    return out.getvalue()


def find_column(columns, names):
    names = {str(x).strip().lower() for x in names}
    for col in columns:
        if str(col).strip().lower() in names:
            return col
    return None


def clean(series):
    return series.fillna("").astype(str).str.strip().str.casefold()


def missing_mask(series, include_no=False):
    vals = clean(series)
    missing = {"", "none", "nan", "null", "na", "n/a", "-"}
    if include_no:
        missing |= {"no", "pending", "not available"}
    return vals.isin(missing)


def status_counts(df, status_col):
    if not status_col or status_col not in df.columns:
        return 0, 0
    s = clean(df[status_col])
    pending = int(s.str.contains("pending", regex=False, na=False).sum())
    completed = int(s.str.contains("complete|completed|done", regex=True, na=False).sum())
    return pending, completed


def contribution(df, amount_col):
    if not amount_col or amount_col not in df.columns:
        return 0
    return pd.to_numeric(df[amount_col], errors="coerce").fillna(0).sum()


def bn_summary(df, bn_col, status_col, amount_col, pran_col, nomination_col, uin_col):
    if not bn_col or df.empty:
        return pd.DataFrame()

    result = df.groupby(bn_col, dropna=False).size().reset_index(name="Total Employees")
    result[bn_col] = result[bn_col].fillna("Blank BN").astype(str)

    if status_col:
        s = clean(df[status_col])
        tmp = df.copy()
        tmp["_status"] = s
        p = tmp[tmp["_status"].str.contains("pending", regex=False, na=False)].groupby(bn_col).size().rename("Pending")
        c = tmp[tmp["_status"].str.contains("complete|completed|done", regex=True, na=False)].groupby(bn_col).size().rename("Completed")
        result = result.merge(p, on=bn_col, how="left").merge(c, on=bn_col, how="left")
    else:
        result["Pending"] = 0
        result["Completed"] = 0

    if pran_col:
        x = df.assign(_missing=missing_mask(df[pran_col])).groupby(bn_col)["_missing"].sum().rename("Missing PRAN")
        result = result.merge(x, on=bn_col, how="left")
    else:
        result["Missing PRAN"] = 0

    if nomination_col:
        x = df.assign(_missing=missing_mask(df[nomination_col], True)).groupby(bn_col)["_missing"].sum().rename("Missing Nomination")
        result = result.merge(x, on=bn_col, how="left")
    else:
        result["Missing Nomination"] = 0

    if uin_col:
        tmp = df.copy()
        tmp["_uin"] = tmp[uin_col].fillna("").astype(str).str.strip()
        tmp["_dup"] = tmp["_uin"].ne("") & tmp["_uin"].duplicated(keep=False)
        x = tmp.groupby(bn_col)["_dup"].sum().rename("Duplicate UIN Records")
        result = result.merge(x, on=bn_col, how="left")
    else:
        result["Duplicate UIN Records"] = 0

    if amount_col:
        tmp = df.copy()
        tmp["_amount"] = pd.to_numeric(tmp[amount_col], errors="coerce").fillna(0)
        x = tmp.groupby(bn_col)["_amount"].sum().rename("Total Contribution")
        result = result.merge(x, on=bn_col, how="left")
    else:
        result["Total Contribution"] = 0

    for col in ["Pending", "Completed", "Missing PRAN", "Missing Nomination", "Duplicate UIN Records"]:
        result[col] = result[col].fillna(0).astype(int)
    result["Total Contribution"] = result["Total Contribution"].fillna(0)
    return result


def safe_sheet(name, used):
    name = re.sub(r'[\[\]:*?/\\]', '_', str(name).strip() or "Blank BN")[:31]
    if name in {"Summary", "BN Summary", "Filtered Data"}:
        name = (name[:25] + "_Data")[:31]
    base = name
    n = 2
    while name in used:
        suffix = f"_{n}"
        name = base[:31-len(suffix)] + suffix
        n += 1
    used.add(name)
    return name


def make_bn_wise_workbook(df, bn_col, status_col, amount_col, pran_col, nomination_col, uin_col, employee_name_col):
    out = BytesIO()
    bns = bn_summary(df, bn_col, status_col, amount_col, pran_col, nomination_col, uin_col)

    if uin_col:
        unique = df[uin_col].fillna("").astype(str).str.strip().replace("", pd.NA).dropna().nunique()
    elif employee_name_col:
        unique = df[employee_name_col].fillna("").astype(str).str.strip().replace("", pd.NA).dropna().nunique()
    else:
        unique = len(df)

    pending, completed = status_counts(df, status_col)
    total_amount = contribution(df, amount_col)
    missing_pran = int(missing_mask(df[pran_col]).sum()) if pran_col else 0
    missing_nom = int(missing_mask(df[nomination_col], True).sum()) if nomination_col else 0
    duplicate = 0
    if uin_col:
        u = df[uin_col].fillna("").astype(str).str.strip()
        duplicate = int((u.ne("") & u.duplicated(keep=False)).sum())

    summary = pd.DataFrame({
        "Particular": ["Total Records", "Pending", "Completed", "Total Contribution", "Unique Employees", "Missing PRAN", "Missing Nomination", "Duplicate UIN Records"],
        "Value": [len(df), pending, completed, total_amount, unique, missing_pran, missing_nom, duplicate]
    })

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        summary.to_excel(writer, index=False, sheet_name="Summary")
        bns.to_excel(writer, index=False, sheet_name="BN Summary")
        df.to_excel(writer, index=False, sheet_name="Filtered Data")

        used = {"Summary", "BN Summary", "Filtered Data"}
        values = df[bn_col].fillna("Blank BN").astype(str).str.strip().replace("", "Blank BN").unique()
        for bn in sorted(values, key=lambda x: str(x).casefold()):
            part = df[df[bn_col].fillna("Blank BN").astype(str).str.strip().replace("", "Blank BN") == bn].copy()
            sheet = safe_sheet(bn, used)
            p, c = status_counts(part, status_col)
            mp = int(missing_mask(part[pran_col]).sum()) if pran_col else 0
            mn = int(missing_mask(part[nomination_col], True).sum()) if nomination_col else 0
            du = 0
            if uin_col:
                u = part[uin_col].fillna("").astype(str).str.strip()
                du = int((u.ne("") & u.duplicated(keep=False)).sum())
            mini = pd.DataFrame({
                "Particular": ["BN", "Total Employees", "Pending", "Completed", "Missing PRAN", "Missing Nomination", "Duplicate UIN Records", "Total Contribution"],
                "Value": [bn, len(part), p, c, mp, mn, du, contribution(part, amount_col)]
            })
            mini.to_excel(writer, index=False, sheet_name=sheet, startrow=0)
            part.to_excel(writer, index=False, sheet_name=sheet, startrow=len(mini)+3)

    out.seek(0)
    return out.getvalue()


uploaded_file = st.file_uploader("📂 Upload Excel File", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        st.success("✅ Excel file uploaded successfully!")
        sheet_name = st.selectbox("📄 Select Sheet", excel_file.sheet_names)
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        df.columns = df.columns.astype(str).str.strip()

        bn_col = find_column(df.columns, {"bn", "battalion", "unit"})
        status_col = find_column(df.columns, {"status", "case status", "case_status"})
        amount_col = find_column(df.columns, {"amount", "contribution", "total amount"})
        uin_col = find_column(df.columns, {"uin", "employee number", "employee no", "employee id", "employee_id"})
        employee_name_col = find_column(df.columns, {"employee name", "name", "employee_name"})
        pran_col = find_column(df.columns, {"pran", "pran no", "pran number", "pran_number"})
        nomination_col = find_column(df.columns, {"nomination", "nomination status", "nominee", "nominee status"})

        st.subheader("🔎 Search & Filter")
        search_text = st.text_input("🔍 Search UIN, Employee Name, BN, PRAN, Mobile, Status etc.")
        filtered_df = df.copy()
        if search_text.strip():
            q = search_text.strip().casefold()
            data = filtered_df.fillna("").astype(str)
            mask = data.apply(lambda col: col.str.casefold().str.contains(re.escape(q), regex=True, na=False)).any(axis=1)
            filtered_df = filtered_df.loc[mask].copy()
            st.info(f"🔍 Search result: {len(filtered_df)} record(s) found")

        if bn_col:
            bn_values = sorted(df[bn_col].dropna().astype(str).unique().tolist())
            selected_bn = st.selectbox("🏢 Select BN", ["All"] + bn_values)
            if selected_bn != "All":
                filtered_df = filtered_df[filtered_df[bn_col].astype(str) == selected_bn]

        if status_col:
            status_values = sorted(df[status_col].dropna().astype(str).unique().tolist())
            selected_status = st.selectbox("📌 Case Status", ["All"] + status_values)
            if selected_status != "All":
                filtered_df = filtered_df[filtered_df[status_col].astype(str) == selected_status]

        st.subheader("📊 Automatic Dashboard")
        total_records = len(filtered_df)
        pending_count, completed_count = status_counts(filtered_df, status_col)
        total_amount = contribution(filtered_df, amount_col)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("📋 Total Records", total_records)
        c2.metric("⏳ Pending", pending_count)
        c3.metric("✅ Completed", completed_count)
        c4.metric("💰 Contribution", f"₹{total_amount:,.0f}")

        if bn_col:
            st.subheader("🏢 BN-wise Summary")
            bns = bn_summary(filtered_df, bn_col, status_col, amount_col, pran_col, nomination_col, uin_col)
            st.dataframe(bns, use_container_width=True, hide_index=True)

        if status_col:
            st.subheader("📌 Case Status Summary")
            status_summary = filtered_df[status_col].fillna("Blank").astype(str).value_counts().reset_index()
            status_summary.columns = ["Case Status", "Total"]
            st.dataframe(status_summary, use_container_width=True, hide_index=True)

        if bn_col:
            st.subheader("📊 BN Report")
            if st.button("📊 Generate BN Report", key="generate_bn_report"):
                report = bn_summary(filtered_df, bn_col, status_col, amount_col, pran_col, nomination_col, uin_col)
                st.dataframe(report, use_container_width=True, hide_index=True)
                st.download_button("⬇️ Download BN Report", make_excel(report, "BN Report"), "BN_Report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_bn_report")

        # NEW FEATURE
        st.subheader("📚 BN-wise Complete Excel Report")
        st.write("एक button से Summary, BN Summary और हर BN की अलग Excel Sheet तैयार करें।")
        if st.button("📚 Generate BN-wise Complete Excel Report", key="generate_bn_wise_complete_report"):
            if not bn_col:
                st.error("❌ BN / Battalion / Unit column नहीं मिला।")
            elif filtered_df.empty:
                st.warning("⚠️ Report बनाने के लिए कोई record उपलब्ध नहीं है।")
            else:
                data = make_bn_wise_workbook(filtered_df, bn_col, status_col, amount_col, pran_col, nomination_col, uin_col, employee_name_col)
                st.success("✅ BN-wise Complete Excel Report तैयार है!")
                st.download_button("⬇️ Download BN-wise Complete Excel Report", data, "BN_Wise_Complete_Report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_bn_wise_complete_report")

        st.subheader("👤 Smart Employee-wise Report")
        employee_report = filtered_df.copy()
        if uin_col:
            unique_employee_count = employee_report[uin_col].dropna().astype(str).str.strip().replace("", pd.NA).dropna().nunique()
        elif employee_name_col:
            unique_employee_count = employee_report[employee_name_col].dropna().astype(str).str.strip().replace("", pd.NA).dropna().nunique()
        else:
            unique_employee_count = len(employee_report)
        st.write(f"👥 Total Unique Employees: **{unique_employee_count}**")
        if not employee_report.empty:
            st.dataframe(employee_report, use_container_width=True, hide_index=True, height=400)
            st.download_button("⬇️ Download Employee Excel Report", make_excel(employee_report, "Employee Report"), "Employee_Wise_Report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_employee_report")
        else:
            st.info("ℹ️ No employee records found.")

        duplicate_uin_report = pd.DataFrame()
        if uin_col:
            st.subheader("⚠️ Duplicate UIN Report")
            u = employee_report[uin_col].fillna("").astype(str).str.strip()
            duplicate_uin_report = employee_report[u.ne("") & u.duplicated(keep=False)].copy()
            if not duplicate_uin_report.empty:
                st.warning(f"⚠️ Duplicate UIN Records: {len(duplicate_uin_report)}")
                st.dataframe(duplicate_uin_report, use_container_width=True, hide_index=True)
                st.download_button("⬇️ Download Duplicate UIN Report", make_excel(duplicate_uin_report, "Duplicate UIN"), "Duplicate_UIN_Report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_duplicate_uin")
            else:
                st.success("✅ No Duplicate UIN found.")

        missing_pran_report = pd.DataFrame()
        if pran_col:
            st.subheader("❌ Missing PRAN Report")
            missing_pran_report = employee_report[missing_mask(employee_report[pran_col])].copy()
            if not missing_pran_report.empty:
                st.warning(f"❌ Employees without PRAN: {len(missing_pran_report)}")
                st.dataframe(missing_pran_report, use_container_width=True, hide_index=True)
                st.download_button("⬇️ Download Missing PRAN Report", make_excel(missing_pran_report, "Missing PRAN"), "Missing_PRAN_Report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_missing_pran")
            else:
                st.success("✅ All employees have PRAN.")

        missing_nomination_report = pd.DataFrame()
        if nomination_col:
            st.subheader("❌ Missing Nomination Report")
            missing_nomination_report = employee_report[missing_mask(employee_report[nomination_col], True)].copy()
            if not missing_nomination_report.empty:
                st.warning(f"❌ Employees with missing nomination: {len(missing_nomination_report)}")
                st.dataframe(missing_nomination_report, use_container_width=True, hide_index=True)
                st.download_button("⬇️ Download Missing Nomination Report", make_excel(missing_nomination_report, "Missing Nomination"), "Missing_Nomination_Report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_missing_nomination")
            else:
                st.success("✅ All employees have nomination.")

        st.subheader("📦 Complete Master Report")
        st.write("एक button से सभी महत्वपूर्ण reports एक ही Excel file में तैयार करें।")
        if st.button("📦 Generate Complete Master Report", key="generate_master_report"):
            out = BytesIO()
            with pd.ExcelWriter(out, engine="openpyxl") as writer:
                pd.DataFrame({
                    "Particular": ["Total Records", "Pending", "Completed", "Total Contribution", "Unique Employees", "Missing PRAN", "Missing Nomination", "Duplicate UIN Records"],
                    "Value": [total_records, pending_count, completed_count, total_amount, unique_employee_count, len(missing_pran_report), len(missing_nomination_report), len(duplicate_uin_report)]
                }).to_excel(writer, index=False, sheet_name="Summary")
                if bn_col:
                    bn_summary(filtered_df, bn_col, status_col, amount_col, pran_col, nomination_col, uin_col).to_excel(writer, index=False, sheet_name="BN Summary")
                employee_report.to_excel(writer, index=False, sheet_name="Employee Report")
                (missing_pran_report if pran_col else pd.DataFrame({"Message": ["PRAN column not found"]})).to_excel(writer, index=False, sheet_name="Missing PRAN")
                (missing_nomination_report if nomination_col else pd.DataFrame({"Message": ["Nomination column not found"]})).to_excel(writer, index=False, sheet_name="Missing Nomination")
                (duplicate_uin_report if uin_col else pd.DataFrame({"Message": ["UIN column not found"]})).to_excel(writer, index=False, sheet_name="Duplicate UIN")
                (status_summary if status_col else pd.DataFrame({"Message": ["Status column not found"]})).to_excel(writer, index=False, sheet_name="Case Status")
                filtered_df.to_excel(writer, index=False, sheet_name="Filtered Data")
            out.seek(0)
            st.success("✅ Complete Master Report तैयार है — सभी reports एक Excel file में हैं।")
            st.download_button("⬇️ Download Complete Master Report", out.getvalue(), "Complete_Master_Report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_complete_master_report")

        st.subheader("📋 Data Preview")
        st.dataframe(filtered_df, use_container_width=True, height=450)
        st.download_button("⬇️ Download Filtered Report", make_excel(filtered_df, "Filtered Data"), "Filtered_Report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_filtered_report")

    except Exception as e:
        st.error(f"❌ Error: {e}")
