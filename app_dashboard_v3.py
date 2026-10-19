import streamlit as st
import pandas as pd
import re
from io import BytesIO

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="Smart Excel Reporting Agent",
    page_icon="X",
    layout="wide"
)

# =========================================================
# CUSTOM UI
# =========================================================
st.markdown("""
<style>
.main-title {
    font-size: 42px;
    font-weight: 800;
    margin-bottom: 5px;
}
.sub-title {
    font-size: 18px;
    color: #6b7280;
    margin-bottom: 20px;
}
</style>
""", unsafe_allow_html=True)

st.markdown(
    '<div class="main-title">Smart Excel Reporting Agent</div>',
    unsafe_allow_html=True
)
st.markdown(
    '<div class="sub-title">Upload and analyze your Excel data - single or multiple files.</div>',
    unsafe_allow_html=True
)

# =========================================================
# HELPERS
# =========================================================
def make_excel(dataframe, sheet_name="Report"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name=sheet_name)
        format_excel_writer(writer, dataframe, sheet_name)
    output.seek(0)
    return output.getvalue()


def safe_sheet_name(name, used_names=None):
    name = str(name)
    name = re.sub(r'[\[\]\:\*\?\/\\]', '_', name)
    name = name[:31] or "Sheet"
    if used_names is None:
        return name

    original = name
    counter = 1
    while name in used_names:
        suffix = f"_{counter}"
        name = (original[:31 - len(suffix)] + suffix)
        counter += 1
    return name


def format_excel_writer(writer, dataframe, sheet_name):
    """Official-looking Excel formatting with automatic widths."""
    try:
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        from openpyxl.worksheet.table import Table, TableStyleInfo

        ws = writer.book[sheet_name]

        header_fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78"
        )
        header_font = Font(
            name="Aptos",
            size=11,
            bold=True,
            color="FFFFFF"
        )
        body_font = Font(
            name="Aptos",
            size=10,
            color="000000"
        )
        thin_border = Border(
            left=Side(style="thin", color="D9E1F2"),
            right=Side(style="thin", color="D9E1F2"),
            top=Side(style="thin", color="D9E1F2"),
            bottom=Side(style="thin", color="D9E1F2")
        )

        # Header
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )
            cell.border = thin_border

        ws.row_dimensions[1].height = 28
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        # Body
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.font = body_font
                cell.alignment = Alignment(
                    vertical="center",
                    wrap_text=False
                )
                cell.border = thin_border

        # Automatic column width
        for col_idx, column_cells in enumerate(
            ws.iter_cols(
                min_row=1,
                max_row=ws.max_row,
                min_col=1,
                max_col=ws.max_column
            ),
            start=1
        ):
            max_length = 0
            for cell in column_cells:
                value = "" if cell.value is None else str(cell.value)
                max_length = max(max_length, len(value))

            width = min(max(max_length + 3, 12), 35)
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        # Date / amount formatting
        for col_idx, col_name in enumerate(dataframe.columns, start=1):
            lower_name = str(col_name).strip().casefold()

            for row_idx in range(2, ws.max_row + 1):
                cell = ws.cell(row=row_idx, column=col_idx)

                if "date" in lower_name:
                    cell.number_format = "dd-mm-yyyy"

                if (
                    "amount" in lower_name
                    or "contribution" in lower_name
                    or "total" in lower_name and "amount" in lower_name
                ):
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0.00'

        # Add Excel table when data exists
        if ws.max_row >= 2 and ws.max_column >= 1:
            table_name = re.sub(
                r"[^A-Za-z0-9_]",
                "",
                f"Table_{sheet_name}"
            )[:200]

            if not table_name or table_name[0].isdigit():
                table_name = "Table_Report"

            try:
                tab = Table(
                    displayName=table_name,
                    ref=ws.dimensions
                )
                style = TableStyleInfo(
                    name="TableStyleMedium2",
                    showFirstColumn=False,
                    showLastColumn=False,
                    showRowStripes=True,
                    showColumnStripes=False
                )
                tab.tableStyleInfo = style
                ws.add_table(tab)
            except Exception:
                pass

    except Exception:
        pass


def normalize_column_key(value):
    """Create a stable key for matching Excel headings across files."""
    text = str(value).replace("\n", " ").replace("\r", " ")
    text = re.sub(r"\s+", " ", text).strip().casefold()
    text = re.sub(r"[^a-z0-9]+", "_", text).strip("_")
    return text


def normalize_columns(columns):
    """Normalize headings and merge duplicate headings safely.

    Example: UIN, Uin, UIN  and U-I-N become the same logical heading.
    Original readable headings are retained where possible.
    """
    result = []
    used = {}

    preferred_names = {
        "uin": "UIN",
        "employee_name": "Employee Name",
        "name": "Name",
        "bn": "BN",
        "battalion": "Battalion",
        "unit": "Unit",
        "pran": "PRAN",
        "pran_no": "PRAN No",
        "pran_number": "PRAN Number",
        "status": "Status",
        "case_status": "Case Status",
        "amount": "Amount",
        "contribution": "Contribution",
        "total_amount": "Total Amount",
        "nomination": "Nomination",
        "nomination_status": "Nomination Status",
        "nominee": "Nominee",
        "nominee_status": "Nominee Status",
        "mobile": "Mobile",
        "mobile_no": "Mobile No"
    }

    for original in columns:
        original_text = str(original).strip()
        key = normalize_column_key(original_text)
        base_name = preferred_names.get(key, original_text)

        if base_name == "":
            base_name = "Unnamed Column"

        if base_name in used:
            used[base_name] += 1
            final_name = f"{base_name}_{used[base_name]}"
        else:
            used[base_name] = 1
            final_name = base_name

        result.append(final_name)

    return result


def find_column(columns, possible_names):
    normalized = {
        str(x).strip().casefold(): x
        for x in columns
    }

    for possible in possible_names:
        key = str(possible).strip().casefold()
        if key in normalized:
            return normalized[key]

    return None


def status_mask(dataframe, status_col, pattern):
    if not status_col:
        return pd.Series(False, index=dataframe.index)

    status_data = (
        dataframe[status_col]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.casefold()
    )

    return status_data.str.contains(
        pattern,
        regex=True,
        na=False
    )


def missing_mask(series, include_no=False):
    data = (
        series
        .fillna("")
        .astype(str)
        .str.strip()
        .str.casefold()
    )

    values = [
        "",
        "none",
        "nan",
        "null",
        "na",
        "n/a",
        "-"
    ]

    if include_no:
        values.extend([
            "no",
            "pending",
            "not available",
            "not available"
        ])

    return data.isin(values)


def build_bn_summary(
    dataframe,
    bn_col,
    status_col,
    amount_col,
    pran_col,
    nomination_col,
    uin_col
):
    if not bn_col:
        return pd.DataFrame()

    rows = []

    for bn_value, group in dataframe.groupby(
        bn_col,
        dropna=False,
        sort=True
    ):
        bn_display = str(bn_value).strip()

        if bn_display.casefold() in ["", "nan", "none"]:
            bn_display = "Blank BN"

        total = len(group)

        pending = (
            int(status_mask(group, status_col, "pending").sum())
            if status_col else 0
        )

        completed = (
            int(
                status_mask(
                    group,
                    status_col,
                    "complete|completed|done"
                ).sum()
            )
            if status_col else 0
        )

        missing_pran = (
            int(missing_mask(group[pran_col]).sum())
            if pran_col else 0
        )

        missing_nomination = (
            int(
                missing_mask(
                    group[nomination_col],
                    include_no=True
                ).sum()
            )
            if nomination_col else 0
        )

        duplicate_uin = 0
        if uin_col:
            uin_data = (
                group[uin_col]
                .fillna("")
                .astype(str)
                .str.strip()
            )
            duplicate_mask = (
                uin_data.ne("")
                & uin_data.duplicated(keep=False)
            )
            duplicate_uin = int(duplicate_mask.sum())

        if amount_col:
            contribution = pd.to_numeric(
                group[amount_col],
                errors="coerce"
            ).fillna(0).sum()
        else:
            contribution = 0

        rows.append({
            "BN": bn_display,
            "Total Employees": total,
            "Pending": pending,
            "Completed": completed,
            "Missing PRAN": missing_pran,
            "Missing Nomination": missing_nomination,
            "Duplicate UIN Records": duplicate_uin,
            "Total Contribution": contribution
        })

    return pd.DataFrame(rows)


# =========================================================
# MULTIPLE EXCEL FILE UPLOAD
# =========================================================
uploaded_files = st.file_uploader(
    "Upload Excel File(s) - Multiple files allowed",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    help="You can upload multiple Excel files together."
)

if not uploaded_files:
    st.info("Please upload one or more Excel files to start.")
    st.stop()

st.success(
    f"{len(uploaded_files)} Excel file(s) uploaded successfully."
)

# =========================================================
# COMBINE ALL FILES + ALL SHEETS
# =========================================================
all_dataframes = []
file_sheet_summary = []

for uploaded_file in uploaded_files:
    try:
        excel_file = pd.ExcelFile(uploaded_file)

        for sheet_name in excel_file.sheet_names:
            temp_df = pd.read_excel(
                uploaded_file,
                sheet_name=sheet_name
            )

            if temp_df.empty:
                continue

            # Clean and normalize column headings before combining files.
            # This makes headings such as "UIN", "Uin", "UIN ", etc.
            # behave as the same column across different Excel files.
            temp_df.columns = normalize_columns(temp_df.columns)

            # Remove completely blank rows. This prevents empty Excel rows
            # from appearing as None records in the combined master data.
            if len(temp_df.columns) > 0:
                blank_row_mask = temp_df.apply(
                    lambda row: row.astype(str)
                    .str.strip()
                    .replace({"nan": "", "None": ""})
                    .eq("")
                    .all(),
                    axis=1
                )
                temp_df = temp_df.loc[~blank_row_mask].copy()

            if temp_df.empty:
                continue

            # Keep source information for audit/tracking
            temp_df["Source File"] = uploaded_file.name
            temp_df["Source Sheet"] = sheet_name

            all_dataframes.append(temp_df)

            file_sheet_summary.append({
                "Source File": uploaded_file.name,
                "Source Sheet": sheet_name,
                "Records": len(temp_df)
            })

    except Exception as e:
        st.warning(
            f"Unable to read {uploaded_file.name}: {e}"
        )

if not all_dataframes:
    st.error("No readable data was found in the uploaded Excel files.")
    st.stop()

# Combine all files
df = pd.concat(
    all_dataframes,
    ignore_index=True,
    sort=False
)

df.columns = normalize_columns(df.columns)

# Remove any rows that are completely blank after the final merge.
# Source File and Source Sheet are ignored for this check because they
# are intentionally populated for every valid record.
data_columns = [
    col for col in df.columns
    if col not in ["Source File", "Source Sheet"]
]

if data_columns:
    final_blank_mask = df[data_columns].apply(
        lambda row: row.astype(str)
        .str.strip()
        .replace({"nan": "", "None": ""})
        .eq("")
        .all(),
        axis=1
    )
    df = df.loc[~final_blank_mask].reset_index(drop=True)

st.success(
    f"All uploaded files were combined successfully. "
    f"Total Records: {len(df):,}"
)

# =========================================================
# SOURCE FILE SUMMARY
# =========================================================
with st.expander("Uploaded Files Summary"):
    source_summary = pd.DataFrame(file_sheet_summary)
    st.dataframe(
        source_summary,
        use_container_width=True,
        hide_index=True
    )

# =========================================================
# COLUMN DETECTION
# =========================================================
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
    {
        "employee name",
        "name",
        "employee_name"
    }
)

pran_col = find_column(
    df.columns,
    {
        "pran",
        "pran no",
        "pran number",
        "pran_number"
    }
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

# =========================================================
# SEARCH & FILTER
# =========================================================
st.subheader("Search and Filter")

search_text = st.text_input(
    "Search UIN, Employee Name, BN, PRAN, Mobile, Status, etc."
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
        f"Search result: {len(filtered_df)} record(s) found."
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
        "Select BN",
        ["All"] + bn_values
    )

    if selected_bn != "All":
        filtered_df = filtered_df[
            filtered_df[bn_col].astype(str) == selected_bn
        ].copy()

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
        "Case Status",
        ["All"] + status_values
    )

    if selected_status != "All":
        filtered_df = filtered_df[
            filtered_df[status_col].astype(str) == selected_status
        ].copy()

# =========================================================
# DASHBOARD V3
# =========================================================
st.subheader("Automatic Dashboard")

# Dashboard figures always follow the currently selected search/BN/status filters.
total_records = len(filtered_df)

pending_count = (
    int(status_mask(filtered_df, status_col, "pending").sum())
    if status_col else 0
)

completed_count = (
    int(
        status_mask(
            filtered_df,
            status_col,
            "complete|completed|done"
        ).sum()
    )
    if status_col else 0
)

total_amount = (
    pd.to_numeric(
        filtered_df[amount_col],
        errors="coerce"
    ).fillna(0).sum()
    if amount_col else 0
)

if uin_col:
    unique_employee_count = (
        filtered_df[uin_col]
        .dropna()
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .nunique()
    )
elif employee_name_col:
    unique_employee_count = (
        filtered_df[employee_name_col]
        .dropna()
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .nunique()
    )
else:
    unique_employee_count = len(filtered_df)

missing_pran_count = (
    int(missing_mask(filtered_df[pran_col]).sum())
    if pran_col else 0
)

missing_nomination_count = (
    int(
        missing_mask(
            filtered_df[nomination_col],
            include_no=True
        ).sum()
    )
    if nomination_col else 0
)

# Compact responsive dashboard cards. On mobile these automatically wrap
# into multiple columns instead of taking the full screen height.
st.markdown(
    f"""
    <style>
    .dashboard-grid {{
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 12px;
        margin: 8px 0 24px 0;
    }}
    .dashboard-card {{
        border: 1px solid rgba(128,128,128,0.28);
        border-radius: 12px;
        padding: 16px 14px;
        background: rgba(128,128,128,0.07);
        min-height: 92px;
    }}
    .dashboard-label {{
        font-size: 14px;
        color: #8b93a1;
        margin-bottom: 8px;
    }}
    .dashboard-value {{
        font-size: 28px;
        font-weight: 750;
        line-height: 1.1;
    }}
    @media (max-width: 700px) {{
        .dashboard-grid {{
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 9px;
        }}
        .dashboard-card {{
            padding: 13px 11px;
            min-height: 82px;
        }}
        .dashboard-label {{
            font-size: 12px;
        }}
        .dashboard-value {{
            font-size: 23px;
        }}
    }}
    </style>

    <div class="dashboard-grid">
        <div class="dashboard-card">
            <div class="dashboard-label">Total Records</div>
            <div class="dashboard-value">{total_records:,}</div>
        </div>
        <div class="dashboard-card">
            <div class="dashboard-label">Unique Employees</div>
            <div class="dashboard-value">{unique_employee_count:,}</div>
        </div>
        <div class="dashboard-card">
            <div class="dashboard-label">Pending</div>
            <div class="dashboard-value">{pending_count:,}</div>
        </div>
        <div class="dashboard-card">
            <div class="dashboard-label">Completed</div>
            <div class="dashboard-value">{completed_count:,}</div>
        </div>
        <div class="dashboard-card">
            <div class="dashboard-label">Missing PRAN</div>
            <div class="dashboard-value">{missing_pran_count:,}</div>
        </div>
        <div class="dashboard-card">
            <div class="dashboard-label">Missing Nomination</div>
            <div class="dashboard-value">{missing_nomination_count:,}</div>
        </div>
        <div class="dashboard-card">
            <div class="dashboard-label">Total Contribution</div>
            <div class="dashboard-value">INR {total_amount:,.0f}</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

# =========================================================
# DASHBOARD CHARTS
# =========================================================
st.markdown("### Dashboard Charts")

chart_col1, chart_col2 = st.columns(2)

with chart_col1:
    st.markdown("**BN-wise Employee Count**")

    if bn_col:
        bn_chart = (
            filtered_df[bn_col]
            .fillna("Blank BN")
            .astype(str)
            .str.strip()
            .replace("", "Blank BN")
            .value_counts()
            .rename_axis("BN")
            .reset_index(name="Employees")
            .set_index("BN")
        )

        if not bn_chart.empty:
            st.bar_chart(bn_chart, use_container_width=True)
        else:
            st.info("No BN data available for the selected filters.")
    else:
        st.info("BN column not found.")

with chart_col2:
    st.markdown("**Case Status**")

    if status_col:
        status_chart = (
            filtered_df[status_col]
            .fillna("Blank")
            .astype(str)
            .str.strip()
            .replace("", "Blank")
            .value_counts()
            .rename_axis("Status")
            .reset_index(name="Records")
            .set_index("Status")
        )

        if not status_chart.empty:
            st.bar_chart(status_chart, use_container_width=True)
        else:
            st.info("No status data available for the selected filters.")
    else:
        st.info("Case Status column not found.")

# =========================================================
# BN-WISE SUMMARY
# =========================================================
bn_summary = pd.DataFrame()

if bn_col:
    st.subheader("BN-wise Summary")

    bn_summary = build_bn_summary(
        filtered_df,
        bn_col,
        status_col,
        amount_col,
        pran_col,
        nomination_col,
        uin_col
    )

    st.dataframe(
        bn_summary,
        use_container_width=True,
        hide_index=True
    )

# =========================================================
# CASE STATUS SUMMARY
# =========================================================
status_summary = pd.DataFrame()

if status_col:
    st.subheader("Case Status Summary")

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

# =========================================================
# BN REPORT
# =========================================================
if bn_col:
    st.subheader("BN Report")

    if st.button(
        "Generate BN Report",
        key="generate_bn_report"
    ):
        st.dataframe(
            bn_summary,
            use_container_width=True,
            hide_index=True
        )

        bn_excel = make_excel(
            bn_summary,
            "BN Report"
        )

        st.download_button(
            "Download BN Report",
            data=bn_excel,
            file_name="BN_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_bn_report"
        )

# =========================================================
# SMART EMPLOYEE-WISE REPORT
# =========================================================
st.subheader("Employee-wise Report")

employee_report = filtered_df.copy()

st.write(
    f"Total Unique Employees: **{unique_employee_count}**"
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
        "Download Employee Excel Report",
        data=employee_excel,
        file_name="Employee_Wise_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_employee_report"
    )
else:
    st.info("No employee records found.")

# =========================================================
# DUPLICATE UIN
# =========================================================
duplicate_uin_report = pd.DataFrame()

if uin_col:
    st.subheader("Duplicate UIN Report")

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
            f"Duplicate UIN Records: "
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
            "Download Duplicate UIN Report",
            data=duplicate_excel,
            file_name="Duplicate_UIN_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_duplicate_uin"
        )
    else:
        st.success("No duplicate UIN found.")

# =========================================================
# MISSING PRAN
# =========================================================
missing_pran_report = pd.DataFrame()

if pran_col:
    st.subheader("Missing PRAN Report")

    missing_pran_mask = missing_mask(
        employee_report[pran_col]
    )

    missing_pran_report = (
        employee_report[missing_pran_mask].copy()
    )

    if len(missing_pran_report) > 0:
        st.warning(
            f"Employees without PRAN: "
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
            "Download Missing PRAN Report",
            data=missing_pran_excel,
            file_name="Missing_PRAN_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_missing_pran"
        )
    else:
        st.success("All employees have PRAN.")

# =========================================================
# MISSING NOMINATION
# =========================================================
missing_nomination_report = pd.DataFrame()

if nomination_col:
    st.subheader("Missing Nomination Report")

    missing_nomination_mask = missing_mask(
        employee_report[nomination_col],
        include_no=True
    )

    missing_nomination_report = (
        employee_report[missing_nomination_mask].copy()
    )

    if len(missing_nomination_report) > 0:
        st.warning(
            f"Employees with missing nomination: "
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
            "Download Missing Nomination Report",
            data=nomination_excel,
            file_name="Missing_Nomination_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_missing_nomination"
        )
    else:
        st.success("All employees have nomination.")

# =========================================================
# BN-WISE COMPLETE EXCEL REPORT
# =========================================================
if bn_col:
    st.subheader("BN-wise Complete Excel Report")

    st.write(
        "One Excel file containing Summary, BN Summary, Filtered Data, "
        "and a separate sheet for each BN."
    )

    if st.button(
        "Generate BN-wise Complete Excel Report",
        key="generate_bn_complete_report"
    ):
        bn_output = BytesIO()
        used_sheet_names = set()

        with pd.ExcelWriter(
            bn_output,
            engine="openpyxl"
        ) as writer:

            # Summary
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
                    len(missing_pran_report),
                    len(missing_nomination_report),
                    len(duplicate_uin_report)
                ]
            })

            summary_data.to_excel(
                writer,
                index=False,
                sheet_name="Summary"
            )
            format_excel_writer(
                writer,
                summary_data,
                "Summary"
            )
            used_sheet_names.add("Summary")

            # BN Summary
            bn_summary.to_excel(
                writer,
                index=False,
                sheet_name="BN Summary"
            )
            format_excel_writer(
                writer,
                bn_summary,
                "BN Summary"
            )
            used_sheet_names.add("BN Summary")

            # Filtered Data
            filtered_df.to_excel(
                writer,
                index=False,
                sheet_name="Filtered Data"
            )
            format_excel_writer(
                writer,
                filtered_df,
                "Filtered Data"
            )
            used_sheet_names.add("Filtered Data")

            # One sheet for each BN
            for bn_value, group in filtered_df.groupby(
                bn_col,
                dropna=False,
                sort=True
            ):
                bn_display = str(bn_value).strip()

                if bn_display.casefold() in ["", "nan", "none"]:
                    bn_display = "Blank BN"

                sheet_name = safe_sheet_name(
                    f"BN-{bn_display}",
                    used_sheet_names
                )
                used_sheet_names.add(sheet_name)

                # Mini summary
                mini_summary = build_bn_summary(
                    group,
                    bn_col,
                    status_col,
                    amount_col,
                    pran_col,
                    nomination_col,
                    uin_col
                )

                # Convert to simple 2-column display
                if not mini_summary.empty:
                    mini_row = mini_summary.iloc[0]

                    mini = pd.DataFrame({
                        "Particular": [
                            "BN",
                            "Total Employees",
                            "Pending",
                            "Completed",
                            "Missing PRAN",
                            "Missing Nomination",
                            "Duplicate UIN Records",
                            "Total Contribution"
                        ],
                        "Value": [
                            mini_row["BN"],
                            mini_row["Total Employees"],
                            mini_row["Pending"],
                            mini_row["Completed"],
                            mini_row["Missing PRAN"],
                            mini_row["Missing Nomination"],
                            mini_row["Duplicate UIN Records"],
                            mini_row["Total Contribution"]
                        ]
                    })
                else:
                    mini = pd.DataFrame({
                        "Particular": ["BN"],
                        "Value": [bn_display]
                    })

                mini.to_excel(
                    writer,
                    index=False,
                    sheet_name=sheet_name,
                    startrow=0
                )

                # Full BN data below summary
                group.to_excel(
                    writer,
                    index=False,
                    sheet_name=sheet_name,
                    startrow=len(mini) + 2
                )

                # Format mini summary
                format_excel_writer(
                    writer,
                    mini,
                    sheet_name
                )

                # Format full data section manually
                try:
                    ws = writer.book[sheet_name]
                    from openpyxl.styles import Font, PatternFill, Alignment
                    from openpyxl.utils import get_column_letter

                    header_row = len(mini) + 3

                    for cell in ws[header_row]:
                        cell.fill = PatternFill(
                            fill_type="solid",
                            fgColor="1F4E78"
                        )
                        cell.font = Font(
                            name="Aptos",
                            size=11,
                            bold=True,
                            color="FFFFFF"
                        )
                        cell.alignment = Alignment(
                            horizontal="center",
                            vertical="center",
                            wrap_text=True
                        )

                    for col_idx in range(
                        1,
                        ws.max_column + 1
                    ):
                        max_len = 0
                        for row_idx in range(
                            1,
                            ws.max_row + 1
                        ):
                            value = ws.cell(
                                row=row_idx,
                                column=col_idx
                            ).value
                            max_len = max(
                                max_len,
                                len(str(value))
                                if value is not None else 0
                            )

                        ws.column_dimensions[
                            get_column_letter(col_idx)
                        ].width = min(
                            max(max_len + 3, 12),
                            35
                        )

                    ws.freeze_panes = f"A{header_row + 1}"

                except Exception:
                    pass

        bn_output.seek(0)

        st.success(
            "BN-wise Complete Excel Report is ready."
        )

        st.download_button(
            "Download BN-wise Complete Excel Report",
            data=bn_output.getvalue(),
            file_name="BN_Wise_Complete_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_bn_complete_report"
        )

# =========================================================
# COMPLETE MASTER REPORT
# =========================================================
st.subheader("Complete Master Report")

st.write(
    "Generate all important reports in a single Excel file."
)

if st.button(
    "Generate Complete Master Report",
    key="generate_master_report"
):
    master_output = BytesIO()

    with pd.ExcelWriter(
        master_output,
        engine="openpyxl"
    ) as writer:

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
                len(missing_pran_report),
                len(missing_nomination_report),
                len(duplicate_uin_report)
            ]
        })

        summary_data.to_excel(
            writer,
            index=False,
            sheet_name="Summary"
        )
        format_excel_writer(
            writer,
            summary_data,
            "Summary"
        )

        if bn_col:
            bn_summary.to_excel(
                writer,
                index=False,
                sheet_name="BN Summary"
            )
            format_excel_writer(
                writer,
                bn_summary,
                "BN Summary"
            )

        employee_report.to_excel(
            writer,
            index=False,
            sheet_name="Employee Report"
        )
        format_excel_writer(
            writer,
            employee_report,
            "Employee Report"
        )

        if pran_col:
            missing_pran_report.to_excel(
                writer,
                index=False,
                sheet_name="Missing PRAN"
            )
        else:
            pd.DataFrame({
                "Message": ["PRAN column not found"]
            }).to_excel(
                writer,
                index=False,
                sheet_name="Missing PRAN"
            )
        format_excel_writer(
            writer,
            missing_pran_report if pran_col else pd.DataFrame({
                "Message": ["PRAN column not found"]
            }),
            "Missing PRAN"
        )

        if nomination_col:
            missing_nomination_report.to_excel(
                writer,
                index=False,
                sheet_name="Missing Nomination"
            )
        else:
            pd.DataFrame({
                "Message": ["Nomination column not found"]
            }).to_excel(
                writer,
                index=False,
                sheet_name="Missing Nomination"
            )
        format_excel_writer(
            writer,
            missing_nomination_report if nomination_col else pd.DataFrame({
                "Message": ["Nomination column not found"]
            }),
            "Missing Nomination"
        )

        if uin_col:
            duplicate_uin_report.to_excel(
                writer,
                index=False,
                sheet_name="Duplicate UIN"
            )
        else:
            pd.DataFrame({
                "Message": ["UIN column not found"]
            }).to_excel(
                writer,
                index=False,
                sheet_name="Duplicate UIN"
            )
        format_excel_writer(
            writer,
            duplicate_uin_report if uin_col else pd.DataFrame({
                "Message": ["UIN column not found"]
            }),
            "Duplicate UIN"
        )

        if status_col:
            status_summary.to_excel(
                writer,
                index=False,
                sheet_name="Case Status"
            )
            format_excel_writer(
                writer,
                status_summary,
                "Case Status"
            )
        else:
            status_empty = pd.DataFrame({
                "Message": ["Status column not found"]
            })
            status_empty.to_excel(
                writer,
                index=False,
                sheet_name="Case Status"
            )
            format_excel_writer(
                writer,
                status_empty,
                "Case Status"
            )

        filtered_df.to_excel(
            writer,
            index=False,
            sheet_name="Filtered Data"
        )
        format_excel_writer(
            writer,
            filtered_df,
            "Filtered Data"
        )

    master_output.seek(0)

    st.success(
        "Complete Master Report is ready."
    )

    st.download_button(
        "Download Complete Master Report",
        data=master_output.getvalue(),
        file_name="Complete_Master_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_complete_master_report"
    )

# =========================================================
# DATA PREVIEW
# =========================================================
st.subheader("Data Preview")

st.dataframe(
    filtered_df,
    use_container_width=True,
    height=450
)

# =========================================================
# FILTERED EXCEL DOWNLOAD
# =========================================================
filtered_excel = make_excel(
    filtered_df,
    "Filtered Data"
)

st.download_button(
    "Download Filtered Excel Report",
    data=filtered_excel,
    file_name="Filtered_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    key="download_filtered_report"
)
