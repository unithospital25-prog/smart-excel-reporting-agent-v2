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
.dashboard-note {
    font-size: 13px;
    color: #6b7280;
    margin-top: -8px;
    margin-bottom: 10px;
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

    safe_df = dataframe.copy()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        safe_df.to_excel(
            writer,
            index=False,
            sheet_name=safe_sheet_name(sheet_name)
        )
        format_excel_writer(
            writer,
            safe_df,
            safe_sheet_name(sheet_name)
        )

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
        name = original[:31 - len(suffix)] + suffix
        counter += 1

    return name


def format_excel_writer(writer, dataframe, sheet_name):
    """Official Excel formatting with automatic column widths."""
    try:
        from openpyxl.styles import (
            Font,
            PatternFill,
            Alignment,
            Border,
            Side
        )
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
        if ws.max_row >= 1:
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

        # Body
        if ws.max_row >= 2:
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.font = body_font
                    cell.alignment = Alignment(
                        vertical="center",
                        wrap_text=False
                    )
                    cell.border = thin_border

        # Freeze and filter
        if ws.max_row >= 1 and ws.max_column >= 1:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

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

            ws.column_dimensions[
                get_column_letter(col_idx)
            ].width = width

        # Date / amount formatting
        for col_idx, col_name in enumerate(dataframe.columns, start=1):
            lower_name = str(col_name).strip().casefold()

            for row_idx in range(2, ws.max_row + 1):
                cell = ws.cell(
                    row=row_idx,
                    column=col_idx
                )

                if "date" in lower_name:
                    cell.number_format = "dd-mm-yyyy"

                if (
                    "amount" in lower_name
                    or "contribution" in lower_name
                    or (
                        "total" in lower_name
                        and "amount" in lower_name
                    )
                ):
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0.00'

        # Add Excel table only for normal one-header-row sheets.
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


def format_excel_section(
    ws,
    header_row,
    start_row,
    end_row,
    start_col=1,
    end_col=None
):
    """Format a separate Excel data section without creating a table."""
    try:
        from openpyxl.styles import (
            Font,
            PatternFill,
            Alignment,
            Border,
            Side
        )
        from openpyxl.utils import get_column_letter

        if end_col is None:
            end_col = ws.max_column

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
        for col in range(start_col, end_col + 1):
            cell = ws.cell(
                row=header_row,
                column=col
            )
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )
            cell.border = thin_border

        ws.row_dimensions[header_row].height = 28

        # Body
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = ws.cell(
                    row=row,
                    column=col
                )
                cell.font = body_font
                cell.alignment = Alignment(
                    vertical="center",
                    wrap_text=False
                )
                cell.border = thin_border

        # Width
        for col in range(start_col, end_col + 1):
            max_len = 0

            for row in range(header_row, end_row + 1):
                value = ws.cell(
                    row=row,
                    column=col
                ).value

                max_len = max(
                    max_len,
                    len(str(value)) if value is not None else 0
                )

            ws.column_dimensions[
                get_column_letter(col)
            ].width = min(max(max_len + 3, 12), 35)

    except Exception:
        pass


def normalize_column_key(value):
    """Create a stable key for matching Excel headings."""
    text = str(value)
    text = text.replace("\n", " ").replace("\r", " ")
    text = re.sub(r"\s+", " ", text).strip().casefold()
    text = re.sub(r"[^a-z0-9]+", "_", text).strip("_")
    return text


def normalize_columns(columns):
    """Normalize different Excel headings to common master headings."""
    result = []
    used = {}

    preferred_names = {
        # UIN / employee number
        "uin": "UIN",
        "employee_number": "UIN",
        "employee_no": "UIN",
        "employee_id": "UIN",
        "employee_code": "UIN",

        # Employee name
        "employee_name": "Employee Name",
        "name": "Employee Name",
        "employee": "Employee Name",

        # BN / Battalion / Unit
        "bn": "BN",
        "battalion": "BN",
        "battalion_no": "BN",
        "battalion_number": "BN",
        "unit": "BN",
        "unit_name": "BN",

        # PRAN
        "pran": "PRAN No",
        "pran_no": "PRAN No",
        "pran_number": "PRAN No",
        "pran_no_": "PRAN No",

        # Status
        "status": "Status",
        "case_status": "Case Status",

        # Amount / contribution
        "amount": "Amount",
        "contribution": "Contribution",
        "total_amount": "Total Amount",

        # Nomination
        "nomination": "Nomination",
        "nomination_status": "Nomination Status",
        "nominee": "Nominee",
        "nominee_status": "Nominee Status",

        # Mobile
        "mobile": "Mobile No",
        "mobile_no": "Mobile No",
        "mobile_number": "Mobile No",
        "phone": "Mobile No",
        "phone_no": "Mobile No"
    }

    for original in columns:
        original_text = str(original).strip()
        key = normalize_column_key(original_text)

        base_name = preferred_names.get(
            key,
            original_text
        )

        if base_name == "":
            base_name = "Unnamed Column"

        if base_name in used:
            used[base_name] += 1
            final_name = (
                f"{base_name}_{used[base_name]}"
            )
        else:
            used[base_name] = 1
            final_name = base_name

        result.append(final_name)

    return result


def merge_duplicate_master_columns(dataframe):
    """Merge duplicate canonical columns safely."""
    df_local = dataframe.copy()

    groups = {}

    for idx, col in enumerate(df_local.columns):
        base = str(col).strip()

        if re.match(r"^(.+)_\d+$", base):
            base_name = re.sub(r"_\d+$", "", base)
        else:
            base_name = base

        groups.setdefault(base_name, []).append(idx)

    # Rebuild columns positionally to avoid pandas duplicate-label issues.
    columns_to_drop = []

    for base_name, positions in groups.items():
        if len(positions) <= 1:
            continue

        combined = pd.Series(
            pd.NA,
            index=df_local.index,
            dtype="object"
        )

        for pos in positions:
            series = df_local.iloc[:, pos].copy()

            clean = series.replace(
                {
                    "": pd.NA,
                    "None": pd.NA,
                    "none": pd.NA,
                    "nan": pd.NA,
                    "NaN": pd.NA,
                    "null": pd.NA,
                    "NULL": pd.NA
                }
            )

            combined = combined.fillna(clean)

        first_pos = positions[0]
        df_local.iloc[:, first_pos] = combined

        columns_to_drop.extend(
            positions[1:]
        )

    if columns_to_drop:
        keep_positions = [
            idx
            for idx in range(len(df_local.columns))
            if idx not in set(columns_to_drop)
        ]

        df_local = df_local.iloc[
            :,
            keep_positions
        ].copy()

    return df_local


def find_column(columns, possible_names):
    normalized = {
        normalize_column_key(x): x
        for x in columns
    }

    for possible in possible_names:
        key = normalize_column_key(possible)

        if key in normalized:
            return normalized[key]

    return None


def status_mask(dataframe, status_col, pattern):
    if not status_col or status_col not in dataframe.columns:
        return pd.Series(
            False,
            index=dataframe.index
        )

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
            "not available"
        ])

    return data.isin(values)


def unique_employee_count(dataframe, uin_col, employee_name_col):
    if uin_col and uin_col in dataframe.columns:
        values = (
            dataframe[uin_col]
            .fillna("")
            .astype(str)
            .str.strip()
            .replace("", pd.NA)
            .dropna()
        )

        return int(values.nunique())

    if employee_name_col and employee_name_col in dataframe.columns:
        values = (
            dataframe[employee_name_col]
            .fillna("")
            .astype(str)
            .str.strip()
            .replace("", pd.NA)
            .dropna()
        )

        return int(values.nunique())

    return len(dataframe)


def duplicate_uin_dataframe(dataframe, uin_col):
    if not uin_col or uin_col not in dataframe.columns:
        return pd.DataFrame(
            columns=dataframe.columns
        )

    uin_data = (
        dataframe[uin_col]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    duplicate_mask = (
        uin_data.ne("")
        & uin_data.duplicated(keep=False)
    )

    return dataframe.loc[
        duplicate_mask
    ].copy()


def build_bn_summary(
    dataframe,
    bn_col,
    status_col,
    amount_col,
    pran_col,
    nomination_col,
    uin_col
):
    if not bn_col or bn_col not in dataframe.columns:
        return pd.DataFrame()

    rows = []

    for bn_value, group in dataframe.groupby(
        bn_col,
        dropna=False,
        sort=True
    ):
        bn_display = str(bn_value).strip()

        if bn_display.casefold() in [
            "",
            "nan",
            "none"
        ]:
            bn_display = "Blank BN"

        total = len(group)

        pending = (
            int(
                status_mask(
                    group,
                    status_col,
                    "pending"
                ).sum()
            )
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
            int(
                missing_mask(
                    group[pran_col]
                ).sum()
            )
            if pran_col and pran_col in group.columns
            else 0
        )

        missing_nomination = (
            int(
                missing_mask(
                    group[nomination_col],
                    include_no=True
                ).sum()
            )
            if nomination_col and nomination_col in group.columns
            else 0
        )

        duplicate_uin = 0

        if uin_col and uin_col in group.columns:
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

            duplicate_uin = int(
                duplicate_mask.sum()
            )

        if amount_col and amount_col in group.columns:
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


def clean_blank_rows(dataframe):
    """Remove rows where all actual data fields are blank."""
    if dataframe.empty:
        return dataframe.copy()

    check_df = dataframe.copy()

    check_df = check_df.replace(
        {
            r"^\s*$": pd.NA,
            "None": pd.NA,
            "none": pd.NA,
            "nan": pd.NA,
            "NaN": pd.NA,
            "null": pd.NA,
            "NULL": pd.NA
        },
        regex=True
    )

    blank_mask = check_df.isna().all(axis=1)

    return dataframe.loc[
        ~blank_mask
    ].reset_index(drop=True)


def calculate_completion(missing_count, total_count):
    if total_count <= 0:
        return 0.0

    value = (
        (total_count - missing_count)
        / total_count
        * 100
    )

    return round(max(0, min(100, value)), 1)


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
    st.info(
        "Please upload one or more Excel files to start."
    )
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
        excel_file = pd.ExcelFile(
            uploaded_file
        )

        for sheet_name in excel_file.sheet_names:
            temp_df = pd.read_excel(
                uploaded_file,
                sheet_name=sheet_name
            )

            if temp_df.empty:
                continue

            # Normalize headings
            temp_df.columns = normalize_columns(
                temp_df.columns
            )

            # Merge duplicate master headings
            temp_df = merge_duplicate_master_columns(
                temp_df
            )

            # Remove blank rows
            temp_df = clean_blank_rows(
                temp_df
            )

            if temp_df.empty:
                continue

            # Source tracking
            temp_df["Source File"] = (
                uploaded_file.name
            )
            temp_df["Source Sheet"] = (
                sheet_name
            )

            all_dataframes.append(
                temp_df
            )

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
    st.error(
        "No readable data was found in the uploaded Excel files."
    )
    st.stop()


# =========================================================
# FINAL MASTER DATA
# =========================================================
df = pd.concat(
    all_dataframes,
    ignore_index=True,
    sort=False
)

df.columns = normalize_columns(
    df.columns
)

df = merge_duplicate_master_columns(
    df
)

# Remove completely blank records while ignoring source fields.
data_columns = [
    col
    for col in df.columns
    if col not in [
        "Source File",
        "Source Sheet"
    ]
]

if data_columns:
    check_df = df[data_columns].copy()

    check_df = check_df.replace(
        {
            r"^\s*$": pd.NA,
            "None": pd.NA,
            "none": pd.NA,
            "nan": pd.NA,
            "NaN": pd.NA,
            "null": pd.NA,
            "NULL": pd.NA
        },
        regex=True
    )

    final_blank_mask = check_df.isna().all(axis=1)

    df = df.loc[
        ~final_blank_mask
    ].reset_index(drop=True)


st.success(
    "All uploaded files were combined successfully. "
    f"Total Records: {len(df):,}"
)


# =========================================================
# SOURCE FILE SUMMARY
# =========================================================
with st.expander("Uploaded Files Summary"):
    source_summary = pd.DataFrame(
        file_sheet_summary
    )

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
    {
        "bn",
        "battalion",
        "unit"
    }
)

status_col = find_column(
    df.columns,
    {
        "status",
        "case status",
        "case_status"
    }
)

amount_col = find_column(
    df.columns,
    {
        "amount",
        "contribution",
        "total amount"
    }
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
        lambda col: col.str.casefold().str.contains(
            re.escape(search_value),
            regex=True,
            na=False
        )
    ).any(axis=1)

    filtered_df = filtered_df.loc[
        mask
    ].copy()

    st.info(
        f"Search result: {len(filtered_df)} record(s) found."
    )


# BN filter
if bn_col:
    bn_values = sorted(
        df[bn_col]
        .dropna()
        .astype(str)
        .str.strip()
        .loc[lambda s: s.ne("")]
        .unique()
        .tolist()
    )

    selected_bn = st.selectbox(
        "Select BN",
        ["All"] + bn_values
    )

    if selected_bn != "All":
        filtered_df = filtered_df[
            filtered_df[bn_col]
            .astype(str)
            .str.strip()
            == selected_bn
        ].copy()


# Status filter
if status_col:
    status_values = sorted(
        df[status_col]
        .dropna()
        .astype(str)
        .str.strip()
        .loc[lambda s: s.ne("")]
        .unique()
        .tolist()
    )

    selected_status = st.selectbox(
        "Case Status",
        ["All"] + status_values
    )

    if selected_status != "All":
        filtered_df = filtered_df[
            filtered_df[status_col]
            .astype(str)
            .str.strip()
            == selected_status
        ].copy()


# =========================================================
# DASHBOARD
# =========================================================
st.subheader("Automatic Dashboard")

st.markdown(
    '<div class="dashboard-note">'
    'Dashboard figures are calculated from the current Search, BN and Case Status filters.'
    '</div>',
    unsafe_allow_html=True
)

total_records = len(filtered_df)

unique_employee_count = unique_employee_count(
    filtered_df,
    uin_col,
    employee_name_col
)

pending_count = (
    int(
        status_mask(
            filtered_df,
            status_col,
            "pending"
        ).sum()
    )
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

missing_pran_count = (
    int(
        missing_mask(
            filtered_df[pran_col]
        ).sum()
    )
    if pran_col and pran_col in filtered_df.columns
    else 0
)

missing_nomination_count = (
    int(
        missing_mask(
            filtered_df[nomination_col],
            include_no=True
        ).sum()
    )
    if nomination_col and nomination_col in filtered_df.columns
    else 0
)

total_amount = (
    pd.to_numeric(
        filtered_df[amount_col],
        errors="coerce"
    ).fillna(0).sum()
    if amount_col and amount_col in filtered_df.columns
    else 0
)

pran_completion = calculate_completion(
    missing_pran_count,
    total_records
)

nomination_completion = calculate_completion(
    missing_nomination_count,
    total_records
)

duplicate_uin_count = len(
    duplicate_uin_dataframe(
        filtered_df,
        uin_col
    )
)


# =========================================================
# DASHBOARD CARDS
# =========================================================
c1, c2, c3, c4 = st.columns(4)

with c1:
    st.metric(
        "Total Records",
        f"{total_records:,}"
    )

with c2:
    st.metric(
        "Unique Employees",
        f"{unique_employee_count:,}"
    )

with c3:
    st.metric(
        "Pending",
        f"{pending_count:,}"
    )

with c4:
    st.metric(
        "Completed",
        f"{completed_count:,}"
    )


c5, c6, c7, c8 = st.columns(4)

with c5:
    st.metric(
        "Missing PRAN",
        f"{missing_pran_count:,}"
    )

with c6:
    st.metric(
        "PRAN Completion",
        f"{pran_completion:.1f}%"
    )

with c7:
    st.metric(
        "Missing Nomination",
        f"{missing_nomination_count:,}"
    )

with c8:
    st.metric(
        "Nomination Completion",
        f"{nomination_completion:.1f}%"
    )


c9, c10 = st.columns(2)

with c9:
    st.metric(
        "Duplicate UIN Records",
        f"{duplicate_uin_count:,}"
    )

with c10:
    st.metric(
        "Total Contribution",
        f"INR {total_amount:,.2f}"
    )


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

    if not bn_summary.empty:

        bn_summary["Pending %"] = (
            bn_summary["Pending"]
            .div(
                bn_summary[
                    "Total Employees"
                ].replace(0, pd.NA)
            )
            .mul(100)
            .fillna(0)
            .round(1)
        )

        bn_summary["Completed %"] = (
            bn_summary["Completed"]
            .div(
                bn_summary[
                    "Total Employees"
                ].replace(0, pd.NA)
            )
            .mul(100)
            .fillna(0)
            .round(1)
        )

        st.dataframe(
            bn_summary,
            use_container_width=True,
            hide_index=True
        )

        chart_data = bn_summary.set_index(
            "BN"
        )[[
            "Total Employees",
            "Pending",
            "Completed",
            "Missing PRAN",
            "Missing Nomination"
        ]]

        st.caption(
            "BN-wise employee and case-status overview"
        )

        st.bar_chart(
            chart_data,
            use_container_width=True
        )

    else:
        st.info(
            "No BN-wise data available for the current filters."
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
        .replace("", "Blank")
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
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
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
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        key="download_employee_report"
    )

else:
    st.info(
        "No employee records found."
    )


# =========================================================
# DUPLICATE UIN
# =========================================================
duplicate_uin_report = duplicate_uin_dataframe(
    employee_report,
    uin_col
)

if uin_col:
    st.subheader("Duplicate UIN Report")

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
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            key="download_duplicate_uin"
        )

    else:
        st.success(
            "No duplicate UIN found."
        )


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
        employee_report[
            missing_pran_mask
        ].copy()
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
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            key="download_missing_pran"
        )

    else:
        st.success(
            "All employees have PRAN."
        )


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
        employee_report[
            missing_nomination_mask
        ].copy()
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
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            key="download_missing_nomination"
        )

    else:
        st.success(
            "All employees have nomination."
        )


# =========================================================
# BN-WISE COMPLETE EXCEL REPORT
# =========================================================
if bn_col:
    st.subheader(
        "BN-wise Complete Excel Report"
    )

    st.write(
        "One Excel file containing Summary, BN Summary, "
        "Filtered Data, and a separate sheet for each BN."
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

            # -------------------------------------------------
            # Summary
            # -------------------------------------------------
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

            summary_sheet = "Summary"

            summary_data.to_excel(
                writer,
                index=False,
                sheet_name=summary_sheet
            )

            format_excel_writer(
                writer,
                summary_data,
                summary_sheet
            )

            used_sheet_names.add(
                summary_sheet
            )

            # -------------------------------------------------
            # BN Summary
            # -------------------------------------------------
            bn_summary_sheet = "BN Summary"

            bn_summary.to_excel(
                writer,
                index=False,
                sheet_name=bn_summary_sheet
            )

            format_excel_writer(
                writer,
                bn_summary,
                bn_summary_sheet
            )

            used_sheet_names.add(
                bn_summary_sheet
            )

            # -------------------------------------------------
            # Filtered Data
            # -------------------------------------------------
            filtered_sheet = "Filtered Data"

            filtered_df.to_excel(
                writer,
                index=False,
                sheet_name=filtered_sheet
            )

            format_excel_writer(
                writer,
                filtered_df,
                filtered_sheet
            )

            used_sheet_names.add(
                filtered_sheet
            )

            # -------------------------------------------------
            # One sheet for each BN
            # -------------------------------------------------
            for bn_value, group in filtered_df.groupby(
                bn_col,
                dropna=False,
                sort=True
            ):
                bn_display = str(
                    bn_value
                ).strip()

                if bn_display.casefold() in [
                    "",
                    "nan",
                    "none"
                ]:
                    bn_display = "Blank BN"

                sheet_name = safe_sheet_name(
                    f"BN-{bn_display}",
                    used_sheet_names
                )

                used_sheet_names.add(
                    sheet_name
                )

                mini_summary = build_bn_summary(
                    group,
                    bn_col,
                    status_col,
                    amount_col,
                    pran_col,
                    nomination_col,
                    uin_col
                )

                if not mini_summary.empty:
                    mini_row = (
                        mini_summary.iloc[0]
                    )

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

                # Mini summary
                mini.to_excel(
                    writer,
                    index=False,
                    sheet_name=sheet_name,
                    startrow=0
                )

                # Full BN data below summary
                data_start_row = len(mini) + 3

                group.to_excel(
                    writer,
                    index=False,
                    sheet_name=sheet_name,
                    startrow=data_start_row - 1
                )

                ws = writer.book[
                    sheet_name
                ]

                # Format mini summary
                format_excel_section(
                    ws,
                    header_row=1,
                    start_row=2,
                    end_row=len(mini) + 1,
                    start_col=1,
                    end_col=2
                )

                # Format full data
                full_header_row = (
                    data_start_row
                )

                full_end_row = (
                    data_start_row
                    + len(group)
                )

                format_excel_section(
                    ws,
                    header_row=full_header_row,
                    start_row=full_header_row + 1,
                    end_row=full_end_row,
                    start_col=1,
                    end_col=len(group.columns)
                )

                ws.freeze_panes = (
                    f"A{full_header_row + 1}"
                )

        bn_output.seek(0)

        st.success(
            "BN-wise Complete Excel Report is ready."
        )

        st.download_button(
            "Download BN-wise Complete Excel Report",
            data=bn_output.getvalue(),
            file_name="BN_Wise_Complete_Report.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            key="download_bn_complete_report"
        )


# =========================================================
# COMPLETE MASTER REPORT
# =========================================================
st.subheader(
    "Complete Master Report"
)

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
    ):

        # -------------------------------------------------
        # Summary
        # -------------------------------------------------
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

        # -------------------------------------------------
        # BN Summary
        # -------------------------------------------------
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

        # -------------------------------------------------
        # Employee Report
        # -------------------------------------------------
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

        # -------------------------------------------------
        # Missing PRAN
        # -------------------------------------------------
        if pran_col:
            pran_export = (
                missing_pran_report
            )
        else:
            pran_export = pd.DataFrame({
                "Message": [
                    "PRAN column not found"
                ]
            })

        pran_export.to_excel(
            writer,
            index=False,
            sheet_name="Missing PRAN"
        )

        format_excel_writer(
            writer,
            pran_export,
            "Missing PRAN"
        )

        # -------------------------------------------------
        # Missing Nomination
        # -------------------------------------------------
        if nomination_col:
            nomination_export = (
                missing_nomination_report
            )
        else:
            nomination_export = pd.DataFrame({
                "Message": [
                    "Nomination column not found"
                ]
            })

        nomination_export.to_excel(
            writer,
            index=False,
            sheet_name="Missing Nomination"
        )

        format_excel_writer(
            writer,
            nomination_export,
            "Missing Nomination"
        )

        # -------------------------------------------------
        # Duplicate UIN
        # -------------------------------------------------
        if uin_col:
            duplicate_export = (
                duplicate_uin_report
            )
        else:
            duplicate_export = pd.DataFrame({
                "Message": [
                    "UIN column not found"
                ]
            })

        duplicate_export.to_excel(
            writer,
            index=False,
            sheet_name="Duplicate UIN"
        )

        format_excel_writer(
            writer,
            duplicate_export,
            "Duplicate UIN"
        )

        # -------------------------------------------------
        # Case Status
        # -------------------------------------------------
        if status_col:
            status_export = (
                status_summary
            )
        else:
            status_export = pd.DataFrame({
                "Message": [
                    "Status column not found"
                ]
            })

        status_export.to_excel(
            writer,
            index=False,
            sheet_name="Case Status"
        )

        format_excel_writer(
            writer,
            status_export,
            "Case Status"
        )

        # -------------------------------------------------
        # Filtered Data
        # -------------------------------------------------
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
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        key="download_complete_master_report"
    )


# =========================================================
# DATA PREVIEW
# =========================================================
st.subheader("Data Preview")

st.dataframe(
    filtered_df,
    use_container_width=True,
    height=450,
    hide_index=True
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
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
    key="download_filtered_report"
)
