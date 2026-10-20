import streamlit as st
import pandas as pd
import re
import logging
from io import BytesIO

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="Smart Excel Reporting Agent V5",
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

.dashboard-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(155px, 1fr));
    gap: 12px;
    width: 100%;
    margin: 8px 0 26px 0;
}

.dashboard-card {
    box-sizing: border-box;
    min-height: 104px;
    padding: 15px;
    border: 1px solid rgba(148,163,184,0.28);
    border-radius: 14px;
    background: linear-gradient(
        145deg,
        rgba(255,255,255,0.075),
        rgba(255,255,255,0.025)
    );
    box-shadow: 0 5px 18px rgba(0,0,0,0.12);
}

.dashboard-top {
    display: flex;
    align-items: center;
    gap: 8px;
}

.dashboard-dot {
    width: 9px;
    height: 9px;
    border-radius: 50%;
    background: currentColor;
}

.dashboard-label {
    font-size: 13px;
    font-weight: 600;
    color: #a7afbd;
}

.dashboard-value {
    margin-top: 10px;
    font-size: 30px;
    font-weight: 800;
    color: #f8fafc;
}

.dashboard-blue { color: #60a5fa; }
.dashboard-purple { color: #a78bfa; }
.dashboard-orange { color: #f59e0b; }
.dashboard-green { color: #4ade80; }
.dashboard-red { color: #f87171; }
.dashboard-pink { color: #f472b6; }
.dashboard-teal { color: #2dd4bf; }

@media (max-width: 520px) {
    .dashboard-grid {
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 9px;
    }

    .dashboard-card {
        min-height: 88px;
        padding: 11px;
    }

    .dashboard-label {
        font-size: 11px;
    }

    .dashboard-value {
        font-size: 22px;
    }
}
</style>
""", unsafe_allow_html=True)

st.markdown(
    '<div class="main-title">Smart Excel Reporting Agent V5</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="sub-title">'
    'Upload, analyze and generate smart Excel reports - single or multiple files.'
    '</div>',
    unsafe_allow_html=True
)


# =========================================================
# EXCEL FUNCTIONS
# =========================================================

def format_excel_writer(writer, dataframe, sheet_name):

    try:
        from openpyxl.styles import (
            Font,
            PatternFill,
            Alignment,
            Border,
            Side
        )
        from openpyxl.utils import get_column_letter

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
            size=10
        )

        border = Border(
            left=Side(style="thin", color="D9E1F2"),
            right=Side(style="thin", color="D9E1F2"),
            top=Side(style="thin", color="D9E1F2"),
            bottom=Side(style="thin", color="D9E1F2")
        )

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )
            cell.border = border

        ws.row_dimensions[1].height = 28
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for row in ws.iter_rows(min_row=2):

            for cell in row:
                cell.font = body_font
                cell.alignment = Alignment(
                    vertical="center"
                )
                cell.border = border

        for col_idx, cells in enumerate(
            ws.iter_cols(
                min_row=1,
                max_row=ws.max_row,
                min_col=1,
                max_col=ws.max_column
            ),
            start=1
        ):

            max_length = 0

            for cell in cells:
                value = "" if cell.value is None else str(cell.value)
                max_length = max(max_length, len(value))

            ws.column_dimensions[
                get_column_letter(col_idx)
            ].width = min(
                max(max_length + 3, 12),
                35
            )

    except Exception as e:

        logger.warning(
            "Excel formatting failed: %s",
            e
        )


def make_excel(dataframe, sheet_name="Report"):

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

        format_excel_writer(
            writer,
            dataframe,
            sheet_name
        )

    output.seek(0)

    return output.getvalue()


def safe_sheet_name(name, used_names=None):

    name = str(name)

    name = re.sub(
        r'[\[\]\:\*\?\/\\]',
        "_",
        name
    )

    name = name[:31] or "Sheet"

    if used_names is None:
        return name

    original = name
    counter = 1

    while name in used_names:

        suffix = f"_{counter}"

        name = (
            original[:31 - len(suffix)]
            + suffix
        )

        counter += 1

    return name


# =========================================================
# CLEANING
# =========================================================

def clean_empty_rows(
    dataframe,
    exclude_columns=None
):

    if dataframe is None or dataframe.empty:
        return dataframe

    df_clean = dataframe.copy()

    if exclude_columns is None:
        exclude_columns = []

    fake_empty_values = {
        "",
        "none",
        "nan",
        "null",
        "n/a",
        "na",
        "nat",
        "-"
    }

    for col in df_clean.columns:

        if col in exclude_columns:
            continue

        if (
            pd.api.types.is_object_dtype(df_clean[col])
            or
            pd.api.types.is_string_dtype(df_clean[col])
        ):

            text = (
                df_clean[col]
                .astype("string")
                .str.replace(
                    "\u00a0",
                    " ",
                    regex=False
                )
                .str.strip()
            )

            blank_mask = (
                text.str.casefold()
                .isin(fake_empty_values)
            )

            df_clean[col] = text.mask(
                blank_mask,
                pd.NA
            )

    data_columns = [
        col
        for col in df_clean.columns
        if col not in exclude_columns
    ]

    if data_columns:

        meaningful_mask = (
            df_clean[data_columns]
            .notna()
            .any(axis=1)
        )

        df_clean = df_clean.loc[
            meaningful_mask
        ].copy()

    return df_clean.reset_index(
        drop=True
    )


# =========================================================
# COLUMN NORMALIZATION
# =========================================================

def normalize_column_key(value):

    text = (
        str(value)
        .replace("\n", " ")
        .replace("\r", " ")
    )

    text = re.sub(
        r"\s+",
        " ",
        text
    ).strip().casefold()

    text = re.sub(
        r"[^a-z0-9]+",
        "_",
        text
    ).strip("_")

    return text


def normalize_columns(columns):

    preferred_names = {

        "uin": "UIN",
        "employee_number": "UIN",
        "employee_no": "UIN",
        "employee_id": "UIN",
        "employee_code": "UIN",

        "employee_name": "Employee Name",
        "name": "Employee Name",
        "employee": "Employee Name",

        "bn": "BN",
        "battalion": "BN",
        "battalion_no": "BN",
        "bn_no": "BN",
        "unit": "BN",

        "pran": "PRAN",
        "pran_no": "PRAN",
        "pran_number": "PRAN",
        "pran_num": "PRAN",

        "status": "Status",
        "case_status": "Status",

        "amount": "Amount",
        "contribution": "Amount",
        "total_amount": "Amount",

        "nomination": "Nomination",
        "nomination_status": "Nomination",
        "nominee": "Nomination",
        "nominee_status": "Nomination",

        "mobile": "Mobile",
        "mobile_no": "Mobile"
    }

    result = []
    used = {}

    for original in columns:

        original_text = str(
            original
        ).strip()

        key = normalize_column_key(
            original_text
        )

        base_name = preferred_names.get(
            key,
            original_text
        )

        if not base_name:
            base_name = "Unnamed Column"

        if base_name in used:

            used[base_name] += 1

            final_name = (
                f"{base_name}_{used[base_name]}"
            )

        else:

            used[base_name] = 1

            final_name = base_name

        result.append(
            final_name
        )

    return result


def find_column(
    columns,
    possible_names
):

    exact = {
        str(x).strip().casefold(): x
        for x in columns
    }

    normalized = {
        normalize_column_key(x): x
        for x in columns
    }

    for possible in possible_names:

        key = str(
            possible
        ).strip().casefold()

        if key in exact:
            return exact[key]

    for possible in possible_names:

        key = normalize_column_key(
            possible
        )

        if key in normalized:
            return normalized[key]

    candidates = []

    aliases = [
        normalize_column_key(x)
        for x in possible_names
    ]

    for col in columns:

        col_key = normalize_column_key(
            col
        )

        for alias in aliases:

            if alias and (
                col_key.startswith(
                    alias + "_"
                )
                or
                col_key.endswith(
                    "_" + alias
                )
            ):

                candidates.append(col)
                break

    candidates = list(
        dict.fromkeys(candidates)
    )

    if len(candidates) == 1:
        return candidates[0]

    return None


# =========================================================
# STATUS / MISSING FUNCTIONS
# =========================================================

def classify_status(series):

    data = (
        series
        .fillna("")
        .astype(str)
        .str.strip()
        .str.casefold()
    )

    pending_pattern = (
        r"pending|under\s*process|"
        r"in\s*progress|awaiting|"
        r"under\s*review|returned|"
        r"objection"
    )

    completed_pattern = (
        r"complete|completed|done|"
        r"approved|settled|paid|closed"
    )

    pending = data.str.contains(
        pending_pattern,
        regex=True,
        na=False
    )

    completed = data.str.contains(
        completed_pattern,
        regex=True,
        na=False
    )

    return pending, completed


def status_mask(
    dataframe,
    status_col,
    pattern
):

    if (
        not status_col
        or status_col not in dataframe.columns
    ):

        return pd.Series(
            False,
            index=dataframe.index
        )

    if pattern == "pending":

        return classify_status(
            dataframe[status_col]
        )[0]

    if pattern == "complete|completed|done":

        return classify_status(
            dataframe[status_col]
        )[1]

    values = (
        dataframe[status_col]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.casefold()
    )

    return values.str.contains(
        pattern,
        regex=True,
        na=False
    )


def missing_mask(
    series,
    include_no=False
):

    values = (
        series
        .fillna("")
        .astype(str)
        .str.strip()
        .str.casefold()
    )

    missing_values = [
        "",
        "none",
        "nan",
        "null",
        "na",
        "n/a",
        "-"
    ]

    if include_no:

        missing_values.extend([
            "no",
            "pending",
            "not available"
        ])

    return values.isin(
        missing_values
    )


# =========================================================
# BN SUMMARY
# =========================================================

def build_bn_summary(
    dataframe,
    bn_col,
    status_col,
    amount_col,
    pran_col,
    nomination_col,
    uin_col
):

    if (
        not bn_col
        or bn_col not in dataframe.columns
    ):

        return pd.DataFrame()

    rows = []

    for bn_value, group in dataframe.groupby(
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

        total = len(group)

        pending = (
            int(
                status_mask(
                    group,
                    status_col,
                    "pending"
                ).sum()
            )
            if status_col
            else 0
        )

        completed = (
            int(
                status_mask(
                    group,
                    status_col,
                    "complete|completed|done"
                ).sum()
            )
            if status_col
            else 0
        )

        missing_pran = (
            int(
                missing_mask(
                    group[pran_col]
                ).sum()
            )
            if pran_col
            else 0
        )

        missing_nomination = (
            int(
                missing_mask(
                    group[nomination_col],
                    include_no=True
                ).sum()
            )
            if nomination_col
            else 0
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
                &
                uin_data.duplicated(
                    keep=False
                )
            )

            duplicate_uin = int(
                duplicate_mask.sum()
            )

        if amount_col:

            contribution = (
                pd.to_numeric(
                    group[amount_col],
                    errors="coerce"
                )
                .fillna(0)
                .sum()
            )

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
# UPLOAD
# =========================================================

st.caption(
    "V5: smart column detection, broader status classification, "
    "duplicate-UIN metrics and enhanced KPIs."
)

uploaded_files = st.file_uploader(
    "Upload Excel File(s) - Multiple files allowed",
    type=["xlsx", "xls"],
    accept_multiple_files=True
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
# COMBINE FILES + SHEETS
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

            temp_df.columns = normalize_columns(
                temp_df.columns
            )

            temp_df = clean_empty_rows(
                temp_df
            )

            if temp_df.empty:
                continue

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

                "Source File":
                    uploaded_file.name,

                "Source Sheet":
                    sheet_name,

                "Records":
                    len(temp_df)
            })

    except Exception as e:

        st.warning(
            f"Unable to read "
            f"{uploaded_file.name}: {e}"
        )


if not all_dataframes:

    st.error(
        "No readable data was found "
        "in the uploaded Excel files."
    )

    st.stop()


df = pd.concat(
    all_dataframes,
    ignore_index=True,
    sort=False
)

df.columns = normalize_columns(
    df.columns
)

df = clean_empty_rows(
    df,
    exclude_columns=[
        "Source File",
        "Source Sheet"
    ]
)

st.success(
    "All uploaded files were combined successfully. "
    f"Total Records: {len(df):,}"
)


# =========================================================
# SOURCE SUMMARY
# =========================================================

with st.expander(
    "Uploaded Files Summary"
):

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
# VALID RECORD FILTER
# =========================================================

primary_columns = [
    uin_col,
    employee_name_col,
    bn_col
]

primary_columns = [
    col
    for col in primary_columns
    if col is not None
    and col in df.columns
]

secondary_columns = [
    pran_col,
    nomination_col,
    status_col,
    amount_col
]

secondary_columns = [
    col
    for col in secondary_columns
    if col is not None
    and col in df.columns
]

empty_tokens = [
    "",
    "none",
    "nan",
    "null",
    "n/a",
    "na",
    "nat",
    "-"
]

if primary_columns:

    check = pd.DataFrame(
        index=df.index
    )

    for col in primary_columns:

        values = (
            df[col]
            .astype("string")
            .str.replace(
                "\u00a0",
                " ",
                regex=False
            )
            .str.strip()
            .str.casefold()
        )

        check[col] = (
            values.notna()
            &
            ~values.isin(
                empty_tokens
            )
        )

    valid_mask = check.any(
        axis=1
    )

    df = df.loc[
        valid_mask
    ].copy()

elif secondary_columns:

    check = pd.DataFrame(
        index=df.index
    )

    for col in secondary_columns:

        values = (
            df[col]
            .astype("string")
            .str.replace(
                "\u00a0",
                " ",
                regex=False
            )
            .str.strip()
            .str.casefold()
        )

        check[col] = (
            values.notna()
            &
            ~values.isin(
                empty_tokens
            )
        )

    valid_mask = check.any(
        axis=1
    )

    df = df.loc[
        valid_mask
    ].copy()


df = clean_empty_rows(
    df,
    exclude_columns=[
        "Source File",
        "Source Sheet"
    ]
)

df.reset_index(
    drop=True,
    inplace=True
)


# =========================================================
# UNIQUE EMPLOYEE MASTER
# =========================================================

raw_df = df.copy()

if uin_col:

    uin_clean = (
        raw_df[uin_col]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    valid_uin = uin_clean.ne("")

    master_df = raw_df.loc[
        valid_uin
    ].copy()

    master_df = (
        master_df
        .drop_duplicates(
            subset=[uin_col],
            keep="first"
        )
        .reset_index(drop=True)
    )

else:

    master_df = (
        raw_df
        .drop_duplicates()
        .reset_index(drop=True)
    )


# =========================================================
# SEARCH & FILTER
# =========================================================

st.subheader(
    "Search and Filter"
)

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

    search_mask = search_data.apply(
        lambda col:
        col.str.casefold()
        .str.contains(
            re.escape(search_value),
            regex=True,
            na=False
        )
    ).any(axis=1)

    filtered_df = filtered_df.loc[
        search_mask
    ].copy()

    st.info(
        f"Search result: "
        f"{len(filtered_df)} record(s) found."
    )


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
            filtered_df[bn_col]
            .astype(str)
            == selected_bn
        ].copy()

else:

    selected_bn = "All"


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
            filtered_df[status_col]
            .astype(str)
            == selected_status
        ].copy()

else:

    selected_status = "All"


# =========================================================
# MASTER FILTER
# =========================================================

master_filtered_df = master_df.copy()


if search_text.strip():

    search_value = (
        search_text
        .strip()
        .casefold()
    )

    master_search_data = (
        master_filtered_df
        .fillna("")
        .astype(str)
    )

    master_search_mask = (
        master_search_data.apply(
            lambda col:
            col.str.casefold()
            .str.contains(
                re.escape(search_value),
                regex=True,
                na=False
            )
        )
        .any(axis=1)
    )

    master_filtered_df = (
        master_filtered_df.loc[
            master_search_mask
        ].copy()
    )


if (
    bn_col
    and selected_bn != "All"
):

    master_filtered_df = (
        master_filtered_df[
            master_filtered_df[bn_col]
            .astype(str)
            == selected_bn
        ].copy()
    )


if (
    status_col
    and selected_status != "All"
):

    master_filtered_df = (
        master_filtered_df[
            master_filtered_df[status_col]
            .astype(str)
            == selected_status
        ].copy()
    )


# =========================================================
# UNIQUE EMPLOYEE MASTER VIEW
# =========================================================

st.subheader(
    "Unique Employee Master"
)

st.caption(
    "One unique record per UIN. "
    "Duplicate UIN records remain available "
    "in the Duplicate UIN Report."
)

st.info(
    "Unique Employee Master: "
    f"{len(master_filtered_df):,} employee(s)"
)

st.dataframe(
    master_filtered_df,
    use_container_width=True,
    hide_index=True,
    height=400
)

master_excel = make_excel(
    master_filtered_df,
    "Employee Master"
)

st.download_button(
    "Download Unique Employee Master",
    data=master_excel,
    file_name="Unique_Employee_Master.xlsx",
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
    key="download_unique_employee_master"
)


# =========================================================
# DASHBOARD
# =========================================================

st.subheader(
    "Automatic Dashboard"
)

total_records = len(
    filtered_df
)

unique_employee_count = len(
    master_filtered_df
)


pending_count = (
    int(
        status_mask(
            master_filtered_df,
            status_col,
            "pending"
        ).sum()
    )
    if status_col
    else 0
)


completed_count = (
    int(
        status_mask(
            master_filtered_df,
            status_col,
            "complete|completed|done"
        ).sum()
    )
    if status_col
    else 0
)


total_amount = (
    pd.to_numeric(
        master_filtered_df[amount_col],
        errors="coerce"
    )
    .fillna(0)
    .sum()
    if amount_col
    else 0
)


missing_pran_count = (
    int(
        missing_mask(
            master_filtered_df[pran_col]
        ).sum()
    )
    if pran_col
    else 0
)


missing_nomination_count = (
    int(
        missing_mask(
            master_filtered_df[nomination_col],
            include_no=True
        ).sum()
    )
    if nomination_col
    else 0
)


pran_completion_pct = (
    round(
        (
            (
                unique_employee_count
                - missing_pran_count
            )
            / unique_employee_count
        )
        * 100,
        1
    )
    if unique_employee_count
    else 0
)


nomination_completion_pct = (
    round(
        (
            (
                unique_employee_count
                - missing_nomination_count
            )
            / unique_employee_count
        )
        * 100,
        1
    )
    if unique_employee_count
    else 0
)


# =========================================================
# DUPLICATE UIN METRIC
# =========================================================

if uin_col:

    dashboard_uin = (
        filtered_df[uin_col]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    duplicate_mask = (
        dashboard_uin.ne("")
        &
        dashboard_uin.duplicated(
            keep=False
        )
    )

    duplicate_uin_unique_count = int(
        dashboard_uin[
            duplicate_mask
        ].nunique()
    )

else:

    duplicate_uin_unique_count = 0


pending_pct = (
    round(
        (
            pending_count
            / unique_employee_count
        )
        * 100,
        1
    )
    if unique_employee_count
    else 0
)


completed_pct = (
    round(
        (
            completed_count
            / unique_employee_count
        )
        * 100,
        1
    )
    if unique_employee_count
    else 0
)


# =========================================================
# DASHBOARD CARDS
# =========================================================

card_data = [

    (
        "Total Records",
        f"{total_records:,}",
        "dashboard-blue"
    ),

    (
        "Unique Employees",
        f"{unique_employee_count:,}",
        "dashboard-purple"
    ),

    (
        "Pending",
        f"{pending_count:,}",
        "dashboard-orange"
    ),

    (
        "Completed",
        f"{completed_count:,}",
        "dashboard-green"
    ),

    (
        "Missing PRAN",
        f"{missing_pran_count:,}",
        "dashboard-red"
    ),

    (
        "PRAN Completion",
        f"{pran_completion_pct:.1f}%",
        "dashboard-teal"
    ),

    (
        "Missing Nomination",
        f"{missing_nomination_count:,}",
        "dashboard-pink"
    ),

    (
        "Nomination Completion",
        f"{nomination_completion_pct:.1f}%",
        "dashboard-blue"
    ),

    (
        "Duplicate UIN",
        f"{duplicate_uin_unique_count:,}",
        "dashboard-red"
    ),

    (
        "Total Contribution",
        f"INR {total_amount:,.0f}",
        "dashboard-teal"
    )
]


card_html = ""

for label, value, tone in card_data:

    card_html += f"""
    <div class="dashboard-card {tone}">
        <div class="dashboard-top">
            <span class="dashboard-dot"></span>
            <span class="dashboard-label">
                {label}
            </span>
        </div>

        <div class="dashboard-value">
            {value}
        </div>
    </div>
    """


st.markdown(
    f"""
    <div class="dashboard-grid">
        {card_html}
    </div>
    """,
    unsafe_allow_html=True
)


# =========================================================
# DASHBOARD CHARTS
# =========================================================

st.markdown(
    "### Dashboard Charts"
)

chart_col1, chart_col2 = st.columns(2)


# ---------------- BN CHART ----------------

with chart_col1:

    st.markdown(
        "**BN-wise Employee Count**"
    )

    if bn_col:

        bn_chart_values = (
            master_filtered_df[bn_col]
            .astype("string")
            .str.strip()
        )

        bn_chart_values = (
            bn_chart_values.mask(
                bn_chart_values
                .str.casefold()
                .isin([
                    "",
                    "none",
                    "nan",
                    "null",
                    "n/a",
                    "na"
                ]),
                "Blank BN"
            )
        )

        bn_chart = (
            bn_chart_values
            .fillna("Blank BN")
            .value_counts()
            .rename_axis("BN")
            .reset_index(
                name="Employees"
            )
            .set_index("BN")
        )

        if not bn_chart.empty:

            st.bar_chart(
                bn_chart,
                use_container_width=True
            )

        else:

            st.info(
                "No BN data available."
            )

    else:

        st.info(
            "BN column not found."
        )


# ---------------- STATUS CHART ----------------

with chart_col2:

    st.markdown(
        "**Case Status**"
    )

    if status_col:

        status_chart_values = (
            master_filtered_df[status_col]
            .astype("string")
            .str.strip()
        )

        status_chart_values = (
            status_chart_values.mask(
                status_chart_values
                .str.casefold()
                .isin([
                    "",
                    "none",
                    "nan",
                    "null",
                    "n/a",
                    "na"
                ]),
                "Blank"
            )
        )

        status_chart = (
            status_chart_values
            .fillna("Blank")
            .value_counts()
            .rename_axis("Status")
            .reset_index(
                name="Employees"
            )
            .set_index("Status")
        )

        if not status_chart.empty:

            st.bar_chart(
                status_chart,
                use_container_width=True
            )

        else:

            st.info(
                "No status data available."
            )

    else:

        st.info(
            "Case Status column not found."
        )


# =========================================================
# BN SUMMARY
# =========================================================

bn_summary = pd.DataFrame()

if bn_col:

    st.subheader(
        "BN-wise Summary"
    )

    bn_summary = build_bn_summary(
        master_filtered_df,
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

    st.subheader(
        "Case Status Summary"
    )

    status_summary = (
        master_filtered_df[
            status_col
        ]
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

    st.caption(
        f"Pending rate: "
        f"{pending_pct:.1f}% | "
        f"Completed rate: "
        f"{completed_pct:.1f}%"
    )


# =========================================================
# EMPLOYEE REPORT
# =========================================================

st.subheader(
    "Employee-wise Report"
)

employee_report = (
    master_filtered_df.copy()
)

st.write(
    "Total Unique Employees: "
    f"**{len(employee_report):,}**"
)

if not employee_report.empty:

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
# DUPLICATE UIN REPORT
# IMPORTANT: RAW FILTERED DATA USED HERE
# =========================================================

duplicate_uin_report = pd.DataFrame()

if uin_col:

    st.subheader(
        "Duplicate UIN Report"
    )

    uin_data = (
        filtered_df[uin_col]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    duplicate_mask = (
        uin_data.ne("")
        &
        uin_data.duplicated(
            keep=False
        )
    )

    duplicate_uin_report = (
        filtered_df.loc[
            duplicate_mask
        ].copy()
    )

    duplicate_uin_unique_count = int(
        uin_data[
            duplicate_mask
        ].nunique()
    )

    if not duplicate_uin_report.empty:

        st.warning(
            "Duplicate UIN Records: "
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
# UNIQUE EMPLOYEE BASIS
# =========================================================

missing_pran_report = pd.DataFrame()

if pran_col:

    st.subheader(
        "Missing PRAN Report"
    )

    missing_pran_mask = missing_mask(
        employee_report[pran_col]
    )

    missing_pran_report = (
        employee_report.loc[
            missing_pran_mask
        ].copy()
    )

    if not missing_pran_report.empty:

        st.warning(
            "Employees without PRAN: "
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
# UNIQUE EMPLOYEE BASIS
# =========================================================

missing_nomination_report = pd.DataFrame()

if nomination_col:

    st.subheader(
        "Missing Nomination Report"
    )

    missing_nomination_mask = missing_mask(
        employee_report[
            nomination_col
        ],
        include_no=True
    )

    missing_nomination_report = (
        employee_report.loc[
            missing_nomination_mask
        ].copy()
    )

    if not missing_nomination_report.empty:

        st.warning(
            "Employees with missing nomination: "
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
# BN REPORT
# =========================================================

if bn_col:

    st.subheader(
        "BN Report"
    )

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
# BN-WISE COMPLETE EXCEL
# =========================================================

if bn_col:

    st.subheader(
        "BN-wise Complete Excel Report"
    )

    st.write(
        "One Excel file containing Summary, "
        "BN Summary, Filtered Data and "
        "separate sheets for each BN."
    )

    if st.button(
        "Generate BN-wise Complete Excel Report",
        key="generate_bn_complete_report"
    ):

        output = BytesIO()
        used_names = set()

        with pd.ExcelWriter(
            output,
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
                    "Duplicate UIN Records",
                    "Unique Duplicate UINs"
                ],

                "Value": [

                    total_records,
                    pending_count,
                    completed_count,
                    total_amount,
                    unique_employee_count,
                    len(missing_pran_report),
                    len(missing_nomination_report),
                    len(duplicate_uin_report),
                    duplicate_uin_unique_count
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

            used_names.add(
                "Summary"
            )


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

            used_names.add(
                "BN Summary"
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

            used_names.add(
                "Filtered Data"
            )


            # UNIQUE DATA IN BN SHEETS
            for bn_value, group in (
                master_filtered_df.groupby(
                    bn_col,
                    dropna=False,
                    sort=True
                )
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
                    used_names
                )

                used_names.add(
                    sheet_name
                )

                group.to_excel(
                    writer,
                    index=False,
                    sheet_name=sheet_name
                )

                format_excel_writer(
                    writer,
                    group,
                    sheet_name
                )

        output.seek(0)

        st.success(
            "BN-wise Complete Excel Report is ready."
        )

        st.download_button(
            "Download BN-wise Complete Excel Report",
            data=output.getvalue(),
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
    "Generate all important reports "
    "in a single Excel file."
)

if st.button(
    "Generate Complete Master Report",
    key="generate_master_report"
):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        # ---------------- SUMMARY ----------------

        summary_data = pd.DataFrame({

            "Particular": [

                "Total Records",
                "Pending",
                "Completed",
                "Total Contribution",
                "Unique Employees",
                "Missing PRAN",
                "Missing Nomination",
                "Duplicate UIN Records",
                "Unique Duplicate UINs"
            ],

            "Value": [

                total_records,
                pending_count,
                completed_count,
                total_amount,
                unique_employee_count,
                len(missing_pran_report),
                len(missing_nomination_report),
                len(duplicate_uin_report),
                duplicate_uin_unique_count
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


        # ---------------- BN SUMMARY ----------------

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


        # ---------------- EMPLOYEE MASTER ----------------

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


        # ---------------- MISSING PRAN ----------------

        missing_pran_report.to_excel(
            writer,
            index=False,
            sheet_name="Missing PRAN"
        )

        format_excel_writer(
            writer,
            missing_pran_report,
            "Missing PRAN"
        )


        # ---------------- MISSING NOMINATION ----------------

        missing_nomination_report.to_excel(
            writer,
            index=False,
            sheet_name="Missing Nomination"
        )

        format_excel_writer(
            writer,
            missing_nomination_report,
            "Missing Nomination"
        )


        # ---------------- DUPLICATE UIN ----------------

        duplicate_uin_report.to_excel(
            writer,
            index=False,
            sheet_name="Duplicate UIN"
        )

        format_excel_writer(
            writer,
            duplicate_uin_report,
            "Duplicate UIN"
        )


        # ---------------- CASE STATUS ----------------

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


        # ---------------- FILTERED RAW DATA ----------------

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


        # ---------------- UNIQUE MASTER ----------------

        master_filtered_df.to_excel(
            writer,
            index=False,
            sheet_name="Unique Employee Master"
        )

        format_excel_writer(
            writer,
            master_filtered_df,
            "Unique Employee Master"
        )


    output.seek(0)

    st.success(
        "Complete Master Report is ready."
    )

    st.download_button(
        "Download Complete Master Report",
        data=output.getvalue(),
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

st.subheader(
    "Data Preview"
)

st.dataframe(
    filtered_df,
    use_container_width=True,
    height=450
)


# =========================================================
# FILTERED DOWNLOAD
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
