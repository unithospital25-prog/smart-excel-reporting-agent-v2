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
        r'[\:\*\?\/\\]',
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
            lambda
