from __future__ import annotations

import io
import math
import os
import re
import sys
from datetime import datetime
from dataclasses import asdict, dataclass, field
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image
import streamlit as st

# ============================================================
# 1. Streamlit Page Config & Custom Styling (Original UI)
# ============================================================
st.set_page_config(
    page_title="WIN-SQUARE | Requirement Sheet Engine",
    layout="wide",
    page_icon="🪟",
    initial_sidebar_state="expanded"
)

# Custom Styling (Restored Original Dashboard Look)
st.markdown("""
    <style>
    header[data-testid="stHeader"] {
        z-index: 99999 !important;
        background: transparent !important;
    }

    button[data-testid="stSidebarCollapsedControl"],
    button[data-testid="stSidebarNavCollapseButton"] {
        display: flex !important;
        visibility: visible !important;
        opacity: 1 !important;
        background-color: #FF4B4B !important;
        color: white !important;
        border-radius: 8px !important;
        position: fixed !important;
        top: 12px !important;
        left: 12px !important;
        z-index: 999999 !important;
        box-shadow: 0px 3px 8px rgba(0,0,0,0.3) !important;
    }

    [data-testid="stStatusWidget"],
    #MainMenu, 
    footer {
        display: none !important;
        visibility: hidden !important;
    }
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
        background-color: #f4f6f9;
        color: #334155;
    }

    .main .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        max-width: 98%;
    }

    .hero-container {
        background: #ffffff;
        border-radius: 12px;
        padding: 20px 24px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
        margin-bottom: 20px;
    }

    .hero-title-text {
        font-size: 22px;
        font-weight: 800;
        color: #0f172a;
    }

    .hero-sub-text {
        font-size: 13px;
        color: #64748b;
    }
    </style>
""", unsafe_allow_html=True)

if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

def get_image_path(filename):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.abspath("."), filename)

# Sidebar
with st.sidebar:
    logo_file = get_image_path("logo.png")
    if os.path.exists(logo_file):
        col_s1, col_s2, col_s3 = st.columns([1, 2, 1])
        with col_s2:
            st.image(Image.open(logo_file), width=110)
    else:
        st.markdown("<h2 style='text-align: center; color:#1e293b;'><b>WIN-SQUARE</b></h2>", unsafe_allow_html=True)
    
    st.markdown("---")
    st.markdown("<b>💡 Quick Guide</b>", unsafe_allow_html=True)
    st.markdown("1. BOQ Files अपलोड करा.<br>2. Merge & Process क्लिक करा.<br>3. Exact Excel Requirement Sheet मिळवा.", unsafe_allow_html=True)

# Main Banner Dashboard
st.markdown(
    """
    <div class="hero-container">
        <div class="hero-title-text">Requirement Sheet Engine</div>
        <div class="hero-sub-text">Automated BOQ Processing & Exact Glass Formatting</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# 2. Parsing Logic & Business Calculations
# ============================================================
HEADER_SCAN_LIMIT = 200
HEADER_REMOVE_PATTERN = r"[^A-Z0-9]"

KEYWORDS = {
    "CODE": ["CODE"],
    "WIDTH": [
        ["GL", "W"], ["GLS", "W"], ["GLASS", "W"],
        ["GL", "WIDTH"], ["GLS", "WIDTH"], ["GLASS", "WIDTH"],
        ["S", "GLS", "W"], ["SGLS", "W"], ["S", "GL", "W"],
    ],
    "HEIGHT": [
        ["GL", "H"], ["GLS", "H"], ["GLASS", "H"],
        ["GL", "HEIGHT"], ["GLS", "HEIGHT"], ["GLASS", "HEIGHT"],
        ["S", "GLS", "H"], ["SGLS", "H"], ["S", "GL", "H"],
    ],
    "QTY": ["QTY"],
    "GLASS": [["GLASS"], ["DESP"]],
}

def standardize_glass_spec(val: str) -> str:
    if pd.isna(val) or not str(val).strip():
        return "NOT SPECIFIED"
    text = str(val).strip()
    if text.lower() == "nan" or not text:
        return "NOT SPECIFIED"
    return re.sub(r"\s+", " ", text).strip()

@dataclass(slots=True)
class HeaderInfo:
    row_index: int
    code_col: Optional[int] = None
    width_col: Optional[int] = None
    height_col: Optional[int] = None
    qty_col: Optional[int] = None
    glass_col: Optional[int] = None
    columns: Dict[str, Optional[int]] = field(default_factory=dict)

@dataclass(slots=True)
class HeaderBlock:
    header: HeaderInfo
    start_row: int
    end_row: int

@dataclass(slots=True)
class GlassRecord:
    WindowCode: str
    Width: int
    Height: int
    Qty: int
    GlassType: str
    SourceFile: str
    SheetName: str

def normalize_header(text: Any) -> str:
    if pd.isna(text):
        return ""
    text = str(text).upper().strip()
    return re.sub(HEADER_REMOVE_PATTERN, "", text)

def normalize_header_row(row: pd.Series) -> List[str]:
    return [normalize_header(val) for val in row.tolist()]

def contains_keywords(text: str, keyword_groups: List[Any]) -> bool:
    if not text:
        return False
    text = normalize_header(text)
    text = re.sub(r"[^A-Z0-9]", "", text)

    for group in keyword_groups:
        if isinstance(group, str):
            group = [group]
        matched = True
        for keyword in group:
            key = re.sub(r"[^A-Z0-9]", "", normalize_header(keyword))
            if key not in text:
                matched = False
                break
        if matched:
            return True
    return False

def detect_column(header_row: List[str], keyword_groups: List[Any]) -> Optional[int]:
    for index, value in enumerate(header_row):
        text = normalize_header(value)
        text = re.sub(r"[^A-Z0-9]", "", text)
        text = (
            text.replace("GLASS", "GL")
            .replace("GLAZING", "GL")
            .replace("SGLS", "GLS")
            .replace("SGL", "GL")
            .replace("GLS.", "GLS")
        )
        for group in keyword_groups:
            if isinstance(group, str):
                group = [group]
            matched = True
            for keyword in group:
                key = re.sub(r"[^A-Z0-9]", "", normalize_header(keyword))
                if key not in text:
                    matched = False
                    break
            if matched:
                return index
    return None

def detect_header_columns(header_row: pd.Series) -> Dict[str, Optional[int]]:
    normalized = normalize_header_row(header_row)
    columns = {
        "code": detect_column(normalized, KEYWORDS["CODE"]),
        "width": detect_column(normalized, KEYWORDS["WIDTH"]),
        "height": detect_column(normalized, KEYWORDS["HEIGHT"]),
        "qty": detect_column(normalized, KEYWORDS["QTY"]),
        "glass": detect_column(normalized, KEYWORDS["GLASS"]),
    }

    if columns["width"] is None:
        for kw in [
            ["S", "GLS", "W"], ["S", "GL", "W"], ["GLS", "W"], ["GL", "W"],
            ["S", "GZ", "W"], ["GZ", "W"], ["FWIDTH"], ["F", "WIDTH"],
            ["SWIDTH"], ["S", "WIDTH"],
        ]:
            col = detect_column(normalized, kw)
            if col is not None:
                columns["width"] = col
                break

    if columns["height"] is None:
        for kw in [
            ["S", "GLS", "H"], ["S", "GL", "H"], ["GLS", "H"], ["GL", "H"],
            ["S", "GZ", "H"], ["GZ", "H"], ["FHEIGHT"], ["F", "HEIGHT"],
            ["SHEIGHT"], ["S", "HEIGHT"],
        ]:
            col = detect_column(normalized, kw)
            if col is not None:
                columns["height"] = col
                break

    return columns

def is_business_header(row: pd.Series) -> bool:
    normalized = normalize_header_row(row)
    has_code = False
    has_qty = False
    has_glass = False

    for value in normalized:
        if contains_keywords(value, KEYWORDS["CODE"]):
            has_code = True
        if contains_keywords(value, KEYWORDS["QTY"]):
            has_qty = True
        if (
            contains_keywords(value, KEYWORDS["GLASS"])
            or contains_keywords(value, KEYWORDS["WIDTH"])
            or contains_keywords(value, KEYWORDS["HEIGHT"])
        ):
            has_glass = True

    return has_code and has_qty and has_glass

def find_header_blocks(dataframe: pd.DataFrame) -> List[HeaderInfo]:
    headers: List[HeaderInfo] = []
    rows = min(len(dataframe), HEADER_SCAN_LIMIT)

    for row_number in range(rows):
        row = dataframe.iloc[row_number]
        if not is_business_header(row):
            continue

        columns = detect_header_columns(row)
        header = HeaderInfo(
            row_index=row_number,
            code_col=columns["code"],
            width_col=columns["width"],
            height_col=columns["height"],
            qty_col=columns["qty"],
            glass_col=columns["glass"],
            columns=columns,
        )
        headers.append(header)

    return headers

def build_header_blocks(dataframe: pd.DataFrame, headers: List[HeaderInfo]) -> List[HeaderBlock]:
    blocks: List[HeaderBlock] = []
    if not headers:
        return blocks

    headers = sorted(headers, key=lambda h: h.row_index)
    for i, header in enumerate(headers):
        start = header.row_index + 1
        end = (
            len(dataframe) - 1
            if i == len(headers) - 1
            else headers[i + 1].row_index - 1
        )
        blocks.append(HeaderBlock(header=header, start_row=start, end_row=end))

    return blocks

def score_business_sheet(df: pd.DataFrame) -> int:
    score = 0
    scan_rows = min(len(df), HEADER_SCAN_LIMIT)

    for r in range(scan_rows):
        row = normalize_header_row(df.iloc[r])
        text = " ".join(row)
        if "CODE" in text:
            score += 5
        if "QTY" in text:
            score += 8
        if "GLASS" in text:
            score += 8
        if re.search(r"\bGL.*W\b", text):
            score += 10
        if re.search(r"\bGL.*H\b", text):
            score += 10

    return score

def find_business_sheets(workbook: Dict[str, pd.DataFrame]) -> List[Tuple[str, pd.DataFrame]]:
    business_sheets = []
    threshold = 35

    for sheet_name, df in workbook.items():
        score = score_business_sheet(df)
        accept_sheet = False

        if score >= threshold:
            accept_sheet = True
        else:
            header_blocks = find_header_blocks(df)
            if header_blocks:
                accept_sheet = True

        if accept_sheet:
            business_sheets.append((sheet_name, df))

    return business_sheets

def safe_numeric(value: Any) -> Optional[int]:
    if pd.isna(value):
        return None
    try:
        val = float(value)
        if val <= 0:
            return None
        return int(math.floor(val + 0.5))
    except Exception:
        return None

def build_window_code(row: pd.Series, header: HeaderInfo) -> Optional[str]:
    if header.code_col is None:
        return None

    val = row.iloc[header.code_col]
    if pd.isna(val):
        return None

    code = str(val).strip()
    if code == "" or code.lower() == "nan":
        return None

    code = re.sub(r"\.0$", "", code)
    code = re.sub(r"\s+", " ", code).strip()

    if header.code_col + 1 < len(row):
        next_val = row.iloc[header.code_col + 1]
        if not pd.isna(next_val):
            next_str = str(next_val).strip()
            next_str = re.sub(r"\.0$", "", next_str)

            if next_str and next_str.lower() != "nan":
                if code.isdigit():
                    code = f"{code} {next_str}"
                elif not re.search(r"\b[A-Z]*\d+\b", code, re.I):
                    if re.match(r"^(W|D|K|CW|NW)\d*$", next_str, re.I):
                        code = f"{code} {next_str}"

    return re.sub(r"\s+", " ", code).strip()

def is_record_start(row: pd.Series, header: HeaderInfo) -> bool:
    try:
        code_parts = []
        for col in range(min(3, len(row))):
            value = row.iloc[col]
            if pd.isna(value):
                continue
            val_str = str(value).strip()
            if val_str == "":
                continue
            code_parts.append(normalize_header(val_str))

        if not code_parts:
            return False
        code = " ".join(code_parts).strip()

        invalid = ("CODE", "FWIDTH", "FHEIGHT", "GLSW", "GLSH", "GLASS", "QTY", "DESCRIPTION")
        if code in invalid or any(inv == code for inv in invalid):
            return False

        upper = code.upper()
        if any(w in upper for w in ["FRAME", "WINDOW", "SECTION", "PROFILE"]):
            return False
        if not re.search(r"[A-Z0-9]", upper):
            return False

        return True
    except Exception:
        return False

def build_record_buffers(dataframe: pd.DataFrame, block: HeaderBlock) -> List[pd.DataFrame]:
    buffers: List[pd.DataFrame] = []
    current_rows = []

    for row_no in range(block.start_row, block.end_row + 1):
        row = dataframe.iloc[row_no]
        if is_record_start(row, block.header):
            if current_rows:
                buffers.append(pd.DataFrame(current_rows))
                current_rows = []
        current_rows.append(row)

    if current_rows:
        buffers.append(pd.DataFrame(current_rows))

    return buffers

def collect_numeric_from_buffer(buffer: pd.DataFrame, column_index: Optional[int]) -> Optional[int]:
    if column_index is None:
        return None
    for _, row in buffer.iterrows():
        if column_index >= len(row):
            continue
        val = safe_numeric(row.iloc[column_index])
        if val is not None:
            return val
    return None

def collect_glass(buffer: pd.DataFrame, header: HeaderInfo) -> Optional[str]:
    parts = []
    seen = set()

    if header.glass_col is not None:
        for _, row in buffer.iterrows():
            if header.glass_col >= len(row):
                continue
            val = row.iloc[header.glass_col]
            if pd.isna(val):
                continue
            val_str = re.sub(r"\s+", " ", str(val).strip()).strip()
            if val_str == "" or val_str.upper() in seen:
                continue
            seen.add(val_str.upper())
            parts.append(val_str)

    if parts:
        return " ".join(parts)

    for _, row in buffer.iterrows():
        for cell in row.values:
            if pd.isna(cell):
                continue
            cell_str = str(cell).strip()
            upper_cell = cell_str.upper()
            if any(k in upper_cell for k in ["LAMINATED", "TOUGHENED", "DGU", "SGU", "CLEAR", "FROSTED", "MM"]):
                if not any(k in upper_cell for k in ["CODE", "HANDLE", "LOCK", "COLOR", "FRAME", "WIDTH", "HEIGHT"]):
                    return cell_str

    return None

def parse_header_block(dataframe: pd.DataFrame, block: HeaderBlock, source_file: str, sheet_name: str) -> List[GlassRecord]:
    records: List[GlassRecord] = []
    buffers = build_record_buffers(dataframe, block)

    for buffer in buffers:
        first_row = buffer.iloc[0]
        window = build_window_code(first_row, block.header)

        width = collect_numeric_from_buffer(buffer, block.header.width_col)
        height = collect_numeric_from_buffer(buffer, block.header.height_col)
        qty = collect_numeric_from_buffer(buffer, block.header.qty_col)
        glass_raw = collect_glass(buffer, block.header)

        if not window or width is None or height is None:
            continue
        if qty is None:
            qty = 1

        if glass_raw and "FROSTED" in str(glass_raw).upper():
            continue

        glass = standardize_glass_spec(glass_raw)

        records.append(
            GlassRecord(
                WindowCode=window,
                Width=width,
                Height=height,
                Qty=qty,
                GlassType=glass,
                SourceFile=source_file,
                SheetName=sheet_name,
            )
        )

    return records

def parse_business_sheet(dataframe: pd.DataFrame, source_file: str, sheet_name: str) -> List[GlassRecord]:
    headers = find_header_blocks(dataframe)
    blocks = build_header_blocks(dataframe, headers)
    all_records: List[GlassRecord] = []

    for block in blocks:
        all_records.extend(parse_header_block(dataframe, block, source_file, sheet_name))

    return all_records

def load_excel_with_calculated_values(file) -> Dict[str, pd.DataFrame]:
    file_bytes = io.BytesIO(file.read())
    file.seek(0)
    wb = openpyxl.load_workbook(file_bytes, data_only=True)
    workbook_dict = {}

    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        data = sheet.values
        cols = next(data, None)
        if cols is None:
            continue

        data_rows = list(data)
        if cols:
            data_rows.insert(0, cols)

        df = pd.DataFrame(data_rows)
        workbook_dict[sheet_name] = df

    return workbook_dict

def process_uploaded_files(uploaded_files) -> pd.DataFrame:
    all_records = []

    for file in uploaded_files:
        try:
            workbook_dict = load_excel_with_calculated_values(file)
            filtered_workbook = {}
            for s_name, df in workbook_dict.items():
                s_name_upper = s_name.upper().strip()
                if any(k in s_name_upper for k in ["CLIENT", "DETAIL", "QUOTE"]):
                    continue
                filtered_workbook[s_name] = df

            business_sheets = find_business_sheets(filtered_workbook)
            for sheet_name, df in business_sheets:
                records = parse_business_sheet(df, file.name, sheet_name)
                all_records.extend(records)

        except Exception as e:
            st.error(f"Error processing file {file.name}: {e}")

    return pd.DataFrame([asdict(r) for r in all_records]).reset_index(drop=True) if all_records else pd.DataFrame()

# ============================================================
# 3. Main Dashboard UI Operations
# ============================================================
uploaded_files = st.file_uploader(
    "Upload BOQ Excel Files",
    type=["xlsx", "xlsm", "xls"],
    accept_multiple_files=True,
    key=f"boq_uploader_{st.session_state['uploader_key']}"
)

col_btn1, col_btn2, _ = st.columns([1.5, 1.5, 7])

with col_btn1:
    if st.button("Merge & Process Files", type="primary"):
        if uploaded_files:
            with st.spinner("Processing BOQ Files..."):
                df_merged = process_uploaded_files(uploaded_files)
                if not df_merged.empty:
                    st.session_state["merged_df"] = df_merged
                    st.toast(f"Successfully processed {len(df_merged)} items!", icon="✅")
                else:
                    st.error("No valid BOQ records found.")
        else:
            st.warning("Please upload files first.")

with col_btn2:
    if st.button("Reset Data"):
        for key in ["merged_df", "req_bytes", "req_generated"]:
            if key in st.session_state:
                del st.session_state[key]
        st.session_state["uploader_key"] += 1
        st.rerun()

# Display Merged Preview Dashboard
if "merged_df" in st.session_state:
    df_merged = st.session_state["merged_df"]

    st.markdown("---")
    st.markdown("### 📊 Extracted BOQ Summary")
    
    # Dashboard Metrics
    m1, m2, m3 = st.columns(3)
    m1.metric("Total Extracted Items", len(df_merged))
    m2.metric("Total Quantity", df_merged["Qty"].sum() if "Qty" in df_merged else 0)
    m3.metric("Glass Types", df_merged["GlassType"].nunique() if "GlassType" in df_merged else 0)

    st.dataframe(df_merged, use_container_width=True)

    st.markdown("---")
    st.markdown("### ⚡ Generate Exact Excel Output Sheet")

    if st.button("⚡ GENERATE REQUIREMENT SHEET", type="primary"):
        with st.spinner("Creating Native Clean Excel File..."):
            
            # Exact Excel Generation using OpenPyXL
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws.views.sheetView[0].showGridLines = True

            # EXACT STYLES FROM SCREENSHOT
            title_font = Font(name="Calibri", size=16, bold=True, color="000000")
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            data_font = Font(name="Calibri", size=11, bold=False, color="000000")
            total_font = Font(name="Calibri", size=11, bold=True, color="000000")

            title_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
            header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")

            thin_border = Border(
                left=Side(style="thin", color="000000"),
                right=Side(style="thin", color="000000"),
                top=Side(style="thin", color="000000"),
                bottom=Side(style="thin", color="000000"),
            )

            # Title
            today_str = datetime.now().strftime("%d %b %Y").upper()
            title_text = f"1 WIN-SQUARE {today_str}"
            st.session_state["generated_title_name"] = title_text

            # ROW 1: TITLE ROW
            ws.merge_cells("A1:H1")
            ws.row_dimensions[1].height = 28
            ws["A1"].value = title_text
            ws["A1"].font = title_font
            ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

            for col_i in range(1, 9):
                cell = ws.cell(row=1, column=col_i)
                cell.fill = title_fill
                cell.border = thin_border

            # ROW 2: HEADER ROW
            headers = ["Sr.No", "WINDOW CODE", "WIDTH", "HEIGHT", "SQFT", "QTY", "TTL SQFT", "REMARKS"]
            ws.row_dimensions[2].height = 22
            for col_i, h_text in enumerate(headers, 1):
                cell = ws.cell(row=2, column=col_i, value=h_text)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border

            # DATA ROWS
            for idx, row in df_merged.iterrows():
                r_idx = idx + 3
                sqft_formula = f"=ROUND((C{r_idx}*D{r_idx})/92903.04, 2)"
                ttl_sqft_formula = f"=E{r_idx}*F{r_idx}"

                row_data = [
                    idx + 1,
                    row["WindowCode"],
                    row["Width"],
                    row["Height"],
                    sqft_formula,
                    row["Qty"],
                    ttl_sqft_formula,
                    row["GlassType"],
                ]

                ws.row_dimensions[r_idx].height = 20

                for col_i, val in enumerate(row_data, 1):
                    cell = ws.cell(row=r_idx, column=col_i, value=val)
                    cell.font = data_font
                    cell.border = thin_border

                    # Exact Cell Formats & Alignment
                    if col_i in [1, 2, 8]:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    elif col_i in [3, 4, 6]:
                        cell.number_format = "0"
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    elif col_i in [5, 7]:
                        cell.number_format = "0.00"
                        cell.alignment = Alignment(horizontal="center", vertical="center")

            # TOTAL ROW
            tot_row = len(df_merged) + 3
            ws.row_dimensions[tot_row].height = 22

            for c in range(1, 9):
                ws.cell(row=tot_row, column=c).border = thin_border

            cell_tot = ws.cell(row=tot_row, column=5, value="TOTAL")
            cell_tot.font = total_font
            cell_tot.alignment = Alignment(horizontal="right", vertical="center")

            qty_sum = ws.cell(row=tot_row, column=6, value=f"=SUM(F3:F{tot_row-1})")
            qty_sum.font = total_font
            qty_sum.number_format = "0"
            qty_sum.alignment = Alignment(horizontal="center", vertical="center")

            ttl_sqft_sum = ws.cell(row=tot_row, column=7, value=f"=SUM(G3:G{tot_row-1})")
            ttl_sqft_sum.font = total_font
            ttl_sqft_sum.number_format = "0.00"
            ttl_sqft_sum.alignment = Alignment(horizontal="center", vertical="center")

            # Column Widths
            col_widths = {'A': 8, 'B': 16, 'C': 10, 'D': 10, 'E': 10, 'F': 8, 'G': 12, 'H': 32}
            for col_letter, width in col_widths.items():
                ws.column_dimensions[col_letter].width = width

            output = io.BytesIO()
            wb.save(output)
            st.session_state["req_bytes"] = output.getvalue()
            st.session_state["req_generated"] = True

    if st.session_state.get("req_generated"):
        file_download_name = f"{st.session_state.get('generated_title_name', '1 WIN-SQUARE')}.xlsx"
        st.download_button(
            label=f"📥 Download Exact Excel ({file_download_name})",
            data=st.session_state["req_bytes"],
            file_name=file_download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
