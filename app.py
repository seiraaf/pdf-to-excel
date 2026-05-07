from io import BytesIO
import re

import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


UNITS = ["Nirwana", "Lovina", "Lembongan", "The Club", "Kanaka"]
COLUMNS = ["Unit", "Supplier", "Item", "Qty", "Unit Item", "Price", "Amount", "Remarks"]
EXPORT_HEADERS = ["Supplier", "No PO", "Inv Name", "Q", "Unit", "Description", "Qty Datang", "Remark"]
UNIT_TITLES = {
    "Nirwana": "NIRWANA BEACH RESORT",
    "Lovina": "LOVINA BEACH RESORT",
    "Lembongan": "LEMBONGAN",
    "The Club": "THE CLUB",
    "Kanaka": "KANAKA",
}
MONTHS_ID = {
    "jan": "JANUARI",
    "january": "JANUARI",
    "feb": "FEBRUARI",
    "february": "FEBRUARI",
    "mar": "MARET",
    "march": "MARET",
    "apr": "APRIL",
    "april": "APRIL",
    "may": "MEI",
    "jun": "JUNI",
    "june": "JUNI",
    "jul": "JULI",
    "july": "JULI",
    "aug": "AGUSTUS",
    "august": "AGUSTUS",
    "sep": "SEPTEMBER",
    "sept": "SEPTEMBER",
    "september": "SEPTEMBER",
    "oct": "OKTOBER",
    "october": "OKTOBER",
    "nov": "NOVEMBER",
    "november": "NOVEMBER",
    "dec": "DESEMBER",
    "december": "DESEMBER",
}

NUMBER_PATTERN = r"\d{1,3}(?:,\d{3})*(?:\.\d+)?|\d+(?:\.\d+)?"
PATTERN_WITH_NO = re.compile(
    rf"^\s*\d+\s+(?P<item>.*?)\s+(?P<qty>{NUMBER_PATTERN})\s+(?P<unit>[A-Za-z]+)\s+"
    rf"(?P<price>{NUMBER_PATTERN})\s+(?P<amount>{NUMBER_PATTERN})(?:\s+(?P<tail>.*))?\s*$",
    re.IGNORECASE,
)
PATTERN_WITHOUT_NO = re.compile(
    rf"^\s*(?P<item>.*?)\s+(?P<qty>{NUMBER_PATTERN})\s+(?P<unit>[A-Za-z]+)\s+"
    rf"(?P<price>{NUMBER_PATTERN})\s+(?P<amount>{NUMBER_PATTERN})(?:\s+(?P<tail>.*))?\s*$",
    re.IGNORECASE,
)

SKIP_KEYWORDS = (
    "subtotal",
    "grand total",
    "purchase order",
    "prepared by",
    "approved by",
    "total :",
)
SUPPLIER_HINTS = (
    "supplier",
    "seafood",
    "ayam",
    "bali",
    "ud ",
    "cv ",
    "pt ",
    "eme ",
)
TABLE_SETTINGS = {
    "vertical_strategy": "text",
    "horizontal_strategy": "text",
    "snap_tolerance": 4,
    "join_tolerance": 4,
    "intersection_tolerance": 6,
    "text_tolerance": 3,
}


st.set_page_config(page_title="PDF to Excel", page_icon="📊", layout="wide")


if "all_data" not in st.session_state:
    st.session_state.all_data = []

if "excel_file" not in st.session_state:
    st.session_state.excel_file = None


def clean_text(value, default="-"):
    value = re.sub(r"\s+", " ", str(value or "")).strip()
    return value if value else default


def normalize_number(value):
    return clean_text(value).replace(" ", "")


def to_excel_number(value):
    value = clean_text(value, "").replace(",", "")
    try:
        number = float(value)
    except ValueError:
        return clean_text(value)
    return int(number) if number.is_integer() else number


def looks_like_number(value):
    return bool(re.fullmatch(NUMBER_PATTERN, clean_text(value, "")))


def looks_like_unit(value):
    value = clean_text(value, "")
    return bool(re.fullmatch(r"[A-Za-z]{1,8}", value))


def should_skip_line(line):
    lowered = line.lower()
    return not line or any(keyword in lowered for keyword in SKIP_KEYWORDS)


def format_po_date(raw_date):
    raw_date = clean_text(raw_date, "")
    match = re.search(r"(\d{1,2})[-/\s]([A-Za-z]+)[-/\s](\d{2,4})", raw_date)
    if not match:
        return ""

    day, month, year = match.groups()
    if len(year) == 2:
        year = "20" + year

    month_name = MONTHS_ID.get(month.lower(), month.upper())
    return f"PO {int(day)} {month_name} {year}"


def extract_po_date(text):
    for line in (text or "").splitlines():
        match = re.search(r"\bDate\b\s*:?[\s-]*(.+)$", line, re.IGNORECASE)
        if match:
            po_date = format_po_date(match.group(1))
            if po_date:
                return po_date
    return ""


def extract_page_remark(text):
    for line in (text or "").splitlines():
        match = re.search(r"\bRemark\b\s*:?[\s-]*(.+)$", line, re.IGNORECASE)
        if match:
            remark = clean_text(match.group(1), "")
            if remark and not remark.lower().startswith(("price", "amount", "supplier")):
                return remark
    return "-"


def split_tail(tail, default_remarks="-"):
    """Pisahkan supplier dan remarks dari teks setelah kolom amount."""
    tail = clean_text(tail, "")
    default_remarks = clean_text(default_remarks)

    if not tail:
        return default_remarks, "Unknown"

    separator_match = re.search(r"\s{2,}|\s[-–—]\s|\s\|\s", tail)
    if separator_match:
        remarks = clean_text(tail[: separator_match.start()], "")
        supplier = clean_text(tail[separator_match.end() :], "")
        if remarks and supplier:
            return remarks, supplier

    lowered_tail = f" {tail.lower()} "
    if default_remarks != "-" or any(hint in lowered_tail for hint in SUPPLIER_HINTS):
        return default_remarks, tail

    parts = tail.split()
    if len(parts) == 1:
        return default_remarks, parts[0]

    supplier_word_count = 2 if len(parts) >= 3 else 1
    remarks = clean_text(" ".join(parts[:-supplier_word_count]))
    supplier = clean_text(" ".join(parts[-supplier_word_count:]), "Unknown")
    return remarks, supplier


def parse_line(line, unit, default_remarks="-", po_date=""):
    line = clean_text(line, "")
    if should_skip_line(line):
        return None

    match = PATTERN_WITH_NO.match(line) or PATTERN_WITHOUT_NO.match(line)
    if not match:
        return None

    data = match.groupdict()
    remarks, supplier = split_tail(data.get("tail", ""), default_remarks)

    return {
        "Unit": unit,
        "Supplier": supplier,
        "Item": clean_text(data["item"]),
        "Qty": normalize_number(data["qty"]),
        "Unit Item": clean_text(data["unit"]).upper(),
        "Price": normalize_number(data["price"]),
        "Amount": normalize_number(data["amount"]),
        "Remarks": remarks,
        "PO Date": po_date,
    }


def parse_table_row(cells, unit, default_remarks="-", po_date=""):
    cells = [clean_text(cell, "") for cell in cells]
    cells = [cell for cell in cells if cell]
    line = " ".join(cells)

    if should_skip_line(line) or "item name" in line.lower() or len(cells) < 4:
        return None

    number_indexes = [index for index, cell in enumerate(cells) if looks_like_number(cell)]
    if len(number_indexes) < 3:
        return None

    qty_index = number_indexes[0]
    amount_index = number_indexes[-1]
    price_index = number_indexes[-2]

    unit_index = None
    for index in range(qty_index + 1, price_index):
        if looks_like_unit(cells[index]):
            unit_index = index
            break

    if unit_index is None:
        return None

    item = clean_text(" ".join(cells[:qty_index]), "")
    if not item:
        return None

    tail = clean_text(" ".join(cells[amount_index + 1 :]), "")
    remarks, supplier = split_tail(tail, default_remarks)

    return {
        "Unit": unit,
        "Supplier": supplier,
        "Item": item,
        "Qty": normalize_number(cells[qty_index]),
        "Unit Item": clean_text(cells[unit_index]).upper(),
        "Price": normalize_number(cells[price_index]),
        "Amount": normalize_number(cells[amount_index]),
        "Remarks": remarks,
        "PO Date": po_date,
    }


def row_key(row):
    return tuple(row[column] for column in ["Unit", "Item", "Qty", "Unit Item", "Price", "Amount"])


def parse_pdf(uploaded_file, unit):
    rows = []
    seen = set()

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=2, y_tolerance=3) or ""
            default_remarks = extract_page_remark(text)
            po_date = extract_po_date(text)

            for table in page.extract_tables(TABLE_SETTINGS) or []:
                for cells in table:
                    row = parse_table_row(cells, unit, default_remarks, po_date)
                    if row and row_key(row) not in seen:
                        rows.append(row)
                        seen.add(row_key(row))

            for line in text.splitlines():
                row = parse_line(line, unit, default_remarks, po_date)
                if row and row_key(row) not in seen:
                    rows.append(row)
                    seen.add(row_key(row))

    return rows


def export_row(row):
    return [
        "" if row.get("Supplier") == "Unknown" else row.get("Supplier", ""),
        "",
        row.get("Item", ""),
        to_excel_number(row.get("Qty", "")),
        row.get("Unit Item", ""),
        "",
        "",
        "" if row.get("Remarks") == "-" else row.get("Remarks", ""),
    ]


def style_sheet(ws, unit_name, po_date, last_row):
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="F2F2F2")

    ws.merge_cells("A1:H1")
    ws.merge_cells("A2:H2")
    ws["A1"] = UNIT_TITLES.get(unit_name, unit_name.upper())
    ws["A2"] = po_date or "PO"
    ws["A1"].font = Font(bold=True, size=20)
    ws["A2"].font = Font(bold=True, size=10)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center")

    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 16
    ws.row_dimensions[3].height = 8
    ws.row_dimensions[4].height = 18

    widths = [18, 18, 44, 10, 8, 48, 16, 16]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width

    for cell in ws[4]:
        cell.font = Font(bold=True, size=9)
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in ws.iter_rows(min_row=5, max_row=max(last_row, 5), min_col=1, max_col=8):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.font = Font(size=10)
        row[2].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        row[3].alignment = Alignment(horizontal="right", vertical="center")
        row[4].alignment = Alignment(horizontal="center", vertical="center")
        row[3].number_format = "#,##0.00"

    ws.auto_filter.ref = f"A4:H{max(last_row, 5)}"
    ws.freeze_panes = "A5"
    ws.sheet_view.showGridLines = True


def build_excel(data):
    output = BytesIO()
    df = pd.DataFrame(data)
    workbook = Workbook()
    workbook.remove(workbook.active)

    for unit_name in UNITS:
        unit_df = df[df["Unit"] == unit_name] if not df.empty else pd.DataFrame()
        if unit_df.empty:
            continue

        ws = workbook.create_sheet(unit_name[:31])
        po_dates = [clean_text(value, "") for value in unit_df.get("PO Date", []) if clean_text(value, "")]
        po_date = po_dates[0] if po_dates else ""

        ws.append([""] * len(EXPORT_HEADERS))
        ws.append([""] * len(EXPORT_HEADERS))
        ws.append([""] * len(EXPORT_HEADERS))
        ws.append(EXPORT_HEADERS)

        for row in unit_df.to_dict("records"):
            ws.append(export_row(row))

        style_sheet(ws, unit_name, po_date, ws.max_row)

    if not workbook.sheetnames:
        ws = workbook.create_sheet("Sheet1")
        ws.append(EXPORT_HEADERS)

    workbook.save(output)
    output.seek(0)
    return output


st.title("📊 PDF to Excel Converter")
st.caption("Upload per unit → Tambah Data → Convert Semua Unit → download 1 Excel multi sheet.")

with st.sidebar:
    st.header("Status Data")
    if st.session_state.all_data:
        status_df = pd.DataFrame(st.session_state.all_data)
        counts = status_df.groupby("Unit").size().reindex(UNITS, fill_value=0)
    else:
        counts = pd.Series(0, index=UNITS)

    for unit_name, total in counts.items():
        st.write(f"**{unit_name}**: {total} baris")

    if st.session_state.all_data and st.button("Reset Data Sementara"):
        st.session_state.all_data = []
        st.session_state.excel_file = None
        st.rerun()

unit = st.selectbox("Pilih Unit Upload", UNITS)

uploaded_files = st.file_uploader(
    f"Upload PDF untuk {unit}",
    type=["pdf"],
    accept_multiple_files=True,
    key=unit,
)

col_add, col_hint = st.columns([1, 2])
with col_add:
    add_clicked = st.button(
        f"Tambah Data {unit}",
        disabled=not uploaded_files,
        use_container_width=True,
    )

with col_hint:
    if uploaded_files:
        st.info(f"{len(uploaded_files)} file PDF siap ditambahkan untuk {unit}.")
    else:
        st.info("Pilih PDF purchase order untuk unit yang sedang dipilih.")

if add_clicked:
    added_rows = 0
    empty_files = []

    with st.spinner(f"Membaca PDF untuk {unit}..."):
        for uploaded_file in uploaded_files:
            rows = parse_pdf(uploaded_file, unit)
            if rows:
                st.session_state.all_data.extend(rows)
                added_rows += len(rows)
            else:
                empty_files.append(uploaded_file.name)

    st.session_state.excel_file = None

    if added_rows:
        st.success(f"Data {unit} berhasil ditambahkan: {added_rows} baris.")
    if empty_files:
        st.warning(
            "Tidak ada baris item yang cocok dari file: "
            + ", ".join(empty_files)
            + ". Cek format PDF atau teks hasil extract."
        )

if st.session_state.all_data:
    df_preview = pd.DataFrame(st.session_state.all_data)
    preview_columns = [column for column in COLUMNS if column in df_preview.columns]

    st.divider()
    st.subheader("Preview Data Sementara")
    st.dataframe(df_preview[preview_columns], use_container_width=True, hide_index=True)

    if st.button("🚀 Convert Semua Unit", use_container_width=True):
        st.session_state.excel_file = build_excel(st.session_state.all_data)
        st.success("Excel berhasil dibuat dengan format report. Silakan download file di bawah.")

    if st.session_state.excel_file:
        st.download_button(
            "⬇️ Download Excel",
            data=st.session_state.excel_file,
            file_name="hasil_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
else:
    st.divider()
    st.info("Belum ada data. Upload PDF untuk salah satu unit, lalu klik Tambah Data.")
