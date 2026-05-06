import streamlit as st
import pdfplumber
import pandas as pd
import re

# 🔥 CONFIG
st.set_page_config(page_title="PDF to Excel", page_icon="📊", layout="wide")

# 🧠 HEADER
st.title("📊 PDF to Excel Converter")
st.caption("Upload per unit → convert jadi 1 Excel ✨")

# 🔍 PATTERN (HARUS DI ATAS!)
pattern1 = re.compile(r"^\d+\s+(.*?)\s+([\d,\.]+)\s+([A-Z]+)\s+([\d,\.]+)\s+([\d,\.]+)")
pattern2 = re.compile(r"^(.*?)\s+([\d,\.]+)\s+([A-Z]+)\s+([\d,\.]+)\s+([\d,\.]+)\s+(.*)$")

# 🧠 SESSION
if "all_data" not in st.session_state:
    st.session_state.all_data = []

# 🏢 PILIH UNIT
unit = st.selectbox(
    "Pilih Unit Upload",
    ["Nirwana", "Lovina", "Lembongan", "The Club", "Kanaka"]
)

# 📤 UPLOAD
uploaded_files = st.file_uploader(
    f"Upload PDF untuk {unit}",
    accept_multiple_files=True,
    key=unit
)

# ➕ TAMBAH DATA
if uploaded_files and st.button(f"Tambah Data {unit}"):
    for uploaded_file in uploaded_files:
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()

                if not text:
                    continue

                lines = text.split("\n")

                for line in lines:
                    line = line.strip()

                    if "Subtotal" in line or "Supplier" in line:
                        continue

                    remarks = "-"

                    match1 = pattern1.match(line)
                    match2 = pattern2.match(line)

                    if match1:
                        name, qty, unit_item, price, amount = match1.groups()
                        supplier = "Unknown"

                    elif match2:
                        name, qty, unit_item, price, amount, tail = match2.groups()

                        parts = tail.split()

                        if len(parts) >= 2:
                            supplier = " ".join(parts[-2:])
                            remarks = " ".join(parts[:-2])
                        else:
                            supplier = tail

                    else:
                        continue

                    st.session_state.all_data.append({
                        "Unit": unit,
                        "Supplier": supplier,
                        "Item": name,
                        "Qty": qty,
                        "Unit Item": unit_item,
                        "Price": price,
                        "Amount": amount,
                        "Remarks": remarks
                    })

    st.success(f"✅ Data {unit} berhasil ditambahkan!")

# 🚀 CONVERT SEMUA
if st.session_state.all_data:
    if st.button("🚀 Convert Semua Unit"):
        df = pd.DataFrame(st.session_state.all_data)

        st.subheader("📊 Preview Data")
        st.dataframe(df, use_container_width=True)

        file = "hasil_final.xlsx"

        with pd.ExcelWriter(file) as writer:
            for u in df["Unit"].unique():
                df[df["Unit"] == u].to_excel(writer, sheet_name=u, index=False)

        with open(file, "rb") as f:
            st.download_button("⬇️ Download Excel", f, file_name=file)
