import streamlit as st
import pdfplumber
import pandas as pd
import re

st.title("PDF to Excel Converter ✨")

uploaded_files = st.file_uploader("Upload PDF", accept_multiple_files=True)

all_data = []

pattern1 = re.compile(r"^\d+\s+(.*?)\s+([\d,\.]+)\s+([A-Z]+)\s+([\d,\.]+)\s+([\d,\.]+)")
pattern2 = re.compile(r"^(.*?)\s+([\d,]+\.\d+)\s+([A-Z]+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+(.+)$")

if uploaded_files:
    if st.button("Convert"):
        for uploaded_file in uploaded_files:
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()

                    if text:
                        lines = text.split("\n")

                        for line in lines:
                            line = line.strip()

                            if "Subtotal" in line or "Supplier" in line:
                                continue

                            match1 = pattern1.match(line)
                            match2 = pattern2.match(line)

                            if match1:
                                name, qty, unit, price, amount = match1.groups()
                                supplier = "Unknown"

                            elif match2:
                                name, qty, unit, price, amount, supplier = match2.groups()

                            else:
                                continue

                            all_data.append({
                                "Supplier": supplier,
                                "Inv Name": name,
                                "Qty": qty,
                                "Unit": unit,
                                "Price": price,
                                "Amount": amount
                            })

        df = pd.DataFrame(all_data)

        if not df.empty:
            file = "hasil.xlsx"
            df.to_excel(file, index=False)

            with open(file, "rb") as f:
                st.download_button("Download Excel", f, file_name="hasil.xlsx")

        else:
            st.error("Data kosong 😭")