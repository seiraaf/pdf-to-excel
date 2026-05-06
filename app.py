import streamlit as st
import pdfplumber
import pandas as pd
import re

# 🔥 CONFIG PAGE
st.set_page_config(page_title="PDF to Excel", page_icon="📊", layout="wide")

# 🎨 STYLE
st.markdown("""
<style>
.main {
    background-color: #f5f7fa;
}
.stButton>button {
    background-color: #6c63ff;
    color: white;
    border-radius: 10px;
    height: 45px;
    width: 100%;
}
.stDownloadButton>button {
    background-color: #00b894;
    color: white;
    border-radius: 10px;
    height: 45px;
    width: 100%;
}
</style>
""", unsafe_allow_html=True)

# 🧠 HEADER
st.title("📊 PDF to Excel Converter")
st.caption("Upload your PDF and convert it instantly ✨")

# 📦 UPLOAD SECTION
st.subheader("📤 Upload File")
uploaded_files = st.file_uploader("Upload PDF", accept_multiple_files=True)

unit = st.selectbox(
    "Pilih Unit",
    ["Nirwana", "Lovina", "Lembongan", "The Club"]
)

all_data = []

pattern1 = re.compile(r"^\d+\s+(.*?)\s+([\d,\.]+)\s+([A-Z]+)\s+([\d,\.]+)\s+([\d,\.]+)")
pattern2 = re.compile(r"^(.*?)\s+([\d,]+\.\d+)\s+([A-Z]+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+(.+)$")

# 🚀 CONVERT BUTTON
if uploaded_files:
    st.success(f"{len(uploaded_files)} file uploaded ✅")

    if st.button("🚀 Convert Now"):
        with st.spinner("Processing... ⏳"):
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
                                    "Item": name,
                                    "Qty": qty,
                                    "Unit": unit,
                                    "Price": price,
                                    "Amount": amount
                                })

        df = pd.DataFrame(all_data)

        if not df.empty:
            st.success("✅ Conversion success!")

            # 📊 PREVIEW DATA
            st.subheader("📊 Preview Data")
            st.dataframe(df, use_container_width=True)

            # 💾 DOWNLOAD
            file = "hasil.xlsx"
            df.to_excel(file, index=False)

            with open(file, "rb") as f:
                st.download_button("⬇️ Download Excel", f, file_name="hasil.xlsx")

        else:
            st.error("Data kosong 😭")
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
