import streamlit as st
import pandas as pd
import pdfplumber
import re
import zipfile
import io
import gc
from openpyxl import Workbook

CSV_COLUMNS = [
    'Bill Date', 'Bill Number', 'Project Name', 'SubTotal',
    'Tax Amount', 'Total', 'GST Identification Number (GSTIN)', 'Currency Code'
]

def clean_amt(val):
    if not val:
        return 0.0
    return float(re.sub(r'[^\d.]', '', val))

def extract_data_from_pdf(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

        project_name = (re.search(r"Tax invoice for\s+(.*)", text) or re.search(r"Campaigns\s+-\s+(.*)", text))
        inv_match = re.search(r"Invoice (?:no\.|ID)\s*([A-Z0-9-]+)", text)
        sub_raw = re.search(r"Subtotal:\s+([\d,.]+)", text)
        igst_raw = re.search(r"IGST\s\(18%\):\s+([\d,.]+)", text)
        date_match = re.search(r"(\d{1,2}\s[A-Za-z]{3}\s\d{4})", text)
        gstin_match = re.search(r"GSTIN:\s+([A-Z0-9]{15})", text)

        sub_val = clean_amt(sub_raw.group(1)) if sub_raw else 0.0
        igst_val = clean_amt(igst_raw.group(1)) if igst_raw else 0.0

        row = {col: "" for col in CSV_COLUMNS}
        row.update({
            'Bill Date': date_match.group(1) if date_match else "N/A",
            'Bill Number': inv_match.group(1) if inv_match else "N/A",
            'Project Name': project_name.group(1).strip() if project_name else "N/A",
            'SubTotal': sub_val,
            'Tax Amount': igst_val,
            'Total': sub_val + igst_val,
            'GST Identification Number (GSTIN)': gstin_match.group(1) if gstin_match else "NA",
            'Currency Code': 'INR'
        })
        return row
    except Exception as e:
        st.warning(f"Failed to parse a PDF: {e}")
        return None

def generate_excel(df):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoices"
    ws.append(list(df.columns))
    for _, row in df.iterrows():
        ws.append(list(row))
    wb.save(output)
    output.seek(0)
    return output

st.title("📂 High-Capacity Invoice Processor")

uploaded_zips = st.file_uploader("Upload ZIP folders (Max 500MB)", type="zip", accept_multiple_files=True)

if uploaded_zips:
    if st.button("▶ Process Invoices"):
        all_rows = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        for z_idx, zip_file in enumerate(uploaded_zips):
            try:
                with zipfile.ZipFile(zip_file) as z:
                    file_list = [f for f in z.namelist() if f.lower().endswith('.pdf')]
                    total_files = len(file_list)

                    if total_files == 0:
                        st.warning(f"No PDFs found in {zip_file.name}")
                        continue

                    for i, pdf_path in enumerate(file_list):
                        status_text.text(f"[ZIP {z_idx+1}] Processing: {pdf_path} ({i+1}/{total_files})")
                        try:
                            with z.open(pdf_path) as f:
                                data = extract_data_from_pdf(io.BytesIO(f.read()))
                                if data:
                                    all_rows.append(data)
                        except Exception as e:
                            st.warning(f"Skipped {pdf_path}: {e}")

                        progress_bar.progress((i + 1) / total_files)

                        if i % 10 == 0:
                            gc.collect()

            except zipfile.BadZipFile:
                st.error(f"{zip_file.name} is not a valid ZIP file.")

        if all_rows:
            df = pd.DataFrame(all_rows, columns=CSV_COLUMNS)
            excel_data = generate_excel(df)

            st.success(f"✅ Successfully processed {len(all_rows)} invoices!")
            st.dataframe(df)

            st.download_button(
                label="📥 Download Excel",
                data=excel_data,
                file_name="invoices_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("No invoice data could be extracted. Check your PDF format.")
