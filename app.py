import streamlit as st
import pandas as pd
import pdfplumber
import re
import zipfile
import io

def clean_amt(value):
    """Removes currency symbols and commas to return a clean float."""
    if not value or value == "N/A":
        return 0.0
    # Strip everything except digits and decimal point
    clean_val = re.sub(r'[^\d.]', '', value)
    try:
        return float(clean_val)
    except ValueError:
        return 0.0

def extract_data_from_pdf(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
    
    # 1. Project Name
    project_match = re.search(r"Tax invoice for\s+(.*)", text)
    project_name = project_match.group(1).strip() if project_match else "N/A"

    # 2. Invoice Number
    inv_match = re.search(r"Invoice (?:no\.|ID)\s*([A-Z0-9-]+)", text)
    invoice_no = inv_match.group(1).strip() if inv_match else "N/A"

    # 3. Financials (Extract and convert to float for math)
    sub_raw = re.search(r"Subtotal:\s+([\d,.]+)", text)
    igst_raw = re.search(r"IGST\s\(18%\):\s+([\d,.]+)", text)
    
    sub_val = clean_amt(sub_raw.group(1)) if sub_raw else 0.0
    igst_val = clean_amt(igst_raw.group(1)) if igst_raw else 0.0
    total_val = sub_val + igst_val 

    # 4. Date
    date_match = re.search(r"(\d{1,2}\s[A-Za-z]{3}\s\d{4})", text)

    # 5. GSTIN
    gstin_match = re.search(r"GSTIN:\s+([A-Z0-9]{15})", text)

    return {
        "Invoice Number": invoice_no,
        "Project Name": project_name,
        "Subtotal": sub_val,
        "IGST": igst_val,
        "Total": total_val,
        "Document Date": date_match.group(1) if date_match else "N/A",
        "GSTIN": gstin_match.group(1) if gstin_match else "NA"
    }

st.set_page_config(page_title="Invoice Processor", layout="wide")
st.title("📂 Multi-Zip Invoice Data Extractor")

uploaded_zips = st.file_uploader("Upload multiple ZIP folders", type="zip", accept_multiple_files=True)

if uploaded_zips:
    all_data = []
    
    for zip_file in uploaded_zips:
        with zipfile.ZipFile(zip_file) as z:
            for file_info in z.infolist():
                if file_info.filename.lower().endswith('.pdf'):
                    with z.open(file_info) as f:
                        pdf_stream = io.BytesIO(f.read())
                        all_data.append(extract_data_from_pdf(pdf_stream))

    if all_data:
        df = pd.DataFrame(all_data)
        df.insert(0, 'Sr.no.', range(1, len(df) + 1))
        
        # UI Preview with Formatting
        st.subheader("Extracted Data Preview")
        st.dataframe(df.style.format({
            "Subtotal": "{:,.2f} INR", 
            "IGST": "{:,.2f} INR", 
            "Total": "{:,.2f} INR"
        }))

        # Excel Export with Currency Formatting
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='InvoiceData')
            
            workbook  = writer.book
            worksheet = writer.sheets['InvoiceData']
            
            # --- THIS IS THE KEY PART ---
            # This format adds the comma and ' INR' while keeping the cell as a number
            inr_format = workbook.add_format({'num_format': '#,##0.00 "INR"'})
            
            # Apply to Columns D (Subtotal), E (IGST), and F (Total)
            worksheet.set_column('D:F', 18, inr_format)
            # ----------------------------

        st.download_button(
            label="📥 Download Excel with INR Formatting",
            data=output.getvalue(),
            file_name="Invoices_Formatted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
