import streamlit as st
import pandas as pd
import pdfplumber
import re
import zipfile
import io

# Define the columns exactly as they appear in your sample CSV
CSV_COLUMNS = [
    'Bill Date', 'Bill Number', 'PurchaseOrder', 'Bill Status', 'Source of Supply', 
    'Destination of Supply', 'GST Treatment', 'GST Identification Number (GSTIN)', 
    'Is Inclusive Tax', 'TDS Percentage', 'TDS Amount', 'TDS Section Code', 
    'TDS Name', 'Vendor Name', 'Due Date', 'Currency Code', 'Exchange Rate', 
    'Attachment ID', 'Attachment Preview ID', 'Attachment Name', 'Attachment Type', 
    'Attachment Size', 'Item Name', 'SKU', 'Item Description', 'Account', 
    'Usage unit', 'Quantity', 'Rate', 'Adjustment', 'Item Type', 'Tax Name', 
    'Tax Percentage', 'Tax Amount', 'Tax Type', 'Item Exemption Code', 
    'Reverse Charge Tax Name', 'Reverse Charge Tax Rate', 'Reverse Charge Tax Type', 
    'Item Total', 'SubTotal', 'Total', 'Balance', 'Vendor Notes', 'Terms & Conditions', 
    'Payment Terms', 'Payment Terms Label', 'Is Billable', 'Customer Name', 
    'Project Name', 'Purchase Order Number', 'Is Discount Before Tax', 
    'Entity Discount Amount', 'Discount Account', 'Is Landed Cost', 'Warehouse Name', 
    'Branch Name', 'CF.Transporte_Name', 'TCS Tax Name', 'TCS Percentage', 
    'Nature Of Collection', 'TCS Amount', 'HSN/SAC', 'Supply Type', 'ITC Eligibility'
]

def clean_amt(value):
    """Removes currency symbols and commas, returning a float."""
    if not value or value == "N/A":
        return 0.0
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

    # 3. Financials
    sub_raw = re.search(r"Subtotal:\s+([\d,.]+)", text)
    igst_raw = re.search(r"IGST\s\(18%\):\s+([\d,.]+)", text)
    
    sub_val = clean_amt(sub_raw.group(1)) if sub_raw else 0.0
    igst_val = clean_amt(igst_raw.group(1)) if igst_raw else 0.0
    total_val = sub_val + igst_val 

    # 4. Date
    date_match = re.search(r"(\d{1,2}\s[A-Za-z]{3}\s\d{4})", text)
    bill_date = date_match.group(1) if date_match else "N/A"

    # 5. GSTIN
    gstin_match = re.search(r"GSTIN:\s+([A-Z0-9]{15})", text)
    gstin = gstin_match.group(1) if gstin_match else "NA"

    # Map to CSV Template
    row = {col: "" for col in CSV_COLUMNS}
    row['Bill Date'] = bill_date
    row['Bill Number'] = invoice_no
    row['Project Name'] = project_name
    row['SubTotal'] = sub_val
    row['Tax Amount'] = igst_val
    row['Tax Name'] = 'IGST18' if igst_val > 0 else ''
    row['Total'] = total_val
    row['GST Identification Number (GSTIN)'] = gstin
    row['Currency Code'] = 'INR'
    
    return row

st.set_page_config(page_title="Pro Invoice Extractor", layout="wide")
st.title("📊 Meta Invoice Aggregator")
st.caption("Upload multiple ZIP folders. Data will be mapped to your custom CSV format.")

uploaded_zips = st.file_uploader("Upload ZIP folders", type="zip", accept_multiple_files=True)

if uploaded_zips:
    all_rows = []
    for zip_file in uploaded_zips:
        with zipfile.ZipFile(zip_file) as z:
            pdfs = [f for f in z.namelist() if f.lower().endswith('.pdf')]
            for pdf_path in pdfs:
                with z.open(pdf_path) as f:
                    pdf_stream = io.BytesIO(f.read())
                    all_rows.append(extract_data_from_pdf(pdf_stream))

    if all_rows:
        df = pd.DataFrame(all_rows)
        # Ensure correct column order and add Sr.no.
        df = df[CSV_COLUMNS]
        df.insert(0, 'Sr.no.', range(1, len(df) + 1))
        
        st.subheader("Data Preview")
        st.dataframe(df.style.format({
            "SubTotal": "{:,.2f} INR", 
            "Tax Amount": "{:,.2f} INR", 
            "Total": "{:,.2f} INR"
        }))

        # Excel Export
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Bills')
            
            workbook  = writer.book
            worksheet = writer.sheets['Bills']
            
            # Format: #,##0.00 "INR" (Excel standard for commas + suffix)
            inr_fmt = workbook.add_format({'num_format': '#,##0.00 "INR"'})
            
            # SubTotal is column AO (41), Tax Amount is AI (35), Total is AP (42)
            # We also map indices: Sr.no is 0, Bill Date is 1, etc.
            # Let's apply to columns containing currency values:
            worksheet.set_column('AL:AL', 15, inr_fmt) # Tax Amount
            worksheet.set_column('AS:AT', 15, inr_fmt) # SubTotal and Total
            
        st.download_button(
            label="📥 Download Formatted Excel",
            data=output.getvalue(),
            file_name="Invoices_Export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
