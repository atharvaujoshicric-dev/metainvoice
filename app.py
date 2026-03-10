import streamlit as st
import pandas as pd
import pdfplumber
import re
import zipfile
import io
import gc

# 1. Exact Columns from your sample_bills.csv
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
    if not value or value == "N/A": return 0.0
    clean_val = re.sub(r'[^\d.]', '', str(value))
    try: return float(clean_val)
    except: return 0.0

def extract_data_from_pdf(pdf_stream):
    try:
        with pdfplumber.open(pdf_stream) as pdf:
            text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
        
        # Regex extraction
        proj_m = re.search(r"Tax invoice for\s+(.*)", text)
        inv_m = re.search(r"Invoice (?:no\.|ID|ID\s)\s*([A-Z0-9-]+)", text)
        sub_m = re.search(r"Subtotal:\s+([\d,.]+)", text)
        tax_m = re.search(r"IGST\s\(18%\):\s+([\d,.]+)", text)
        date_m = re.search(r"(\d{1,2}\s[A-Za-z]{3}\s\d{4})", text)
        gstin_m = re.search(r"GSTIN:\s+([A-Z0-9]{15})", text)

        sub_v = clean_amt(sub_m.group(1)) if sub_m else 0.0
        tax_v = clean_amt(tax_m.group(1)) if tax_m else 0.0

        row = {col: "" for col in CSV_COLUMNS}
        row.update({
            'Bill Date': date_m.group(1) if date_m else "N/A",
            'Bill Number': inv_m.group(1) if inv_m else "N/A",
            'Project Name': proj_m.group(1).strip() if proj_m else "N/A",
            'SubTotal': sub_v,
            'Tax Amount': tax_v,
            'Total': sub_v + tax_v,
            'GST Identification Number (GSTIN)': gstin_m.group(1) if gstin_m else "NA",
            'Currency Code': 'INR',
            'Bill Status': 'Paid'
        })
        return row
    except:
        return None

def process_zip_recursively(zip_bytes):
    """Deeply scans for PDFs inside ZIPs and ZIPs inside ZIPs."""
    results = []
    with zipfile.ZipFile(zip_bytes) as z:
        for file_info in z.infolist():
            filename = file_info.filename
            # Ignore metadata folders
            if filename.startswith('__MACOSX') or '.DS_Store' in filename:
                continue
                
            # If it's an inner ZIP, open it and scan its contents
            if filename.lower().endswith('.zip'):
                with z.open(filename) as inner_zip_file:
                    inner_zip_data = io.BytesIO(inner_zip_file.read())
                    results.extend(process_zip_recursively(inner_zip_data))
            
            # If it's a PDF, extract the data
            elif filename.lower().endswith('.pdf'):
                with z.open(filename) as pdf_file:
                    pdf_data = io.BytesIO(pdf_file.read())
                    data = extract_data_from_pdf(pdf_data)
                    if data:
                        results.append(data)
    return results

# --- UI Setup ---
st.set_page_config(page_title="Nested Invoice Extractor", layout="wide")
st.title("📂 Deep-Scan Invoice Processor")
st.info("This version will scan for PDFs inside nested ZIP folders (ZIPs within ZIPs).")

uploads = st.file_uploader("Upload ZIP folders", type="zip", accept_multiple_files=True)

if uploads:
    all_data = []
    status = st.empty()
    
    for uploaded_file in uploads:
        status.text(f"Scanning {uploaded_file.name}...")
        all_data.extend(process_zip_recursively(io.BytesIO(uploaded_file.read())))
        gc.collect()

    if all_data:
        df = pd.DataFrame(all_data)
        df = df[CSV_COLUMNS] # Ensure column order
        df.insert(0, 'Sr.no.', range(1, len(df) + 1))
        
        st.subheader(f"✅ Found {len(all_data)} Invoices")
        st.dataframe(df.head(20), use_container_width=True)

        # Excel Generation
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Invoices')
            wb = writer.book
            ws = writer.sheets['Invoices']
            
            # Format with Commas and INR suffix
            inr_fmt = wb.add_format({'num_format': '#,##0.00 "INR"'})
            
            # Map Column Letters from the CSV_COLUMNS list
            # SubTotal is usually col AO, Tax Amount AI, Total AP
            ws.set_column('AL:AL', 15, inr_fmt) # Tax Amount
            ws.set_column('AS:AT', 18, inr_fmt) # SubTotal & Total
            
        st.download_button("📥 Download Final Excel Report", output.getvalue(), "Consolidated_Bills.xlsx")
