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

def process_pdf(pdf_bytes):
    """Processes a single PDF byte stream and returns a data dictionary."""
    try:
        with pdfplumber.open(pdf_bytes) as pdf:
            text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
        
        # Extract Fields
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
            'Vendor Name': "Meta / Facebook India",
            'Currency Code': "INR",
            'SubTotal': sub_v,
            'Tax Amount': tax_v,
            'Total': sub_v + tax_v,
            'GST Identification Number (GSTIN)': gstin_m.group(1) if gstin_m else "NA",
            'Bill Status': 'Paid'
        })
        return row
    except:
        return None

# --- UI Setup ---
st.set_page_config(page_title="Invoice Processor", layout="wide")
st.title("📂 High-Volume Invoice Processor")
st.markdown("Processing limit is set to **500MB**. If it stays stuck on 'Uploading', check your upload speed.")

# Step 1: Multiple Uploads
zips = st.file_uploader("Drop ZIP files here", type="zip", accept_multiple_files=True)

if zips:
    all_data = []
    # Create placeholders to avoid screen flickering
    prog_bar = st.progress(0)
    msg = st.empty()

    for z_file in zips:
        with zipfile.ZipFile(z_file) as z:
            # Find all PDFs including those in subfolders
            pdf_names = [n for n in z.namelist() if n.lower().endswith('.pdf') and not n.startswith('__MACOSX')]
            
            for i, name in enumerate(pdf_names):
                msg.text(f"Extracting: {name.split('/')[-1]}")
                with z.open(name) as f:
                    res = process_pdf(io.BytesIO(f.read()))
                    if res: all_data.append(res)
                
                prog_bar.progress((i + 1) / len(pdf_names))
                if i % 20 == 0: gc.collect() # Frequent cleanup for 300MB+ files

    if all_data:
        df = pd.DataFrame(all_data)
        df.insert(0, 'Sr.no.', range(1, len(df) + 1))
        
        st.subheader(f"✅ Processed {len(all_data)} Invoices")
        st.dataframe(df[['Sr.no.', 'Bill Number', 'Project Name', 'Total', 'GST Identification Number (GSTIN)']], use_container_width=True)

        # Download
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Invoices')
            wb = writer.book
            ws = writer.sheets['Invoices']
            curr_fmt = wb.add_format({'num_format': '#,##0.00 "INR"'})
            # Applying format based on CSV column indices (approx AL, AS, AT)
            ws.set_column('AL:AL', 15, curr_fmt) # Tax Amount
            ws.set_column('AS:AT', 18, curr_fmt) # SubTotal & Total
            
        st.download_button("📥 Download Excel Report", out.getvalue(), "Invoices_Master.xlsx")
