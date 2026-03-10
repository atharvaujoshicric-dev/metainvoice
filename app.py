import streamlit as st
import pandas as pd
import pdfplumber
import re
import zipfile
import io
import os

# --- 1. UI Enhancements & Config ---
st.set_page_config(
    page_title="Invoice Intelligence Pro", 
    page_icon="🧾", 
    layout="wide"
)

# Custom CSS for a modern look
st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #007bff; color: white; }
    .stDataFrame { border: 1px solid #e6e9ef; border-radius: 10px; }
    h1 { color: #1e3d59; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. Core Logic Functions ---

def clean_amt(value):
    if not value or value == "N/A": return 0.0
    clean_val = re.sub(r'[^\d.]', '', str(value))
    try:
        return float(clean_val)
    except ValueError:
        return 0.0

def extract_data_from_pdf(pdf_stream):
    """Extracts specific fields from a PDF byte stream."""
    try:
        with pdfplumber.open(pdf_stream) as pdf:
            text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
        
        # Regex patterns
        project_match = re.search(r"Tax invoice for\s+(.*)", text)
        inv_match = re.search(r"Invoice (?:no\.|ID)\s*([A-Z0-9-]+)", text)
        sub_raw = re.search(r"Subtotal:\s+([\d,.]+)", text)
        igst_raw = re.search(r"IGST\s\(18%\):\s+([\d,.]+)", text)
        date_match = re.search(r"(\d{1,2}\s[A-Za-z]{3}\s\d{4})", text)
        gstin_match = re.search(r"GSTIN:\s+([A-Z0-9]{15})", text)

        sub_val = clean_amt(sub_raw.group(1)) if sub_raw else 0.0
        igst_val = clean_amt(igst_raw.group(1)) if igst_raw else 0.0

        return {
            "Invoice Number": inv_match.group(1).strip() if inv_match else "N/A",
            "Project Name": project_match.group(1).strip() if project_match else "N/A",
            "Subtotal": sub_val,
            "IGST": igst_val,
            "Total": sub_val + igst_val,
            "Document Date": date_match.group(1) if date_match else "N/A",
            "GSTIN": gstin_match.group(1) if gstin_match else "N/A"
        }
    except Exception as e:
        return None

def process_nested_zip(zip_bytes, data_list):
    """Recursively finds and processes PDFs within ZIPs and nested ZIPs."""
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
        for filename in z.namelist():
            # Skip directories
            if filename.endswith('/'):
                continue
            
            with z.open(filename) as f:
                content = f.read()
                if filename.lower().endswith('.pdf'):
                    pdf_data = extract_data_from_pdf(io.BytesIO(content))
                    if pdf_data:
                        data_list.append(pdf_data)
                elif filename.lower().endswith('.zip'):
                    # Recurse into nested zip
                    process_nested_zip(content, data_list)

# --- 3. Streamlit UI ---

st.title("📂 Invoice Data Extractor")
st.markdown("Upload your ZIP files (even nested ones) to extract structured financial data.")

# 500MB Limit message (Streamlit handles actual enforcement via server config)
uploaded_zips = st.file_uploader(
    "Upload ZIP folders (Max 500MB total)", 
    type="zip", 
    accept_multiple_files=True
)

if uploaded_zips:
    all_extracted_data = []
    
    # Progress visualization
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    total_zips = len(uploaded_zips)
    
    with st.spinner("Digging through folders..."):
        for i, zip_file in enumerate(uploaded_zips):
            # Update progress
            percent_complete = int(((i) / total_zips) * 100)
            progress_bar.progress(percent_complete)
            status_text.text(f"Processing: {zip_file.name} ({percent_complete}%)")
            
            # Process the zip (including nested content)
            process_nested_zip(zip_file.read(), all_extracted_data)
        
        progress_bar.progress(100)
        status_text.text("Processing Complete! 100%")

    if all_extracted_data:
        df = pd.DataFrame(all_extracted_data)
        df.insert(0, 'Sr.no.', range(1, len(df) + 1))

        st.success(f"Successfully extracted {len(df)} invoices!")

        # --- Tabbed UI for Cleanliness ---
        tab1, tab2 = st.tabs(["📊 Data Preview", "💾 Export"])
        
        with tab1:
            st.dataframe(df.style.format({
                "Subtotal": "{:,.2f} INR", 
                "IGST": "{:,.2f} INR", 
                "Total": "{:,.2f} INR"
            }), use_container_width=True)

        with tab2:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='InvoiceData')
                workbook  = writer.book
                worksheet = writer.sheets['InvoiceData']
                
                # Excel Formatting
                inr_format = workbook.add_format({'num_format': '#,##0.00 "INR"', 'align': 'center'})
                header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
                
                # Apply formats
                worksheet.set_column('D:F', 18, inr_format)
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)

            st.download_button(
                label="🚀 Download Formatted Excel Report",
                data=output.getvalue(),
                file_name="Master_Invoice_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("No PDF invoices were found in the uploaded files.")
