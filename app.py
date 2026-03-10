import streamlit as st
import pandas as pd
import pdfplumber
import re
import zipfile
import io

# --- 1. Page Config ---
st.set_page_config(page_title="Invoice Intel", page_icon="💎", layout="wide")

# UI Polish
st.markdown("""
    <style>
    .block-container { padding-top: 2rem; }
    .stMetric { border: 1px solid #e1e4e8; padding: 10px; border-radius: 8px; background: white; }
    </style>
""", unsafe_allow_html=True)

# --- 2. Logic & Caching ---

def clean_amt(value):
    if not value: return 0.0
    clean_val = re.sub(r'[^\d.]', '', str(value))
    try: return float(clean_val)
    except: return 0.0

@st.cache_data(show_spinner="Extracting files from ZIPs...")
def get_all_pdfs(uploaded_files):
    """Recursively finds all PDFs and returns a list of (filename, bytes)."""
    pdf_list = []
    
    def process_zip(zip_data):
        with zipfile.ZipFile(io.BytesIO(zip_data)) as z:
            for info in z.infolist():
                if info.is_dir(): continue
                content = z.read(info.filename)
                if info.filename.lower().endswith('.pdf'):
                    pdf_list.append((info.filename, content))
                elif info.filename.lower().endswith('.zip'):
                    process_zip(content) # Recursive call

    for uploaded_file in uploaded_files:
        process_zip(uploaded_file.read())
    return pdf_list

def extract_data(pdf_content):
    try:
        with pdfplumber.open(io.BytesIO(pdf_content)) as pdf:
            text = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])
        
        # Robust Regex
        patterns = {
            "Invoice Number": r"Invoice (?:no\.|ID)\s*([A-Z0-9-]+)",
            "Project": r"Tax invoice for\s+(.*)",
            "Subtotal": r"Subtotal:\s+([\d,.]+)",
            "IGST": r"IGST\s\(18%\):\s+([\d,.]+)",
            "Date": r"(\d{1,2}\s[A-Za-z]{3}\s\d{4})",
            "GSTIN": r"GSTIN:\s+([A-Z0-9]{15})"
        }
        
        data = {k: (re.search(v, text).group(1) if re.search(v, text) else "N/A") for k, v in patterns.items()}
        
        # Calculations
        sub = clean_amt(data["Subtotal"])
        igst = clean_amt(data["IGST"])
        
        return {
            "Invoice #": data["Invoice Number"],
            "Project Name": data["Project"].strip(),
            "Subtotal": sub,
            "IGST": igst,
            "Total": sub + igst,
            "Date": data["Date"],
            "GSTIN": data["GSTIN"]
        }
    except: return None

# --- 3. UI ---
st.title("💎 Invoice Intel Pro")
st.info("Upload your folders/ZIPs below. I'll dig through all nested layers for you.")

files = st.file_uploader("Drop ZIP files here", type="zip", accept_multiple_files=True)

if files:
    # Get all PDFs once (Cached)
    all_pdfs = get_all_pdfs(files)
    
    if all_pdfs:
        st.write(f"📂 Found **{len(all_pdfs)}** PDF invoices. Starting deep-scan...")
        
        results = []
        prog_container = st.empty()
        bar = st.progress(0)
        
        for i, (name, content) in enumerate(all_pdfs):
            # Calculate and update progress
            percent = (i + 1) / len(all_pdfs)
            bar.progress(percent)
            prog_container.caption(f"Processing {i+1}/{len(all_pdfs)}: **{name[:40]}...**")
            
            data = extract_data(content)
            if data: results.append(data)
        
        bar.empty()
        prog_container.empty()

        if results:
            df = pd.DataFrame(results)
            
            # Dashboard Metrics
            m1, m2, m3 = st.columns(3)
            m1.metric("Invoices", len(df))
            m2.metric("Total (INR)", f"₹{df['Total'].sum():,.2f}")
            m3.metric("Tax Amount", f"₹{df['IGST'].sum():,.2f}")

            # Elegant Table
            st.dataframe(df.style.format({"Subtotal": "{:,.2f}", "IGST": "{:,.2f}", "Total": "{:,.2f}"}), use_container_width=True)

            # Excel Download
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Report')
                # Currency format for Excel
                workbook = writer.book
                fmt = workbook.add_format({'num_format': '#,##0.00 "INR"'})
                writer.sheets['Report'].set_column('C:E', 15, fmt)

            st.download_button("📥 Download Master Excel Report", buf.getvalue(), "Invoice_Report.xlsx", "application/vnd.ms-excel")
    else:
        st.warning("No PDFs found in those ZIPs.")
