import streamlit as st
import pandas as pd
import pdfplumber
import re
import zipfile
import io

# --- 1. Page Configuration & UI Styling ---
st.set_page_config(page_title="Invoice Intel Pro", page_icon="🧾", layout="wide")

# Custom CSS for a modern, "SaaS-style" look
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .stButton>button { width: 100%; border-radius: 8px; height: 3em; background-image: linear-gradient(to right, #4facfe 0%, #00f2fe 100%); color: white; border: none; font-weight: bold; }
    .stDataFrame { border-radius: 10px; }
    .css-1kyx7g3 { background-color: #ffffff; } /* Sidebar color */
    </style>
    """, unsafe_allow_html=True)

# --- 2. Data Extraction Logic ---

def clean_amt(value):
    if not value or value == "N/A": return 0.0
    clean_val = re.sub(r'[^\d.]', '', str(value))
    try:
        return float(clean_val)
    except ValueError:
        return 0.0

def extract_from_pdf(pdf_stream):
    """Core extraction logic for a single PDF."""
    try:
        with pdfplumber.open(pdf_stream) as pdf:
            text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
        
        project_match = re.search(r"Tax invoice for\s+(.*)", text)
        inv_match = re.search(r"Invoice (?:no\.|ID)\s*([A-Z0-9-]+)", text)
        sub_raw = re.search(r"Subtotal:\s+([\d,.]+)", text)
        igst_raw = re.search(r"IGST\s\(18%\):\s+([\d,.]+)", text)
        date_match = re.search(r"(\d{1,2}\s[A-Za-z]{3}\s\d{4})", text)
        gstin_match = re.search(r"GSTIN:\s+([A-Z0-9]{15})", text)

        sub_val = clean_amt(sub_raw.group(1)) if sub_raw else 0.0
        igst_val = clean_amt(igst_raw.group(1)) if igst_raw else 0.0

        return {
            "Invoice Number": inv_match.group(1) if inv_match else "N/A",
            "Project Name": project_match.group(1).strip() if project_match else "N/A",
            "Subtotal": sub_val,
            "IGST": igst_val,
            "Total": sub_val + igst_val,
            "Date": date_match.group(1) if date_match else "N/A",
            "GSTIN": gstin_match.group(1) if gstin_match else "N/A"
        }
    except: return None

def get_all_files_recursively(zip_bytes):
    """Deep dives into ZIPs and nested ZIPs to find all PDFs."""
    files_to_process = []
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
        for info in z.infolist():
            if info.is_dir(): continue
            content = z.read(info.filename)
            if info.filename.lower().endswith('.pdf'):
                files_to_process.append((info.filename, content))
            elif info.filename.lower().endswith('.zip'):
                # Recursively add files from the nested zip
                files_to_process.extend(get_all_files_recursively(content))
    return files_to_process

# --- 3. Sidebar & File Upload ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/2611/2611030.png", width=80)
    st.title("Settings")
    st.info("The app automatically digs through folders and nested ZIP files.")
    uploaded_zips = st.file_uploader("Upload ZIP Folders", type="zip", accept_multiple_files=True)
    st.caption("Max limit: 500MB (via config)")

# --- 4. Processing Main View ---
st.title("⚡ Invoice Intelligence")
st.markdown("Automated PDF extraction for complex folder structures.")

if uploaded_zips:
    all_data = []
    
    # Pre-scan to count total PDFs (for accurate progress bar)
    with st.status("🔍 Scanning folders for PDFs...", expanded=False) as status:
        all_files = []
        for zip_file in uploaded_zips:
            all_files.extend(get_all_files_recursively(zip_file.read()))
        status.update(label=f"Found {len(all_files)} PDFs!", state="complete")

    if all_files:
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for idx, (name, content) in enumerate(all_files):
            # Calculate Percentage
            percent = (idx + 1) / len(all_files)
            progress_bar.progress(percent)
            status_text.text(f"Processing ({int(percent*100)}%): {name[:40]}...")
            
            # Extract
            result = extract_from_pdf(io.BytesIO(content))
            if result: all_data.append(result)

        # --- 5. Results & Visualization ---
        if all_data:
            df = pd.DataFrame(all_data)
            
            # Quick Stats Row
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Invoices", len(df))
            c2.metric("Total Value", f"₹ {df['Total'].sum():,.2f}")
            c3.metric("Total Tax (IGST)", f"₹ {df['IGST'].sum():,.2f}")

            # Data Display
            st.subheader("📋 Extracted Records")
            st.dataframe(df.style.format({
                "Subtotal": "{:,.2f}", "IGST": "{:,.2f}", "Total": "{:,.2f}"
            }), use_container_width=True)

            # Excel Export
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Invoices')
                workbook = writer.book
                worksheet = writer.sheets['Invoices']
                
                # Pro Formatting for Excel
                fmt_header = workbook.add_format({'bold': True, 'bg_color': '#1f77b4', 'font_color': 'white', 'border': 1})
                fmt_currency = workbook.add_format({'num_format': '#,##0.00 "INR"', 'border': 1})
                
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, fmt_header)
                worksheet.set_column('C:E', 18, fmt_currency)
                worksheet.set_column('A:B', 20)

            st.download_button(
                label="📥 Download Formatted Excel Report",
                data=output.getvalue(),
                file_name="Invoice_Data_Master.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("👈 Please upload your ZIP files in the sidebar to begin.")
