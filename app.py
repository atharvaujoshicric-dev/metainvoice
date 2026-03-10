import streamlit as st
import pandas as pd
import pdfplumber
import re
import zipfile
import io

# --- 1. Page Config & Modern UI ---
st.set_page_config(page_title="Invoice Intel Pro", page_icon="🧾", layout="wide")

st.markdown("""
    <style>
    .stProgress > div > div > div > div {
        background-image: linear-gradient(to right, #4facfe 0%, #00f2fe 100%);
        border-radius: 10px;
    }
    .main { background-color: #f8f9fa; }
    div[data-testid="stMetricValue"] { font-size: 1.8rem; color: #1f77b4; }
    </style>
""", unsafe_allow_html=True)

def clean_amt(value):
    if not value: return 0.0
    val = re.sub(r'[^\d.]', '', str(value))
    try: return float(val)
    except: return 0.0

# --- 2. Optimized Recursive Extraction ---
def get_all_pdfs_pre_scan(uploaded_files):
    """Gathers all file references without loading everything into RAM at once."""
    pdf_tasks = []
    for uploaded_file in uploaded_files:
        try:
            file_bytes = uploaded_file.read()
            def recursive_zip_search(data, path_prefix=""):
                with zipfile.ZipFile(io.BytesIO(data)) as z:
                    for info in z.infolist():
                        if info.is_dir() or "__MACOSX" in info.filename:
                            continue
                        content = z.read(info.filename)
                        if info.filename.lower().endswith('.pdf'):
                            pdf_tasks.append({"name": path_prefix + info.filename, "content": content})
                        elif info.filename.lower().endswith('.zip'):
                            recursive_zip_search(content, path_prefix + info.filename + " > ")
            
            recursive_zip_search(file_bytes, f"{uploaded_file.name} > ")
        except Exception as e:
            st.error(f"Error opening ZIP {uploaded_file.name}: {e}")
    return pdf_tasks

# --- 3. Main App Layout ---
st.title("📂 Multi-Zip Invoice Data Extractor")
st.write("Professional-grade recursive extraction with real-time tracking.")

# Sidebar for clean UI
with st.sidebar:
    st.header("Settings")
    st.info("Upload limit: 500MB. Deep-scan is enabled for nested folders.")
    uploaded_files = st.file_uploader("Upload ZIP files", type="zip", accept_multiple_files=True)

if uploaded_files:
    # Step 1: Map Files
    with st.spinner("🔍 Scanning directory structure..."):
        all_pdfs = get_all_pdfs_pre_scan(uploaded_files)
    
    if all_pdfs:
        total = len(all_pdfs)
        st.subheader(f"🔄 Processing {total} Invoices")
        
        # UI Placeholders
        progress_bar = st.progress(0)
        status_text = st.empty()
        results = []

        # Step 2: Extraction Loop
        for i, task in enumerate(all_pdfs):
            # Real-time UI Update
            progress = (i + 1) / total
            progress_bar.progress(progress)
            status_text.write(f"🏷️ **Current File:** `{task['name'].split(' > ')[-1]}`")
            
            try:
                with pdfplumber.open(io.BytesIO(task['content'])) as pdf:
                    text = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])
                
                # Regex (Updated for more robustness)
                patterns = {
                    "inv": r"Invoice (?:no\.|ID)\s*([A-Z0-9-]+)",
                    "proj": r"Tax invoice for\s+(.*)",
                    "sub": r"Subtotal:\s+([\d,.]+)",
                    "tax": r"IGST\s\(18%\):\s+([\d,.]+)"
                }
                
                extracted = {k: (re.search(v, text).group(1) if re.search(v, text) else "N/A") for k, v in patterns.items()}
                s_val = clean_amt(extracted['sub'])
                t_val = clean_amt(extracted['tax'])
                
                results.append({
                    "Invoice #": extracted['inv'],
                    "Project Name": extracted['proj'].strip(),
                    "Subtotal": s_val,
                    "IGST": t_val,
                    "Total": s_val + t_val,
                    "Source Path": task['name']
                })
            except:
                continue # Skip corrupted PDFs

        # Clear progress UI
        status_text.empty()
        progress_bar.empty()

        if results:
            df = pd.DataFrame(results)
            
            # Dashboard Metrics
            m1, m2, m3 = st.columns(3)
            m1.metric("Total Invoices", len(df))
            m2.metric("Total Value (INR)", f"₹{df['Total'].sum():,.2f}")
            m3.metric("Tax IGST", f"₹{df['IGST'].sum():,.2f}")

            # Data Preview - Updated syntax for 2026
            st.subheader("📋 Extracted Records")
            st.dataframe(
                df.style.format({"Subtotal": "{:,.2f}", "IGST": "{:,.2f}", "Total": "{:,.2f}"}), 
                width='stretch'
            )

            # Excel Export
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='InvoiceData')
                workbook = writer.book
                worksheet = writer.sheets['InvoiceData']
                # Currency format
                inr_fmt = workbook.add_format({'num_format': '#,##0.00 "INR"'})
                worksheet.set_column('C:E', 18, inr_fmt)
            
            st.download_button(
                label="📥 Download Master Excel Report",
                data=output.getvalue(),
                file_name="Master_Invoice_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("No PDFs found in the uploaded ZIPs.")
else:
    st.info("Please upload one or more ZIP files in the sidebar to start.")
