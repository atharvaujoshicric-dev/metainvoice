import streamlit as st
import pandas as pd
import pdfplumber
import re
import zipfile
import io

# --- 1. Page Config ---
st.set_page_config(page_title="Invoice Intel Pro", page_icon="💎", layout="wide")

# Modern Styling
st.markdown("""
    <style>
    .stMetric { border: 1px solid #e1e4e8; padding: 15px; border-radius: 10px; background-color: white; }
    .main { background-color: #f8f9fa; }
    </style>
""", unsafe_allow_html=True)

# --- 2. Helper Functions ---

def clean_amt(value):
    if not value or value == "N/A": return 0.0
    clean_val = re.sub(r'[^\d.]', '', str(value))
    try: return float(clean_val)
    except: return 0.0

@st.cache_data(show_spinner="Unpacking ZIP layers...")
def recursive_get_pdfs(zip_bytes):
    """Deep scan for PDFs in nested ZIP structures."""
    pdf_files = []
    def extract_recursive(data):
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            for info in z.infolist():
                if info.is_dir(): continue
                content = z.read(info.filename)
                if info.filename.lower().endswith('.pdf'):
                    pdf_files.append((info.filename, content))
                elif info.filename.lower().endswith('.zip'):
                    extract_recursive(content)
    extract_recursive(zip_bytes)
    return pdf_files

def parse_pdf(content):
    """Extracts data from PDF bytes with error handling."""
    try:
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            text = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])
        
        # Regex patterns
        patterns = {
            "inv": r"Invoice (?:no\.|ID)\s*([A-Z0-9-]+)",
            "proj": r"Tax invoice for\s+(.*)",
            "sub": r"Subtotal:\s+([\d,.]+)",
            "tax": r"IGST\s\(18%\):\s+([\d,.]+)",
            "date": r"(\d{1,2}\s[A-Za-z]{3}\s\d{4})",
            "gst": r"GSTIN:\s+([A-Z0-9]{15})"
        }
        
        res = {k: (re.search(v, text).group(1) if re.search(v, text) else "N/A") for k, v in patterns.items()}
        s_val, t_val = clean_amt(res['sub']), clean_amt(res['tax'])
        
        return {
            "Invoice #": res['inv'],
            "Project Name": res['proj'].strip(),
            "Subtotal": s_val,
            "IGST": t_val,
            "Total": s_val + t_val,
            "Date": res['date'],
            "GSTIN": res['gst']
        }
    except Exception:
        return None

# --- 3. UI Flow ---

st.title("💎 Invoice Intel Pro")
st.write("Upload folders or ZIPs. I'll handle the extraction and formatting.")

uploaded_zips = st.file_uploader("Upload ZIP files", type="zip", accept_multiple_files=True)

if uploaded_zips:
    all_pdfs = []
    for f in uploaded_zips:
        all_pdfs.extend(recursive_get_pdfs(f.read()))
    
    if all_pdfs:
        st.success(f"🔍 Found {len(all_pdfs)} PDFs. Extracting data...")
        
        # Progress Tracking
        progress_bar = st.progress(0)
        status_text = st.empty()
        results = []

        for idx, (name, content) in enumerate(all_pdfs):
            # Update Progress
            curr_progress = (idx + 1) / len(all_pdfs)
            progress_bar.progress(curr_progress)
            status_text.caption(f"Processing {idx+1}/{len(all_pdfs)}: **{name[:40]}**")
            
            data = parse_pdf(content)
            if data: results.append(data)
        
        status_text.empty()
        progress_bar.empty()

        if results:
            df = pd.DataFrame(results)
            
            # Dashboard
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Invoices", len(df))
            c2.metric("Net Amount", f"₹{df['Total'].sum():,.2f}")
            c3.metric("Tax Collected", f"₹{df['IGST'].sum():,.2f}")

            st.dataframe(df.style.format({"Subtotal": "{:,.2f}", "IGST": "{:,.2f}", "Total": "{:,.2f}"}), use_container_width=True)

            # Excel Generation
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Data')
                book = writer.book
                sheet = writer.sheets['Data']
                curr_fmt = book.add_format({'num_format': '#,##0.00 "INR"'})
                sheet.set_column('C:E', 15, curr_fmt)

            st.download_button("📥 Download Excel Report", out.getvalue(), "Invoices.xlsx", "application/vnd.ms-excel")
    else:
        st.warning("No PDFs found inside those ZIP files.")
