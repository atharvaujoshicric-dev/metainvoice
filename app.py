import streamlit as st
import pandas as pd
import pdfplumber
import re
import zipfile
import io
import gc  # Garbage Collector to clear memory

# --- 1. Page Config (2026 Compliant) ---
st.set_page_config(page_title="Invoice Intel Pro", page_icon="🧾", layout="wide")

def clean_amt(value):
    if not value: return 0.0
    val = re.sub(r'[^\d.]', '', str(value))
    try: return float(val)
    except: return 0.0

# --- 2. Memory-Safe Extraction ---
def extract_single_pdf(content):
    """Processes one PDF and immediately returns data to keep RAM low."""
    try:
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() or ""
        
        patterns = {
            "inv": r"Invoice (?:no\.|ID)\s*([A-Z0-9-]+)",
            "proj": r"Tax invoice for\s+(.*)",
            "sub": r"Subtotal:\s+([\d,.]+)",
            "tax": r"IGST\s\(18%\):\s+([\d,.]+)"
        }
        
        res = {k: (re.search(v, text).group(1) if re.search(v, text) else "N/A") for k, v in patterns.items()}
        s_val, t_val = clean_amt(res['sub']), clean_amt(res['tax'])
        
        return {
            "Invoice #": res['inv'],
            "Project": res['proj'].strip(),
            "Subtotal": s_val,
            "IGST": t_val,
            "Total": s_val + t_val
        }
    except:
        return None

# --- 3. UI Implementation ---
st.title("📂 Heavy-Duty Invoice Extractor")
st.info("This version is optimized for large 500MB uploads on limited memory.")

uploaded_files = st.file_uploader("Upload ZIP files", type="zip", accept_multiple_files=True)

if uploaded_files:
    all_data = []
    
    # Placeholders for real-time UI
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    # Process files one by one without pre-loading
    for up_file in uploaded_files:
        with zipfile.ZipFile(up_file) as z:
            # Filter only PDFs and ZIPs to avoid system junk
            target_files = [f for f in z.namelist() if f.lower().endswith(('.pdf', '.zip')) and "__MACOSX" not in f]
            
            for idx, filename in enumerate(target_files):
                # Update UI
                progress = (idx + 1) / len(target_files)
                progress_bar.progress(progress)
                status_text.text(f"Processing: {filename[:50]}...")

                if filename.lower().endswith('.pdf'):
                    with z.open(filename) as f:
                        result = extract_single_pdf(f.read())
                        if result:
                            all_data.append(result)
                
                # Manual memory cleanup
                if idx % 10 == 0:
                    gc.collect()

    status_text.empty()
    progress_bar.empty()

    if all_data:
        df = pd.DataFrame(all_data)
        
        # Metrics
        c1, c2, c3 = st.columns(3)
        c1.metric("Invoices", len(df))
        c2.metric("Total Value", f"₹{df['Total'].sum():,.2f}")
        c3.metric("Total Tax", f"₹{df['IGST'].sum():,.2f}")

        # Dataframe with 2026 'width' syntax
        st.dataframe(df, width=1500) # Using fixed width or 'stretch' equivalent

        # Excel Download
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 Download Excel", buf.getvalue(), "Report.xlsx")
