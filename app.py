import streamlit as st
import pandas as pd
import pdfplumber
import re
import zipfile
import io

st.set_page_config(page_title="Invoice Intel Pro", page_icon="💎", layout="wide")

# --- UI Styling ---
st.markdown("""
    <style>
    .stMetric { border: 1px solid #e1e4e8; padding: 10px; border-radius: 10px; background-color: #ffffff; }
    </style>
""", unsafe_allow_html=True)

def clean_amt(value):
    if not value: return 0.0
    val = re.sub(r'[^\d.]', '', str(value))
    try: return float(val)
    except: return 0.0

def extract_pdf_data(pdf_stream):
    """Memory-efficient extraction."""
    try:
        with pdfplumber.open(pdf_stream) as pdf:
            text = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])
        
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
            "Project Name": res['proj'].strip()[:50], # Truncate long names
            "Subtotal": s_val,
            "IGST": t_val,
            "Total": s_val + t_val,
            "Date": res['date'],
            "GSTIN": res['gst']
        }
    except: return None

# --- Main Logic ---
st.title("💎 Invoice Intel Pro")
st.caption("Deep-recursive extraction for high-volume uploads")

files = st.file_uploader("Upload ZIPs", type="zip", accept_multiple_files=True)

if files:
    all_data = []
    
    # We use a placeholder to avoid screen-jump
    status = st.status("🚀 Initializing deep scan...", expanded=True)
    
    # Recursive processing inside a flat loop to save memory
    zips_to_process = [f.read() for f in files]
    
    processed_count = 0
    
    # Process files one by one
    for zip_bytes in zips_to_process:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
            for filename in z.namelist():
                if filename.lower().endswith('.pdf'):
                    with z.open(filename) as f:
                        data = extract_pdf_data(io.BytesIO(f.read()))
                        if data:
                            all_data.append(data)
                            processed_count += 1
                            status.write(f"✅ Processed: {filename[:40]}...")

    if all_data:
        status.update(label=f"Extraction Complete! ({processed_count} files)", state="complete")
        df = pd.DataFrame(all_data)
        
        # Dashboard
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Files", len(df))
        col2.metric("Total Value", f"₹{df['Total'].sum():,.2f}")
        col3.metric("Total Tax", f"₹{df['IGST'].sum():,.2f}")
        
        st.dataframe(df, use_container_width=True)
        
        # Download
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 Download Excel Report", buf.getvalue(), "Invoices.xlsx")
    else:
        status.update(label="No PDFs found.", state="error")
