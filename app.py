import streamlit as st
import pandas as pd
import pdfplumber
import re
import zipfile
import io

# --- 1. Basic Config ---
st.set_page_config(page_title="Invoice Intel Pro", page_icon="🧾", layout="wide")

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

def process_pdfs_recursively(zip_bytes):
    """Memory-efficient recursive PDF finder."""
    extracted_data = []
    
    def walk_zip(data):
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            for filename in z.namelist():
                if filename.startswith('__MACOSX'): continue # Ignore Mac system files
                
                content = z.read(filename)
                if filename.lower().endswith('.pdf'):
                    # Process PDF
                    with pdfplumber.open(io.BytesIO(content)) as pdf:
                        text = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])
                    
                    # Regex Extraction
                    inv = re.search(r"Invoice (?:no\.|ID)\s*([A-Z0-9-]+)", text)
                    proj = re.search(r"Tax invoice for\s+(.*)", text)
                    sub = re.search(r"Subtotal:\s+([\d,.]+)", text)
                    tax = re.search(r"IGST\s\(18%\):\s+([\d,.]+)", text)
                    
                    s_val = clean_amt(sub.group(1)) if sub else 0.0
                    t_val = clean_amt(tax.group(1)) if tax else 0.0
                    
                    extracted_data.append({
                        "Invoice #": inv.group(1) if inv else "N/A",
                        "Project": proj.group(1).strip() if proj else "N/A",
                        "Subtotal": s_val,
                        "IGST": t_val,
                        "Total": s_val + t_val,
                        "File Source": filename.split('/')[-1]
                    })
                elif filename.lower().endswith('.zip'):
                    walk_zip(content) # Dig deeper

    walk_zip(zip_bytes)
    return extracted_data

# --- 2. Main UI ---
st.title("🧾 Invoice Intel Pro")
st.write("Deep-scan recursive PDF extractor")

files = st.file_uploader("Upload ZIP files", type="zip", accept_multiple_files=True)

if files:
    all_results = []
    with st.status("🔍 Deep scanning and extracting...", expanded=True) as status:
        for f in files:
            try:
                data = process_pdfs_recursively(f.read())
                all_results.extend(data)
            except Exception as e:
                st.error(f"Error reading {f.name}: {e}")
        status.update(label="Processing Complete!", state="complete")

    if all_results:
        df = pd.DataFrame(all_results)
        
        # Dashboard
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Invoices", len(df))
        c2.metric("Total (INR)", f"₹{df['Total'].sum():,.2f}")
        c3.metric("Tax Collected", f"₹{df['IGST'].sum():,.2f}")

        st.dataframe(df, use_container_width=True)

        # Excel Export
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Data')
            workbook = writer.book
            sheet = writer.sheets['Data']
            fmt = workbook.add_format({'num_format': '#,##0.00 "INR"'})
            sheet.set_column('C:E', 15, fmt)
        
        st.download_button("📥 Download Excel", buf.getvalue(), "Invoices.xlsx")
