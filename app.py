import streamlit as st
import pandas as pd
import pdfplumber
import re
import zipfile
import io

# --- 1. Page Config & Visuals ---
st.set_page_config(page_title="Invoice Intel Pro", page_icon="🧾", layout="wide")

st.markdown("""
    <style>
    .stProgress > div > div > div > div {
        background-image: linear-gradient(to right, #00dbde 0%, #fc00ff 100%);
    }
    .metric-container {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 10px;
        border: 1px solid #e1e4e8;
    }
    </style>
""", unsafe_allow_html=True)

def clean_amt(value):
    if not value: return 0.0
    val = re.sub(r'[^\d.]', '', str(value))
    try: return float(val)
    except: return 0.0

# --- 2. Recursive File Discovery ---
def get_file_list(uploaded_files):
    """Gathers all PDF paths first so we can calculate a real percentage."""
    all_pdfs = []
    
    def walk_zip(data, prefix=""):
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            for info in z.infolist():
                if info.is_dir() or info.filename.startswith('__MACOSX'): continue
                
                content = z.read(info.filename)
                if info.filename.lower().endswith('.pdf'):
                    all_pdfs.append((prefix + info.filename, content))
                elif info.filename.lower().endswith('.zip'):
                    walk_zip(content, prefix=info.filename + " > ")
                    
    for f in uploaded_files:
        walk_zip(f.read(), prefix=f.name + " > ")
    return all_pdfs

# --- 3. Main UI ---
st.title("🧾 Real-Time Invoice Extractor")
st.write("Processing multiple nested ZIPs with live tracking.")

files = st.file_uploader("Upload ZIP folders", type="zip", accept_multiple_files=True)

if files:
    # 1. Pre-scan (Fast)
    with st.spinner("🔍 Mapping folder structure..."):
        all_pdf_files = get_file_list(files)
    
    if all_pdf_files:
        total_files = len(all_pdf_files)
        st.write(f"📂 Found **{total_files}** invoices to process.")
        
        # 2. Setup Real-Time UI Elements
        progress_text = st.empty()
        progress_bar = st.progress(0)
        status_box = st.empty()
        
        all_results = []

        # 3. Process with Real-Time Updates
        for idx, (name, content) in enumerate(all_pdf_files):
            # Calculate Percentage
            percent_val = (idx + 1) / total_files
            
            # Update UI in real-time
            progress_bar.progress(percent_val)
            progress_text.markdown(f"**Processing:** `{name.split(' > ')[-1]}`")
            status_box.caption(f"File {idx + 1} of {total_files} ({int(percent_val * 100)}%)")
            
            try:
                with pdfplumber.open(io.BytesIO(content)) as pdf:
                    text = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])
                
                # Extraction
                inv = re.search(r"Invoice (?:no\.|ID)\s*([A-Z0-9-]+)", text)
                proj = re.search(r"Tax invoice for\s+(.*)", text)
                sub = re.search(r"Subtotal:\s+([\d,.]+)", text)
                tax = re.search(r"IGST\s\(18%\):\s+([\d,.]+)", text)
                
                s_val = clean_amt(sub.group(1)) if sub else 0.0
                t_val = clean_amt(tax.group(1)) if tax else 0.0
                
                all_results.append({
                    "Invoice #": inv.group(1) if inv else "N/A",
                    "Project": proj.group(1).strip() if proj else "N/A",
                    "Subtotal": s_val,
                    "IGST": t_val,
                    "Total": s_val + t_val,
                    "Location": name
                })
            except Exception as e:
                st.error(f"Could not read {name}: {e}")

        # Clear progress elements once done
        progress_text.empty()
        progress_bar.empty()
        status_box.success(f"🎯 Successfully processed {len(all_results)} invoices!")

        # --- 4. Final Results ---
        if all_results:
            df = pd.DataFrame(all_results)
            
            # Dashboard Metrics
            m1, m2, m3 = st.columns(3)
            with m1: st.metric("Invoices", len(df))
            with m2: st.metric("Total Revenue", f"₹{df['Total'].sum():,.2f}")
            with m3: st.metric("Tax Amount", f"₹{df['IGST'].sum():,.2f}")

            st.subheader("📊 Data Preview")
            st.dataframe(df, use_container_width=True)

            # Excel Export
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Invoices')
                workbook = writer.book
                sheet = writer.sheets['Invoices']
                fmt = workbook.add_format({'num_format': '#,##0.00 "INR"'})
                sheet.set_column('C:E', 18, fmt)
            
            st.download_button(
                label="📥 Download Formatted Excel",
                data=buf.getvalue(),
                file_name="Invoice_Master_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("No PDFs found in the uploaded ZIPs.")
