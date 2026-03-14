import streamlit as st
import pandas as pd
import pdfplumber
import re
import zipfile
import io
import gc

st.set_page_config(page_title="Invoice Intel Pro", page_icon="🧾", layout="wide")

def clean_amt(value):
    if not value: return 0.0
    val = re.sub(r'[^\d.]', '', str(value))
    try: return float(val)
    except: return 0.0

def extract_single_pdf(content, filename=""):
    try:
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            text = ""
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"

        patterns = {
            "inv": [
                r"Invoice\s*(?:No\.?|ID|#|Number)[:\s]*([A-Z0-9/_-]+)",
                r"Inv\s*#?\s*[:\s]*([A-Z0-9/_-]+)",
            ],
            "proj": [
                # Specifically targets the name under the Campaign/HSN section
                r"Campaigns\s*-\s*Advertising\s*service\s*HSN\s*code/SAC\s*code\s*\n\s*(.*)",
                r"Description\s*[:\-]\s*(.*)",
            ],
            "gstin": [
                # Extract Indian GSTIN format
                r"GSTIN\s*[:\-]?\s*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})",
            ],
            "sub": [
                r"Sub\s*[Tt]otal\s*[:\-]?\s*([\d,. ]+)",
                r"Net\s+Amount\s*[:\-]?\s*([\d,. ]+)",
            ],
            "tax": [
                r"IGST\s*(?:\(\s*\d+\s*%\s*\))?\s*[:\-]?\s*([\d,. ]+)",
                r"GST\s*(?:\(\s*\d+\s*%\s*\))?\s*[:\-]?\s*([\d,. ]+)",
                r"Tax\s+Amount\s*[:\-]?\s*([\d,. ]+)",
            ],
            "total": [
                r"Grand\s+Total\s*[:\-]?\s*([\d,. ]+)",
                r"Total\s+Amount\s*[:\-]?\s*([\d,. ]+)",
                r"\bTotal\b\s*[:\-]?\s*([\d,. ]+)",
            ],
        }

        def find_first(key):
            for pat in patterns[key]:
                m = re.search(pat, text, re.IGNORECASE)
                if m: return m.group(1).strip()
            return "N/A"

        inv   = find_first("inv")
        proj  = find_first("proj")
        gstin = find_first("gstin")
        s_val = clean_amt(find_first("sub"))
        t_val = clean_amt(find_first("tax"))
        tot_s = find_first("total")
        grand = clean_amt(tot_s) if tot_s != "N/A" else s_val + t_val

        return {
            "File": filename,
            "Invoice #": inv,
            "Project": proj[:100] if proj != "N/A" else "N/A",
            "GSTIN": gstin,
            "Subtotal": s_val,
            "IGST/Tax": t_val,
            "Total": grand,
            "_raw_text": text[:2000],
        }
    except Exception as e:
        return {"File": filename, "Invoice #": "ERROR", "Project": str(e), "GSTIN": "N/A", "Subtotal": 0.0, "IGST/Tax": 0.0, "Total": 0.0}

def walk_zip(zip_bytes, parent_path="", results=None, status_fn=None):
    if results is None: results = []
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
            for entry in z.namelist():
                if "__MACOSX" in entry or entry.startswith(".") or entry.endswith("/"):
                    continue
                full_path = f"{parent_path}/{entry}" if parent_path else entry
                
                # If nested ZIP, dive deeper
                if entry.lower().endswith(".zip"):
                    with z.open(entry) as f:
                        walk_zip(f.read(), full_path, results, status_fn)
                # If PDF, extract
                elif entry.lower().endswith(".pdf"):
                    if status_fn: status_fn(f"📄 Processing: {entry}")
                    with z.open(entry) as f:
                        results.append(extract_single_pdf(f.read(), full_path))
                    gc.collect()
    except Exception as e:
        results.append({"File": parent_path, "Invoice #": "ZIP ERROR", "Project": str(e)})
    return results

# ── UI ──────────────────────────────────────────────────────────────────────────
st.title("📂 Universal Invoice Intel")
st.caption("Upload PDFs, ZIPs, or nested ZIPs. We'll find every invoice.")

debug_mode = st.sidebar.checkbox("🐛 Debug mode", value=False)

uploaded_files = st.file_uploader(
    "Drop files here (PDF or ZIP)", 
    type=["pdf", "zip"], 
    accept_multiple_files=True
)

if uploaded_files:
    all_data = []
    status_text = st.empty()
    
    for up_file in uploaded_files:
        file_bytes = up_file.read()
        
        if up_file.name.lower().endswith(".pdf"):
            status_text.text(f"📄 Processing direct PDF: {up_file.name}")
            all_data.append(extract_single_pdf(file_bytes, up_file.name))
        
        elif up_file.name.lower().endswith(".zip"):
            status_text.text(f"📦 Unpacking ZIP: {up_file.name}")
            walk_zip(file_bytes, up_file.name, all_data, lambda m: status_text.text(m))

    status_text.empty()

    if all_data:
        df = pd.DataFrame(all_data)
        display_cols = ["File", "Invoice #", "Project", "GSTIN", "Subtotal", "IGST/Tax", "Total"]
        # Ensure columns exist
        for c in display_cols: 
            if c not in df.columns: df[c] = "N/A"
        
        st.dataframe(df[display_cols], use_container_width=True)
        
        # Total Summary Metrics
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Invoices", len(df))
        c2.metric("Total Tax (IGST)", f"₹{df['IGST/Tax'].sum():,.2f}")
        c3.metric("Grand Total", f"₹{df['Total'].sum():,.2f}")

        # Excel Download
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
            df[display_cols].to_excel(writer, index=False)
        st.download_button("📥 Download Report", buf.getvalue(), "Invoice_Report.xlsx")
        
        if debug_mode:
            for row in all_data:
                with st.expander(f"Raw Text: {row['File']}"):
                    st.text(row.get("_raw_text", "No text"))
    else:
        st.warning("No data extracted. Please check your files.")
