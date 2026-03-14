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

        # Refined patterns based on your specific invoice layout
        patterns = {
            "inv": [
                r"Invoice\s*(?:No\.?|ID|#|Number)[:\s]*([A-Z0-9/_-]+)",
                r"Invoice\s*([A-Z0-9/_-]{4,})",
            ],
            "proj": [
                # Looks for the line starting with '1' right after the HSN code header
                r"HSN\s*code/SAC\s*code:?\s*\d*\s*\n\s*1\s+(.*)",
                # Fallback for different spacing
                r"Campaigns\s*–\s*Advertising\s*service[^\n]*\n\s*1\s+(.*)",
            ],
            "gstin": [
                r"GSTIN\s*[:\-]?\s*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})",
            ],
            "sub": [
                r"Sub\s*[Tt]otal\s*[:\-]?\s*([\d,. ]+)",
            ],
            "tax": [
                r"IGST\s*(?:\(\s*\d+\s*%\s*\))?\s*[:\-]?\s*([\d,. ]+)",
                r"GST\s*(?:\(\s*\d+\s*%\s*\))?\s*[:\-]?\s*([\d,. ]+)",
            ],
            "total": [
                r"Total\s+Amount\s*[:\-]?\s*([\d,. ]+)",
                r"Grand\s+Total\s*[:\-]?\s*([\d,. ]+)",
            ],
        }

        def find_first(key):
            for pat in patterns[key]:
                m = re.search(pat, text, re.IGNORECASE)
                if m:
                    # Clean up the project string if it captures the price at the end
                    res = m.group(1).strip()
                    if key == "proj":
                        res = re.sub(r'\s+[\d,.]+\s*INR.*$', '', res)
                    return res
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
            "Project": proj,
            "GSTIN": gstin,
            "Subtotal": s_val,
            "IGST/Tax": t_val,
            "Total": grand,
            "_raw_text": text[:2500],
        }
    except Exception as e:
        return {"File": filename, "Invoice #": "ERROR", "Project": str(e), "GSTIN": "N/A"}

def walk_zip(zip_bytes, parent_path="", results=None, status_fn=None):
    if results is None: results = []
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
            for entry in z.namelist():
                if "__MACOSX" in entry or entry.startswith(".") or entry.endswith("/"):
                    continue
                full_path = f"{parent_path}/{entry}" if parent_path else entry
                
                if entry.lower().endswith(".zip"):
                    with z.open(entry) as f:
                        walk_zip(f.read(), full_path, results, status_fn)
                elif entry.lower().endswith(".pdf"):
                    if status_fn: status_fn(f"📄 Extracting: {entry}")
                    with z.open(entry) as f:
                        results.append(extract_single_pdf(f.read(), full_path))
    except Exception as e:
        results.append({"File": parent_path, "Invoice #": "ZIP ERROR", "Project": str(e)})
    return results

# ── UI ──────────────────────────────────────────────────────────────────────────
st.title("📂 Universal Invoice Intel Pro")
st.markdown("Upload **PDFs**, **ZIPs**, or **Nested ZIPs**. The system will recursively find all invoices.")

debug_mode = st.sidebar.checkbox("🐛 Debug Mode (Raw Text)", value=False)

uploaded_files = st.file_uploader(
    "Drag and drop any files here", 
    type=["pdf", "zip"], 
    accept_multiple_files=True
)

if uploaded_files:
    all_data = []
    status = st.empty()
    
    for up_file in uploaded_files:
        file_bytes = up_file.read()
        if up_file.name.lower().endswith(".pdf"):
            status.text(f"📄 Processing: {up_file.name}")
            all_data.append(extract_single_pdf(file_bytes, up_file.name))
        elif up_file.name.lower().endswith(".zip"):
            status.text(f"📦 Unpacking ZIP: {up_file.name}")
            walk_zip(file_bytes, up_file.name, all_data, lambda m: status.text(m))

    status.empty()

    if all_data:
        df = pd.DataFrame(all_data)
        cols = ["File", "Invoice #", "Project", "GSTIN", "Subtotal", "IGST/Tax", "Total"]
        for c in cols: 
            if c not in df.columns: df[c] = "N/A"
        
        # UI Metrics
        m1, m2, m3 = st.columns(3)
        m1.metric("Invoices", len(df))
        m2.metric("Total Tax", f"₹{df['IGST/Tax'].sum():,.2f}")
        m3.metric("Grand Total", f"₹{df['Total'].sum():,.2f}")

        st.dataframe(df[cols], use_container_width=True, height=500)

        # Excel Export
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df[cols].to_excel(writer, index=False, sheet_name='Data')
        st.download_button("📥 Download Excel Report", output.getvalue(), "Invoice_Summary.xlsx")

        if debug_mode:
            for item in all_data:
                with st.expander(f"Raw Text: {item['File']}"):
                    st.text(item.get("_raw_text", "No text found"))

    else:
        st.warning("No invoices were found in the uploaded files.")
