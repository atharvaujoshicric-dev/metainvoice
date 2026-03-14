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

        # Enhanced Patterns
        patterns = {
            "inv": [
                r"Invoice\s*(?:No\.?|ID|#|Number)[:\s]*([A-Z0-9/_-]+)",
                r"Inv\s*#?\s*[:\s]*([A-Z0-9/_-]+)",
            ],
            "proj": [
                # Specific request: Name under "Campaigns - Advertising service..."
                r"Campaigns\s*-\s*Advertising\s*service\s*[^\n]*\n\s*(.*)",
                r"Project\s*[:\-]\s*(.*)",
                r"Description\s*[:\-]\s*(.*)",
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
            "gstin": [
                # Standard Indian GSTIN Regex
                r"GSTIN\s*[:\-]?\s*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})",
            ]
        }

        def find_first(key):
            for pat in patterns[key]:
                m = re.search(pat, text, re.IGNORECASE)
                if m:
                    return m.group(1).strip()
            return "N/A"

        inv   = find_first("inv")
        proj  = find_first("proj")
        gstin = find_first("gstin")
        sub_s = find_first("sub")
        tax_s = find_first("tax")
        tot_s = find_first("total")

        s_val = clean_amt(sub_s)
        t_val = clean_amt(tax_s)
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
        return {
            "File": filename,
            "Invoice #": "ERROR",
            "Project": str(e),
            "GSTIN": "N/A",
            "Subtotal": 0.0,
            "IGST/Tax": 0.0,
            "Total": 0.0,
            "_raw_text": "",
        }

def walk_zip(zip_bytes, parent_path="", results=None, status_fn=None):
    if results is None: results = []
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
            for entry in z.namelist():
                if "__MACOSX" in entry or entry.startswith(".") or entry.endswith("/"):
                    continue
                full_path = f"{parent_path}/{entry}" if parent_path else entry
                if entry.lower().endswith(".pdf"):
                    if status_fn: status_fn(f"📄 {full_path[:70]}…")
                    with z.open(entry) as f:
                        pdf_bytes = f.read()
                    result = extract_single_pdf(pdf_bytes, filename=full_path)
                    if result: results.append(result)
                    gc.collect()
                elif entry.lower().endswith(".zip"):
                    if status_fn: status_fn(f"📦 Opening nested ZIP: {full_path[:70]}…")
                    with z.open(entry) as f:
                        nested_bytes = f.read()
                    walk_zip(nested_bytes, parent_path=full_path, results=results, status_fn=status_fn)
    except Exception as e:
        results.append({"File": parent_path or "unknown", "Invoice #": "ZIP ERROR", "Project": str(e)})
    return results

# ── UI ──────────────────────────────────────────────────────────────────────────
st.title("📂 Invoice Intel Pro")
st.caption("Recursively extracts PDF invoices from any ZIP structure.")

debug_mode = st.sidebar.checkbox("🐛 Debug mode (show raw PDF text)", value=False)
uploaded_files = st.file_uploader("Upload ZIP file(s)", type="zip", accept_multiple_files=True)

if uploaded_files:
    all_data = []
    status_text = st.empty()
    progress_bar = st.progress(0)

    for i, up_file in enumerate(uploaded_files):
        status_text.text(f"🔍 Scanning {up_file.name}…")
        zip_bytes = up_file.read()
        results = walk_zip(zip_bytes, parent_path=up_file.name, status_fn=lambda msg: status_text.text(msg))
        all_data.extend(results)
        progress_bar.progress((i + 1) / len(uploaded_files))

    status_text.empty()
    progress_bar.empty()

    if all_data:
        # Define specific column order including new GSTIN
        display_cols = ["File", "Invoice #", "Project", "GSTIN", "Subtotal", "IGST/Tax", "Total"]
        df_full = pd.DataFrame(all_data)
        
        # Ensure all columns exist even if error occurred
        for col in display_cols:
            if col not in df_full.columns: df_full[col] = "N/A"
            
        df = df_full[display_cols]

        # ── Metrics ──
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("📄 Invoices Found", len(df))
        c2.metric("💰 Total Value", f"₹{df['Total'].sum():,.2f}")
        c3.metric("🧾 Total Tax", f"₹{df['IGST/Tax'].sum():,.2f}")
        c4.metric("🏢 Unique GSTINs", df['GSTIN'].nunique() if 'GSTIN' in df else 0)

        st.divider()
        st.subheader("Extracted Records")
        st.dataframe(df, use_container_width=True, height=450)

        # ── Excel download ──
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name="Invoices")
        st.download_button("📥 Download Excel Report", data=buf.getvalue(), file_name="InvoiceReport.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        if debug_mode:
            st.divider()
            st.subheader("🐛 Raw Extracted Text")
            for row in all_data:
                with st.expander(f"📄 {row['File']}"):
                    st.text(row.get("_raw_text", "") or "(no text found)")
    else:
        st.error("❌ No PDFs found.")
