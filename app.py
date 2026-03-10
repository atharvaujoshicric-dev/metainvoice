import streamlit as st
import pandas as pd
import pdfplumber
import re
import zipfile
import io

# --- 1. UI Setup ---
st.set_page_config(page_title="Invoice Intel Pro", page_icon="🧾", layout="wide")

# Modern Professional Styling
st.markdown("""
    <style>
    .main { background-color: #f4f7f9; }
    .stProgress > div > div > div > div { background-image: linear-gradient(to right, #4facfe 0%, #00f2fe 100%); }
    div.stStatus { border-radius: 10px; border: 1px solid #d1d9e0; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. Logic with Caching (The Fix for Blinking) ---

def clean_amt(value):
    if not value or value == "N/A": return 0.0
    clean_val = re.sub(r'[^\d.]', '', str(value))
    try: return float(clean_val)
    except: return 0.0

@st.cache_data(show_spinner=False)
def get_all_files_recursively(zip_bytes):
    """Recursively finds all PDFs. Cached to prevent re-processing on every frame."""
    files_to_process = []
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
        for info in z.infolist():
            if info.is_dir(): continue
            content = z.read(info.filename)
            if info.filename.lower().endswith('.pdf'):
                files_to_process.append((info.filename, content))
            elif info.filename.lower().endswith('.zip'):
                files_to_process.extend(get_all_files_recursively(content))
    return files_to_process

def extract_from_pdf(pdf_content):
    """Extracts data from PDF bytes."""
    try:
        with pdfplumber.open(io.BytesIO(pdf_content)) as pdf:
            text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
        
        # Extraction patterns
        project = re.search(r"Tax invoice for\s+(.*)", text)
        inv = re.search(r"Invoice (?:no\.|ID)\s*([A-Z0-9-]+)", text)
        sub_raw = re.search(r"Subtotal:\s+([\d,.]+)", text)
        igst_raw = re.search(r"IGST\s\(18%\):\s+([\d,.]+)", text)
        date = re.search(r"(\d{1,2}\s[A-Za-z]{3}\s\d{4})", text)
        gstin = re.search(r"GSTIN:\s+([A-Z0-9]{15})", text)

        sub_val = clean_amt(sub_raw.group(1)) if sub_raw else 0.0
        igst_val = clean_amt(igst_raw.group(1)) if igst_raw else 0.0

        return {
            "Invoice Number": inv.group(1) if inv else "N/A",
            "Project Name": project.group(1).strip() if project else "N/A",
            "Subtotal": sub_val,
            "IGST": igst_val,
            "Total": sub_val + igst_val,
            "Date": date.group(1) if date else "N/A",
            "GSTIN": gstin.group(1) if gstin else "N/A"
        }
    except: return None

# --- 3. Main Interface ---
st.title("⚡ Invoice Intelligence Pro")
st.caption("Deep-scan recursive PDF extractor")

uploaded_zips = st.file_uploader("Upload ZIP folders (Supports nested ZIPs & Folders)", type="zip", accept_multiple_files=True)

if uploaded_zips:
    # Use a container to hold the UI so it doesn't jump around
    main_container = st.container()
    
    with main_container:
        # Step 1: Gather all files
        all_pdf_files = []
        for uploaded_file in uploaded_zips:
            # We pass the file name to help caching identify unique files
            file_data = uploaded_file.read()
            all_pdf_files.extend(get_all_files_recursively(file_data))

        if all_pdf_files:
            st.write(f"🔍 Found **{len(all_pdf_files)}** PDFs across all levels. Starting extraction...")
            
            # Step 2: Progress Tracking
            progress_bar = st.progress(0)
            status_text = st.empty()
            all_extracted_data = []

            for idx, (name, content) in enumerate(all_pdf_files):
                # Update UI
                progress = (idx + 1) / len(all_pdf_files)
                progress_bar.progress(progress)
                status_text.markdown(f"**Processing ({int(progress*100)}%):** `{name[:50]}...`")
                
                # Process
                res = extract_from_pdf(content)
                if res: all_extracted_data.append(res)

            # Step 3: Display Results
            if all_extracted_data:
                status_text.success(f"✅ Finished! Processed {len(all_extracted_data)} invoices.")
                df = pd.DataFrame(all_extracted_data)
                
                # Visual Dashboard
                col1, col2, col3 = st.columns(3)
                col1.metric("Total Invoices", len(df))
                col2.metric("Total Revenue", f"₹{df['Total'].sum():,.2f}")
                col3.metric("Avg. Invoice", f"₹{df['Total'].mean():,.2f}")

                st.subheader("Data Preview")
                st.dataframe(df, use_container_width=True)

                # Export Logic
                export_buffer = io.BytesIO()
                with pd.ExcelWriter(export_buffer, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                
                st.download_button(
                    label="📥 Download Master Excel",
                    data=export_buffer.getvalue(),
                    file_name="Invoice_Export.xlsx",
                    mime="application/vnd.ms-excel"
                )
        else:
            st.error("No PDFs found inside the uploaded ZIPs.")
