import streamlit as st
import pandas as pd
import pdfplumber
import re
import zipfile
import io

def extract_data_from_pdf(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        # Extract text from all pages and join into one string
        full_text = ""
        for page in pdf.pages:
            full_text += page.extract_text() + "\n"
    
    # --- Extraction Logic ---
    
    # 1. Project Name (Extracts text after 'Tax invoice for')
    project_match = re.search(r"Tax invoice for\s+(.*)", full_text)
    project_name = project_match.group(1).strip() if project_match else "N/A"

    # 2. Invoice Number (Extracts the ID following 'Tax invoice ID' or the bottom 'Invoice no.')
    inv_match = re.search(r"Invoice no\.\s+([A-Z0-9-]+)", full_text)
    invoice_no = inv_match.group(1).strip() if inv_match else "N/A"

    # 3. Financials (Subtotal, IGST, Total)
    subtotal = re.search(r"Subtotal:\s+([\d,.]+)", full_text)
    igst = re.search(r"IGST\s\(18%\):\s+([\d,.]+)", full_text)
    total = re.search(r"Paid\s+([\d,.]+)", full_text)

    # 4. Document Date
    date_match = re.search(r"(\d{1,2}\s[A-Za-z]{3}\s\d{4})", full_text)

    # 5. GSTIN (Looking for the specific pattern in the footer)
    gstin_match = re.search(r"GSTIN:\s+([A-Z0-9]{15})", full_text)

    return {
        "Invoice Number": invoice_no,
        "Project Name": project_name,
        "Subtotal": subtotal.group(1) if subtotal else "0.00",
        "IGST": igst.group(1) if igst else "0.00",
        "Total": total.group(1) if total else "0.00",
        "Document Date": date_match.group(1) if date_match else "N/A",
        "GSTIN": gstin_match.group(1) if gstin_match else "NA"
    }

# --- Streamlit UI ---
st.set_page_config(page_title="Invoice Processor", layout="wide")
st.title("📊 Multi-Zip Invoice Data Extractor")
st.info("Upload your ZIP folders containing PDF invoices. The system will extract GSTIN and financial details into Excel.")

uploaded_zips = st.file_uploader("Upload ZIP folders", type="zip", accept_multiple_files=True)

if uploaded_zips:
    all_data = []
    
    for zip_file in uploaded_zips:
        with zipfile.ZipFile(zip_file) as z:
            # Filter for PDF files inside the zip
            pdfs = [f for f in z.namelist() if f.lower().endswith('.pdf')]
            
            for pdf_path in pdfs:
                with z.open(pdf_path) as f:
                    # Use BytesIO to make the file readable by pdfplumber
                    pdf_stream = io.BytesIO(f.read())
                    extracted_row = extract_data_from_pdf(pdf_stream)
                    all_data.append(extracted_row)

    if all_data:
        df = pd.DataFrame(all_data)
        # Add Serial Number
        df.insert(0, 'Sr.no.', range(1, len(df) + 1))
        
        # Reorder columns as requested
        column_order = ["Sr.no.", "Invoice Number", "Project Name", "Subtotal", "IGST", "Total", "Document Date", "GSTIN"]
        df = df[column_order]

        st.subheader("Extracted Data Preview")
        st.dataframe(df, use_container_width=True)

        # Generate Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Invoice Summary')
        
        st.download_button(
            label="📥 Download Excel Report",
            data=output.getvalue(),
            file_name="Consolidated_Invoices.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("No PDF files were detected inside the ZIP folders.")
