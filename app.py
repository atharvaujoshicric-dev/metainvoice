import streamlit as st
import pandas as pd
import pdfplumber
import re
import zipfile
import io
import gc

def clean_amt(value):
    if not value or value == "N/A":
        return 0.0
    clean_val = re.sub(r'[^\d.]', '', value)
    try:
        return float(clean_val)
    except ValueError:
        return 0.0

def extract_data_from_pdf(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

        project_match = re.search(r"Tax invoice for\s+(.*)", text)
        project_name = project_match.group(1).strip() if project_match else "N/A"

        inv_match = re.search(r"Invoice (?:no\.|ID)\s*([A-Z0-9-]+)", text)
        invoice_no = inv_match.group(1).strip() if inv_match else "N/A"

        sub_raw = re.search(r"Subtotal:\s+([\d,.]+)", text)
        igst_raw = re.search(r"IGST\s\(18%\):\s+([\d,.]+)", text)

        sub_val = clean_amt(sub_raw.group(1)) if sub_raw else 0.0
        igst_val = clean_amt(igst_raw.group(1)) if igst_raw else 0.0
        total_val = sub_val + igst_val

        date_match = re.search(r"(\d{1,2}\s[A-Za-z]{3}\s\d{4})", text)
        gstin_match = re.search(r"GSTIN:\s+([A-Z0-9]{15})", text)

        return {
            "Invoice Number": invoice_no,
            "Project Name": project_name,
            "Subtotal": sub_val,
            "IGST": igst_val,
            "Total": total_val,
            "Document Date": date_match.group(1) if date_match else "N/A",
            "GSTIN": gstin_match.group(1) if gstin_match else "NA"
        }
    except Exception as e:
        st.warning(f"Skipped a PDF due to error: {e}")
        return None


def process_outer_zip(outer_zip_file, status, progress):
    """
    Handles structure: outer.zip → folder/ → inner1.zip, inner2.zip → PDFs
    """
    all_data = []

    with zipfile.ZipFile(outer_zip_file) as outer_z:
        # Find all inner ZIPs inside the outer ZIP (regardless of folder depth)
        inner_zips = [f for f in outer_z.namelist() if f.lower().endswith('.zip')]

        if not inner_zips:
            # Fallback: maybe PDFs are directly inside
            pdf_files = [f for f in outer_z.namelist() if f.lower().endswith('.pdf')]
            status.text(f"No inner ZIPs found. Trying {len(pdf_files)} PDFs directly...")
            for pdf_path in pdf_files:
                with outer_z.open(pdf_path) as f:
                    result = extract_data_from_pdf(io.BytesIO(f.read()))
                    if result:
                        result["Source ZIP"] = "direct"
                        all_data.append(result)
            return all_data

        total_inner = len(inner_zips)
        status.text(f"Found {total_inner} inner ZIPs to process...")

        for z_idx, inner_zip_path in enumerate(inner_zips):
            inner_zip_name = inner_zip_path.split('/')[-1]  # Just the filename
            status.text(f"Opening inner ZIP {z_idx+1}/{total_inner}: {inner_zip_name}")

            try:
                # Read inner ZIP bytes from outer ZIP
                with outer_z.open(inner_zip_path) as inner_zip_bytes:
                    inner_zip_data = io.BytesIO(inner_zip_bytes.read())

                with zipfile.ZipFile(inner_zip_data) as inner_z:
                    pdf_files = [f for f in inner_z.namelist() if f.lower().endswith('.pdf')]
                    total_pdfs = len(pdf_files)

                    for i, pdf_path in enumerate(pdf_files):
                        status.text(f"[{inner_zip_name}] Processing PDF {i+1}/{total_pdfs}: {pdf_path.split('/')[-1]}")
                        try:
                            with inner_z.open(pdf_path) as f:
                                result = extract_data_from_pdf(io.BytesIO(f.read()))
                                if result:
                                    result["Source ZIP"] = inner_zip_name  # Track which inner ZIP it came from
                                    all_data.append(result)
                        except Exception as e:
                            st.warning(f"Skipped {pdf_path}: {e}")

                        if i % 10 == 0:
                            gc.collect()

            except zipfile.BadZipFile:
                st.warning(f"Skipped invalid ZIP: {inner_zip_name}")

            # Update outer progress
            progress.progress((z_idx + 1) / total_inner)

    return all_data


st.set_page_config(page_title="Invoice Processor", layout="wide")
st.title("📂 Multi-Zip Invoice Data Extractor")

uploaded_zips = st.file_uploader("Upload ZIP folders (Max 500MB)", type="zip", accept_multiple_files=True)

if uploaded_zips:
    if st.button("▶ Process Invoices"):
        all_data = []
        progress = st.progress(0)
        status = st.empty()

        for outer_zip in uploaded_zips:
            status.text(f"Opening {outer_zip.name}...")
            data = process_outer_zip(outer_zip, status, progress)
            all_data.extend(data)

        if all_data:
            df = pd.DataFrame(all_data)
            df.insert(0, 'Sr.no.', range(1, len(df) + 1))

            st.subheader("Extracted Data Preview")
            st.dataframe(df.style.format({
                "Subtotal": "{:,.2f} INR",
                "IGST": "{:,.2f} INR",
                "Total": "{:,.2f} INR"
            }))

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='InvoiceData')
                workbook = writer.book
                worksheet = writer.sheets['InvoiceData']
                inr_format = workbook.add_format({'num_format': '#,##0.00 "INR"'})
                worksheet.set_column('E:G', 18, inr_format)  # Subtotal, IGST, Total

            output.seek(0)

            st.success(f"✅ Processed {len(all_data)} invoices!")
            st.download_button(
                label="📥 Download Excel with INR Formatting",
                data=output.getvalue(),
                file_name="Invoices_Formatted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("No invoice data extracted. Check your ZIP structure or PDF format.")
