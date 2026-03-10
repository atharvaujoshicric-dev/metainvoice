import streamlit as st
import pandas as pd
import pdfplumber
import re
import zipfile
import io
import gc

# ── Exact column order from sample_bills.csv ──────────────────────────────────
CSV_COLUMNS = [
    "Bill Date", "Bill Number", "PurchaseOrder", "Bill Status",
    "Source of Supply", "Destination of Supply", "GST Treatment",
    "GST Identification Number (GSTIN)", "Is Inclusive Tax",
    "TDS Percentage", "TDS Amount", "TDS Section Code", "TDS Name",
    "Vendor Name", "Due Date", "Currency Code", "Exchange Rate",
    "Attachment ID", "Attachment Preview ID", "Attachment Name",
    "Attachment Type", "Attachment Size",
    "Item Name", "SKU", "Item Description", "Account", "Usage unit",
    "Quantity", "Rate", "Adjustment", "Item Type",
    "Tax Name", "Tax Percentage", "Tax Amount", "Tax Type",
    "Item Exemption Code", "Reverse Charge Tax Name",
    "Reverse Charge Tax Rate", "Reverse Charge Tax Type", "Item Total",
    "SubTotal", "Total", "Balance",
    "Vendor Notes", "Terms & Conditions", "Payment Terms",
    "Payment Terms Label", "Is Billable", "Customer Name", "Project Name",
    "Purchase Order Number", "Is Discount Before Tax",
    "Entity Discount Amount", "Discount Account", "Is Landed Cost",
    "Warehouse Name", "Branch Name", "CF.Transporte_Name",
    "TCS Tax Name", "TCS Percentage", "Nature Of Collection", "TCS Amount",
    "HSN/SAC", "Supply Type", "ITC Eligibility"
]

def clean_amt(value):
    if not value:
        return 0.0
    clean_val = re.sub(r'[^\d.]', '', str(value))
    try:
        return float(clean_val)
    except ValueError:
        return 0.0

def extract_data_from_pdf(pdf_file, source_zip=""):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            text = "\n".join([
                page.extract_text() for page in pdf.pages if page.extract_text()
            ])

        # ── Bill Date ─────────────────────────────────────────────────────────
        date_match = re.search(r"(\d{1,2}\s[A-Za-z]{3,9}\s\d{4})", text)
        if not date_match:
            date_match = re.search(r"(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})", text)
        bill_date = date_match.group(1).strip() if date_match else ""

        # ── Bill / Invoice Number ─────────────────────────────────────────────
        inv_match = re.search(
            r"Invoice\s*(?:no\.?|number|ID|#)[:\s]*([A-Z0-9/_-]+)", text, re.IGNORECASE
        )
        if not inv_match:
            inv_match = re.search(r"Bill\s*(?:No\.?|Number)[:\s]*([A-Z0-9/_-]+)", text, re.IGNORECASE)
        invoice_no = inv_match.group(1).strip() if inv_match else ""

        # ── Vendor / Billed By ────────────────────────────────────────────────
        vendor_match = re.search(
            r"(?:From|Billed by|Vendor|Supplier)[:\s]+([A-Za-z0-9 ,.\-]+?)(?:\n|GSTIN|GST)", text, re.IGNORECASE
        )
        vendor_name = vendor_match.group(1).strip() if vendor_match else ""

        # ── Project Name / Campaign ───────────────────────────────────────────
        project_match = (
            re.search(r"Tax invoice for\s+(.+)", text, re.IGNORECASE) or
            re.search(r"Campaigns?\s*[-:]\s*(.+)", text, re.IGNORECASE) or
            re.search(r"Project\s*[-:]\s*(.+)", text, re.IGNORECASE)
        )
        project_name = project_match.group(1).strip() if project_match else ""

        # ── GSTIN ─────────────────────────────────────────────────────────────
        gstin_match = re.search(r"\b([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})\b", text)
        gstin = gstin_match.group(1) if gstin_match else ""

        # ── Source & Destination of Supply (state codes from GSTIN) ──────────
        # Seller GSTIN is usually the first one; buyer GSTIN second
        all_gstins = re.findall(r"\b([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})\b", text)
        STATE_CODES = {
            "01":"JK","02":"HP","03":"PB","04":"CH","05":"UT","06":"HR",
            "07":"DL","08":"RJ","09":"UP","10":"BR","11":"SK","12":"AR",
            "13":"NL","14":"MN","15":"MZ","16":"TR","17":"ML","18":"AS",
            "19":"WB","20":"JH","21":"OR","22":"CG","23":"MP","24":"GJ",
            "25":"DD","26":"DN","27":"MH","28":"AP","29":"KA","30":"GA",
            "31":"LD","32":"KL","33":"TN","34":"PY","35":"AN","36":"TS",
            "37":"AP"
        }
        source_of_supply = STATE_CODES.get(all_gstins[0][:2], "") if len(all_gstins) > 0 else ""
        destination_of_supply = STATE_CODES.get(all_gstins[1][:2], "") if len(all_gstins) > 1 else ""

        # ── HSN/SAC ───────────────────────────────────────────────────────────
        hsn_match = re.search(r"HSN[/SAC]*[:\s]+(\d{4,8})", text, re.IGNORECASE)
        if not hsn_match:
            hsn_match = re.search(r"SAC[:\s]+(\d{4,8})", text, re.IGNORECASE)
        hsn_sac = hsn_match.group(1) if hsn_match else ""

        # ── Item details ──────────────────────────────────────────────────────
        item_desc_match = re.search(
            r"(?:Description|Particulars|Item)[:\s]+([^\n]+)", text, re.IGNORECASE
        )
        item_description = item_desc_match.group(1).strip() if item_desc_match else ""

        qty_match = re.search(r"Qty[:\s]*([\d.]+)|Quantity[:\s]*([\d.]+)", text, re.IGNORECASE)
        quantity = float(qty_match.group(1) or qty_match.group(2)) if qty_match else 1

        rate_match = re.search(r"Rate[:\s]*([\d,.]+)", text, re.IGNORECASE)
        rate = clean_amt(rate_match.group(1)) if rate_match else 0.0

        # ── Financials ────────────────────────────────────────────────────────
        sub_raw   = re.search(r"Sub\s*[Tt]otal[:\s]+([\d,.]+)", text)
        igst_raw  = re.search(r"IGST\s*(?:\([\d.]+%\))?[:\s]+([\d,.]+)", text)
        cgst_raw  = re.search(r"CGST\s*(?:\([\d.]+%\))?[:\s]+([\d,.]+)", text)
        sgst_raw  = re.search(r"SGST\s*(?:\([\d.]+%\))?[:\s]+([\d,.]+)", text)
        total_raw = re.search(r"(?:Grand\s*)?Total[:\s]+([\d,.]+)", text)

        sub_val   = clean_amt(sub_raw.group(1))   if sub_raw   else 0.0
        igst_val  = clean_amt(igst_raw.group(1))  if igst_raw  else 0.0
        cgst_val  = clean_amt(cgst_raw.group(1))  if cgst_raw  else 0.0
        sgst_val  = clean_amt(sgst_raw.group(1))  if sgst_raw  else 0.0
        total_val = clean_amt(total_raw.group(1)) if total_raw else 0.0

        # Determine tax type: IGST (inter-state) or CGST+SGST (intra-state)
        if igst_val > 0:
            tax_name = "IGST18"
            tax_pct  = 18
            tax_amt  = igst_val
            tax_type = "ItemAmount"
        elif cgst_val > 0 or sgst_val > 0:
            tax_name = "GST18"
            tax_pct  = 18
            tax_amt  = cgst_val + sgst_val
            tax_type = "ItemAmount"
        else:
            tax_name = ""
            tax_pct  = ""
            tax_amt  = 0.0
            tax_type = ""

        if total_val == 0:
            total_val = sub_val + tax_amt

        # ── GST Treatment ─────────────────────────────────────────────────────
        gst_treatment = "business_gst" if gstin else "business_none"

        # ── PO Number ────────────────────────────────────────────────────────
        po_match = re.search(r"P\.?O\.?\s*(?:No\.?|Number|#)[:\s]*([A-Z0-9/_-]+)", text, re.IGNORECASE)
        po_number = po_match.group(1).strip() if po_match else ""

        # ── Place of Supply ───────────────────────────────────────────────────
        pos_match = re.search(r"Place of [Ss]upply[:\s]+([A-Za-z ]+)", text)
        place_of_supply = pos_match.group(1).strip() if pos_match else destination_of_supply

        # ── Build final row ───────────────────────────────────────────────────
        row = {col: "" for col in CSV_COLUMNS}
        row.update({
            "Bill Date":                         bill_date,
            "Bill Number":                       invoice_no,
            "Bill Status":                       "Open",
            "Source of Supply":                  source_of_supply,
            "Destination of Supply":             place_of_supply or destination_of_supply,
            "GST Treatment":                     gst_treatment,
            "GST Identification Number (GSTIN)": gstin,
            "Is Inclusive Tax":                  "FALSE",
            "Vendor Name":                       vendor_name,
            "Due Date":                          bill_date,
            "Currency Code":                     "INR",
            "Exchange Rate":                     1,
            "Item Description":                  item_description,
            "Account":                           "Cost of Goods Sold",
            "Quantity":                          quantity,
            "Rate":                              rate if rate > 0 else sub_val,
            "Adjustment":                        0,
            "Item Type":                         "goods",
            "Tax Name":                          tax_name,
            "Tax Percentage":                    tax_pct,
            "Tax Amount":                        tax_amt,
            "Tax Type":                          tax_type,
            "Item Total":                        sub_val,
            "SubTotal":                          sub_val,
            "Total":                             total_val,
            "Balance":                           total_val,
            "Payment Terms":                     0,
            "Payment Terms Label":               "Due on Receipt",
            "Is Billable":                       "FALSE",
            "Project Name":                      project_name,
            "Purchase Order Number":             po_number,
            "Is Discount Before Tax":            "TRUE",
            "Discount Account":                  "Purchase Discounts",
            "Is Landed Cost":                    "FALSE",
            "HSN/SAC":                           hsn_sac,
            "ITC Eligibility":                   "eligible",
        })
        return row

    except Exception as e:
        st.warning(f"Skipped a PDF: {e}")
        return None


def process_outer_zip(outer_zip_file, status, progress):
    all_data = []

    with zipfile.ZipFile(outer_zip_file) as outer_z:
        inner_zips = [f for f in outer_z.namelist() if f.lower().endswith('.zip')]

        if not inner_zips:
            pdf_files = [f for f in outer_z.namelist() if f.lower().endswith('.pdf')]
            status.text(f"No inner ZIPs — processing {len(pdf_files)} PDFs directly...")
            for pdf_path in pdf_files:
                with outer_z.open(pdf_path) as f:
                    result = extract_data_from_pdf(io.BytesIO(f.read()))
                    if result:
                        all_data.append(result)
            return all_data

        total_inner = len(inner_zips)
        for z_idx, inner_zip_path in enumerate(inner_zips):
            inner_zip_name = inner_zip_path.split('/')[-1]
            status.text(f"Opening inner ZIP {z_idx+1}/{total_inner}: {inner_zip_name}")
            try:
                with outer_z.open(inner_zip_path) as inner_zip_bytes:
                    inner_zip_data = io.BytesIO(inner_zip_bytes.read())

                with zipfile.ZipFile(inner_zip_data) as inner_z:
                    pdf_files = [f for f in inner_z.namelist() if f.lower().endswith('.pdf')]
                    total_pdfs = len(pdf_files)

                    for i, pdf_path in enumerate(pdf_files):
                        status.text(f"[{inner_zip_name}] PDF {i+1}/{total_pdfs}: {pdf_path.split('/')[-1]}")
                        try:
                            with inner_z.open(pdf_path) as f:
                                result = extract_data_from_pdf(io.BytesIO(f.read()), inner_zip_name)
                                if result:
                                    all_data.append(result)
                        except Exception as e:
                            st.warning(f"Skipped {pdf_path}: {e}")
                        if i % 10 == 0:
                            gc.collect()

            except zipfile.BadZipFile:
                st.warning(f"Skipped invalid ZIP: {inner_zip_name}")

            progress.progress((z_idx + 1) / total_inner)

    return all_data


# ── UI ─────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Invoice Processor", layout="wide")
st.title("📂 Multi-Zip Invoice Data Extractor")

uploaded_zips = st.file_uploader(
    "Upload ZIP folders (Max 500MB)", type="zip", accept_multiple_files=True
)

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

            # Enforce exact column order, fill missing with ""
            for col in CSV_COLUMNS:
                if col not in df.columns:
                    df[col] = ""
            df = df[CSV_COLUMNS]
            df.insert(0, "Sr.no.", range(1, len(df) + 1))

            st.subheader(f"Preview — {len(df)} invoices extracted")
            st.dataframe(df.style.format({
                "SubTotal":   "{:,.2f}",
                "Tax Amount": "{:,.2f}",
                "Total":      "{:,.2f}",
                "Balance":    "{:,.2f}",
                "Rate":       "{:,.2f}",
                "Item Total": "{:,.2f}",
            }, na_rep=""))

            # ── Excel export ──────────────────────────────────────────────────
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Bills')
                wb  = writer.book
                ws  = writer.sheets['Bills']

                header_fmt = wb.add_format({
                    'bold': True, 'bg_color': '#D9E1F2',
                    'border': 1, 'text_wrap': True, 'valign': 'top'
                })
                money_fmt = wb.add_format({'num_format': '#,##0.00'})
                pct_fmt   = wb.add_format({'num_format': '0.00'})

                money_cols = {
                    "SubTotal", "Total", "Balance", "Tax Amount", "TDS Amount",
                    "TCS Amount", "Rate", "Item Total", "Entity Discount Amount",
                    "Adjustment"
                }
                pct_cols = {"Tax Percentage", "TDS Percentage", "TCS Percentage",
                            "Exchange Rate", "Quantity"}

                for col_num, col_name in enumerate(df.columns):
                    worksheet_col = col_name
                    ws.write(0, col_num, worksheet_col, header_fmt)
                    if worksheet_col in money_cols:
                        ws.set_column(col_num, col_num, 16, money_fmt)
                    elif worksheet_col in pct_cols:
                        ws.set_column(col_num, col_num, 12, pct_fmt)
                    else:
                        ws.set_column(col_num, col_num, 20)

                ws.freeze_panes(1, 0)  # Freeze header row

            output.seek(0)
            status.text("✅ Done!")
            st.success(f"✅ Processed {len(df)} invoices!")
            st.download_button(
                label="📥 Download Excel (Zoho Bills Format)",
                data=output.getvalue(),
                file_name="Invoices_ZohoBills.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("No data extracted. Check your ZIP structure or PDF format.")
