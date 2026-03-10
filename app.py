import streamlit as st
import zipfile
import io
import os
import json
import base64
import tempfile
import anthropic
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

MAX_FILE_SIZE_MB = 500
MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024

st.set_page_config(page_title="Invoice PDF Extractor", page_icon="🧾", layout="wide")

st.markdown("""
<style>
    .main-title { font-size: 2rem; font-weight: 700; color: #1F3864; }
    .sub-title { font-size: 1rem; color: #555; margin-bottom: 1.5rem; }
    .status-box { padding: 0.75rem 1rem; border-radius: 8px; margin: 0.3rem 0; }
    .stProgress > div > div { background-color: #2E75B6; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">🧾 Invoice PDF Extractor</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Upload a ZIP file containing project folders with PDF invoices. Max 500 MB per file.</div>', unsafe_allow_html=True)

# ─── Helpers ───────────────────────────────────────────────────────────────────

def extract_pdfs_from_zip(zip_bytes: bytes) -> dict[str, bytes]:
    """Recursively extract all PDFs from nested zips. Returns {path: pdf_bytes}"""
    pdfs = {}
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        for name in zf.namelist():
            if name.endswith('/'):
                continue
            data = zf.read(name)
            if name.lower().endswith('.pdf'):
                pdfs[name] = data
            elif name.lower().endswith('.zip'):
                # Nested zip
                try:
                    nested = extract_pdfs_from_zip(data)
                    for sub_name, sub_data in nested.items():
                        pdfs[f"{name}/{sub_name}"] = sub_data
                except Exception:
                    pass
    return pdfs


def pdf_to_base64(pdf_bytes: bytes) -> str:
    return base64.standard_b64encode(pdf_bytes).decode("utf-8")


def extract_invoice_data(pdf_bytes: bytes, filename: str, client: anthropic.Anthropic) -> dict:
    """Call Claude API to extract structured invoice data from a PDF."""
    prompt = """Extract all invoice details from this PDF and return ONLY a valid JSON object with these exact keys:
{
  "bill_date": "",
  "bill_number": "",
  "invoice_id": "",
  "document_date": "",
  "transaction_id": "",
  "payment_method": "",
  "currency_code": "",
  "vendor_name": "",
  "vendor_address": "",
  "vendor_gstin": "",
  "vendor_pan": "",
  "buyer_name": "",
  "buyer_address": "",
  "buyer_gstin": "",
  "buyer_pan": "",
  "gst_treatment": "",
  "source_of_supply": "",
  "destination_of_supply": "",
  "place_of_supply": "",
  "reverse_charge": "",
  "is_inclusive_tax": "",
  "hsn_sac": "",
  "line_items": [
    {
      "line_no": "",
      "campaign_name": "",
      "item_name": "",
      "item_type": "",
      "impressions": "",
      "rate": "",
      "tax_name": "",
      "tax_percentage": "",
      "tax_amount": "",
      "item_total": ""
    }
  ],
  "subtotal": "",
  "igst_amount": "",
  "total_amount": "",
  "vendor_notes": ""
}
Return ONLY the JSON. No markdown, no explanation."""

    b64 = pdf_to_base64(pdf_bytes)
    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=2000,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "document",
                    "source": {"type": "base64", "media_type": "application/pdf", "data": b64}
                },
                {"type": "text", "text": prompt}
            ]
        }]
    )
    raw = response.content[0].text.strip()
    raw = raw.replace("```json", "").replace("```", "").strip()
    return json.loads(raw)


# ─── Excel Builder ─────────────────────────────────────────────────────────────

def build_excel(all_data: list[dict]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoices"

    # Styles
    hf = PatternFill("solid", start_color="1F3864")
    lf = PatternFill("solid", start_color="D9E1F2")
    gf = PatternFill("solid", start_color="E2EFDA")
    h_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    d_font = Font(name="Arial", size=10)
    b_font = Font(name="Arial", bold=True, size=10)
    thin = Side(style="thin", color="B0B0B0")
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)
    ctr = Alignment(horizontal="center", vertical="center", wrap_text=True)
    lft = Alignment(horizontal="left", vertical="center", wrap_text=True)

    headers = [
        "Bill Date", "Bill Number", "Invoice ID", "Document Date", "Transaction ID",
        "Payment Method", "Currency", "Vendor Name", "Vendor GSTIN", "Vendor PAN",
        "Buyer Name", "Buyer GSTIN", "Buyer PAN", "GST Treatment",
        "Source of Supply", "Destination of Supply", "Place of Supply",
        "Reverse Charge", "Line No.", "Campaign Name", "Item Name",
        "HSN/SAC", "Item Type", "Impressions", "Rate (INR)",
        "Tax Name", "Tax %", "Tax Amount (INR)", "Item Total (INR)",
        "SubTotal (INR)", "IGST Amount (INR)", "Grand Total (INR)"
    ]

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = h_font
        cell.fill = hf
        cell.alignment = ctr
        cell.border = bdr

    ws.row_dimensions[1].height = 28

    row = 2
    for inv in all_data:
        items = inv.get("line_items", [{}])
        if not items:
            items = [{}]
        for idx, item in enumerate(items):
            fill = lf if row % 2 == 0 else PatternFill("solid", start_color="FFFFFF")
            vals = [
                inv.get("bill_date", ""),
                inv.get("bill_number", ""),
                inv.get("invoice_id", ""),
                inv.get("document_date", ""),
                inv.get("transaction_id", ""),
                inv.get("payment_method", ""),
                inv.get("currency_code", ""),
                inv.get("vendor_name", ""),
                inv.get("vendor_gstin", ""),
                inv.get("vendor_pan", ""),
                inv.get("buyer_name", ""),
                inv.get("buyer_gstin", ""),
                inv.get("buyer_pan", ""),
                inv.get("gst_treatment", ""),
                inv.get("source_of_supply", ""),
                inv.get("destination_of_supply", ""),
                inv.get("place_of_supply", ""),
                inv.get("reverse_charge", ""),
                item.get("line_no", ""),
                item.get("campaign_name", ""),
                item.get("item_name", ""),
                inv.get("hsn_sac", ""),
                item.get("item_type", ""),
                item.get("impressions", ""),
                item.get("rate", ""),
                item.get("tax_name", ""),
                item.get("tax_percentage", ""),
                item.get("tax_amount", ""),
                item.get("item_total", ""),
                inv.get("subtotal", "") if idx == 0 else "",
                inv.get("igst_amount", "") if idx == 0 else "",
                inv.get("total_amount", "") if idx == 0 else "",
            ]
            for col, val in enumerate(vals, 1):
                c = ws.cell(row=row, column=col, value=val)
                c.font = d_font
                c.fill = fill
                c.alignment = ctr if col > 7 else lft
                c.border = bdr
            ws.row_dimensions[row].height = 20
            row += 1

    # Column widths
    widths = [12, 22, 22, 16, 36, 16, 10, 32, 20, 14, 36, 20, 14,
              14, 18, 20, 18, 14, 10, 28, 28, 12, 12, 12, 14, 12, 8, 16, 16, 16, 16, 16]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─── UI ────────────────────────────────────────────────────────────────────────

api_key = st.sidebar.text_input("🔑 Anthropic API Key", type="password",
                                 help="Required to extract invoice data using Claude AI")

st.sidebar.markdown("---")
st.sidebar.markdown("**Instructions:**")
st.sidebar.markdown("""
1. Enter your Anthropic API key  
2. Upload a ZIP file (max 500 MB)  
3. The ZIP may contain:  
   - Project folders  
   - Nested ZIPs  
   - PDF invoices at any depth  
4. Click **Extract & Download Excel**
""")

uploaded = st.file_uploader(
    "Upload ZIP file (max 500 MB)",
    type=["zip"],
    help="ZIP containing project subfolders with PDF invoices"
)

if uploaded:
    file_size = uploaded.size
    size_mb = file_size / (1024 * 1024)

    if file_size > MAX_FILE_SIZE_BYTES:
        st.error(f"❌ File too large: {size_mb:.1f} MB. Maximum allowed is {MAX_FILE_SIZE_MB} MB.")
        st.stop()

    st.info(f"📦 File: **{uploaded.name}** | Size: **{size_mb:.2f} MB**")

    if st.button("🚀 Extract & Download Excel", type="primary"):
        if not api_key:
            st.error("⚠️ Please enter your Anthropic API key in the sidebar.")
            st.stop()

        client = anthropic.Anthropic(api_key=api_key)

        with st.spinner("Reading ZIP file..."):
            zip_bytes = uploaded.read()
            try:
                pdfs = extract_pdfs_from_zip(zip_bytes)
            except Exception as e:
                st.error(f"❌ Failed to read ZIP: {e}")
                st.stop()

        if not pdfs:
            st.warning("⚠️ No PDF files found in the uploaded ZIP.")
            st.stop()

        st.success(f"✅ Found **{len(pdfs)} PDF(s)** across all folders.")

        with st.expander("📄 PDF files found", expanded=False):
            for p in sorted(pdfs.keys()):
                st.text(f"  • {p}")

        all_invoice_data = []
        errors = []

        progress = st.progress(0)
        status_area = st.empty()

        for i, (path, pdf_bytes) in enumerate(sorted(pdfs.items())):
            fname = os.path.basename(path)
            status_area.markdown(f"⏳ Processing **{fname}** ({i+1}/{len(pdfs)})...")
            try:
                data = extract_invoice_data(pdf_bytes, fname, client)
                data["_source_file"] = path
                all_invoice_data.append(data)
                status_area.markdown(f"✅ **{fname}** extracted successfully")
            except Exception as e:
                errors.append((path, str(e)))
                status_area.markdown(f"❌ **{fname}** failed: {e}")

            progress.progress((i + 1) / len(pdfs))

        status_area.empty()
        progress.empty()

        if errors:
            st.warning(f"⚠️ {len(errors)} file(s) failed to process:")
            for path, err in errors:
                st.text(f"  • {path}: {err}")

        if all_invoice_data:
            st.success(f"✅ Successfully extracted **{len(all_invoice_data)}** invoice(s)!")

            with st.spinner("Building Excel file..."):
                excel_bytes = build_excel(all_invoice_data)

            st.download_button(
                label="📥 Download Excel",
                data=excel_bytes,
                file_name="extracted_invoices.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

            # Preview
            st.markdown("### 📊 Preview (first 5 invoices)")
            import pandas as pd
            preview_rows = []
            for inv in all_invoice_data[:5]:
                preview_rows.append({
                    "Invoice ID": inv.get("invoice_id", ""),
                    "Date": inv.get("bill_date", ""),
                    "Vendor": inv.get("vendor_name", ""),
                    "Buyer": inv.get("buyer_name", ""),
                    "Total (INR)": inv.get("total_amount", ""),
                    "Source File": inv.get("_source_file", ""),
                })
            st.dataframe(pd.DataFrame(preview_rows), use_container_width=True)
        else:
            st.error("❌ No invoices could be extracted. Please check your API key and PDF files.")
