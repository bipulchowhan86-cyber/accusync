from fastapi import APIRouter, UploadFile, File, Form, HTTPException
from fastapi.responses import Response, StreamingResponse
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import re
import os
import shutil

router = APIRouter()

# Global memory cache for current session export
CURRENT_STATEMENT_DATA = {
    "metadata": {},
    "transactions": [],
    "totals": {}
}

def clean_amount(val):
    if val is None:
        return 0.0
    s = str(val).strip()
    if not s or s in ['-', '--', 'None', 'NA', 'N/A']:
        return 0.0
    # Strip symbols
    s = re.sub(r'[₹\s,`"\'RsINRrsinr]', '', s)
    s = re.sub(r'[^\d.-]', '', s)
    try:
        n = float(s)
        return abs(n)
    except:
        return 0.0

def detect_bank_metadata(text_blob):
    """Auto-detect Bank Name, Account Holder, Account Number, IFSC from header text"""
    meta = {
        "bank_name": "Standard Indian Bank",
        "account_holder": "Not Detected",
        "account_no": "Not Detected",
        "ifsc": "Not Detected",
        "branch": "Main Branch"
    }

    # Bank detection
    text_lower = text_blob.lower()
    if "state bank of india" in text_lower or "sbi" in text_lower:
        meta["bank_name"] = "State Bank of India (SBI)"
    elif "hdfc" in text_lower:
        meta["bank_name"] = "HDFC Bank Ltd"
    elif "icici" in text_lower:
        meta["bank_name"] = "ICICI Bank Ltd"
    elif "axis" in text_lower:
        meta["bank_name"] = "Axis Bank Ltd"
    elif "punjab national" in text_lower or "pnb" in text_lower:
        meta["bank_name"] = "Punjab National Bank"
    elif "canara" in text_lower:
        meta["bank_name"] = "Canara Bank"
    elif "bank of baroda" in text_lower or "bob" in text_lower:
        meta["bank_name"] = "Bank of Baroda"

    # IFSC detection (Standard 11-char regex)
    ifsc_match = re.search(r'\b[A-Z]{4}0[A-Z0-9]{6}\b', text_blob)
    if ifsc_match:
        meta["ifsc"] = ifsc_match.group(0)

    # Account Number detection
    acc_match = re.search(r'(?:account\s*no|a/c\s*no|ac\s*no|account\s*number)[:.\s]*([0-9]{9,18})', text_blob, re.IGNORECASE)
    if acc_match:
        meta["account_no"] = acc_match.group(1)
    else:
        # Fallback 11 to 16 digit series
        digits_match = re.search(r'\b[0-9]{11,18}\b', text_blob)
        if digits_match:
            meta["account_no"] = digits_match.group(0)

    # Name detection
    name_match = re.search(r'(?:mr\.|ms\.|shri|name\s*:|customer\s*name)[:\s]*([A-Z\s]{4,35})', text_blob, re.IGNORECASE)
    if name_match:
        meta["account_holder"] = name_match.group(1).strip()

    return meta

@router.post("/parse-bank-statement")
async def parse_bank_statement(
    file: UploadFile = File(...),
    password: str = Form(None)
):
    global CURRENT_STATEMENT_DATA

    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Kripya PDF format statement upload karein.")

    os.makedirs("uploads", exist_ok=True)
    file_path = os.path.join("uploads", file.filename)

    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    all_raw_tables = []
    header_text_blob = ""

    try:
        with pdfplumber.open(file_path, password=password if password else None) as pdf:
            # Extract header text from first 2 pages for metadata
            for p_idx in range(min(2, len(pdf.pages))):
                header_text_blob += (pdf.pages[p_idx].extract_text() or "") + "\n"

            for page in pdf.pages:
                tables = page.extract_tables({
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "snap_tolerance": 3,
                })
                if not tables:
                    tables = page.extract_tables() # fallback text-based

                for table in tables:
                    for row in table:
                        clean_row = [str(col).strip() if col is not None else "" for col in row]
                        if any(clean_row):
                            all_raw_tables.append(clean_row)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PDF reading error (Password protected ho sakti hai): {str(e)}")

    if not all_raw_tables:
        raise HTTPException(status_code=400, detail="PDF se koi transaction table extract nahi hui.")

    # 1. Bank Metadata
    meta = detect_bank_metadata(header_text_blob)

    # 2. Find Table Header Row
    header_row_idx = 0
    for idx, r in enumerate(all_raw_tables[:10]):
        row_str = " ".join(r).lower()
        if any(w in row_str for w in ["date", "particulars", "description", "narration", "debit", "withdrawal"]):
            header_row_idx = idx
            break

    raw_headers = all_raw_tables[header_row_idx]
    data_rows = all_raw_tables[header_row_idx + 1:]

    # Map column indexes
    date_idx, desc_idx, chq_idx, debit_idx, credit_idx, bal_idx = -1, -1, -1, -1, -1, -1

    for i, h in enumerate(raw_headers):
        t = str(h).lower()
        if "date" in t and date_idx == -1:
            date_idx = i
        elif any(k in t for k in ["particular", "description", "narration", "details"]) and desc_idx == -1:
            desc_idx = i
        elif any(k in t for k in ["chq", "cheque", "ref", "utr", "tran id"]) and chq_idx == -1:
            chq_idx = i
        elif any(k in t for k in ["debit", "withdrawal", "dr"]) and debit_idx == -1:
            debit_idx = i
        elif any(k in t for k in ["credit", "deposit", "cr"]) and credit_idx == -1:
            credit_idx = i
        elif any(k in t for k in ["balance", "bal"]) and bal_idx == -1:
            bal_idx = i

    # Fallback indexes if not detected
    if desc_idx == -1 and len(raw_headers) > 2: desc_idx = 2
    if debit_idx == -1 and len(raw_headers) > 3: debit_idx = 3
    if credit_idx == -1 and len(raw_headers) > 4: credit_idx = 4
    if bal_idx == -1 and len(raw_headers) > 5: bal_idx = 5

    # 3. Standardize Transactions into 6 core columns
    standard_rows = []
    tot_debit = 0.0
    tot_credit = 0.0

    for r in data_rows:
        # Must have at least date or description
        r_date = r[date_idx] if 0 <= date_idx < len(r) else ""
        r_desc = r[desc_idx] if 0 <= desc_idx < len(r) else ""
        r_chq  = r[chq_idx] if 0 <= chq_idx < len(r) else ""
        r_deb  = clean_amount(r[debit_idx]) if 0 <= debit_idx < len(r) else 0.0
        r_crd  = clean_amount(r[credit_idx]) if 0 <= credit_idx < len(r) else 0.0
        r_bal  = clean_amount(r[bal_idx]) if 0 <= bal_idx < len(r) else 0.0

        # Skip headers repeating across pages
        if "date" in str(r_date).lower() or "particular" in str(r_desc).lower():
            continue

        # Valid row if has description and (amount or balance)
        if r_desc or r_deb > 0 or r_crd > 0:
            tot_debit += r_deb
            tot_credit += r_crd

            # Default Auto Ledger mapping
            ledger = "Suspense A/c"
            d_lower = r_desc.lower()
            if any(k in d_lower for k in ["charge", "fee", "gst", "sms"]):
                ledger = "Bank Charges"
            elif any(k in d_lower for k in ["fuel", "petrol", "hpcl", "bpcl", "iocl"]):
                ledger = "Fuel & Conveyance"
            elif any(k in d_lower for k in ["salary", "payroll"]):
                ledger = "Salary Expense"
            elif any(k in d_lower for k in ["swiggy", "zomato", "restaurant", "hotel"]):
                ledger = "Staff Welfare / Refreshment"
            elif r_crd > 0:
                ledger = "Sundry Debtors / Revenue"
            elif r_deb > 0:
                ledger = "Sundry Creditors / Vendor"

            standard_rows.append({
                "date": r_date,
                "narration": r_desc,
                "ref_no": r_chq,
                "debit": r_deb,
                "credit": r_crd,
                "balance": r_bal,
                "ledger": ledger
            })

    # Cache for exports
    CURRENT_STATEMENT_DATA["metadata"] = meta
    CURRENT_STATEMENT_DATA["transactions"] = standard_rows
    CURRENT_STATEMENT_DATA["totals"] = {
        "debit": round(tot_debit, 2),
        "credit": round(tot_credit, 2),
        "count": len(standard_rows)
    }

    return {
        "status": "success",
        "filename": file.filename,
        "metadata": meta,
        "total_rows_found": len(standard_rows),
        "total_debit": round(tot_debit, 2),
        "total_credit": round(tot_credit, 2),
        "transactions": standard_rows
    }

@router.get("/export/excel")
def export_excel():
    """Download professionally formatted Excel with total formulas and metadata banner"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bank Statement"
    ws.views.sheetView[0].showGridLines = True

    meta = CURRENT_STATEMENT_DATA["metadata"]
    txns = CURRENT_STATEMENT_DATA["transactions"]
    totals = CURRENT_STATEMENT_DATA["totals"]

    # Brand Title
    ws.merge_cells("A1:G1")
    ws["A1"] = f"{meta.get('bank_name', 'Bank Statement')} - Account Statement"
    ws["A1"].font = Font(name="Calibri", size=16, bold=True, color="1E3A8A")

    # Meta Banner
    ws["A3"] = "Account Holder:"
    ws["B3"] = meta.get("account_holder", "BIPUL CHAUHAN")
    ws["A4"] = "Account Number:"
    ws["B4"] = meta.get("account_no", "43822500216")
    ws["D3"] = "IFSC Code:"
    ws["E3"] = meta.get("ifsc", "SBIN0006878")
    ws["D4"] = "Bank Branch:"
    ws["E4"] = meta.get("branch", "WEST GUWAHATI")

    for cell in ["A3", "A4", "D3", "D4"]:
        ws[cell].font = Font(bold=True, color="475569")

    # Headers
    headers = ["Date", "Narration / Description", "Cheque / Ref No", "Debit (Withdrawal)", "Credit (Deposit)", "Balance", "Tally Ledger"]
    ws.append([]) # Row 5 blank
    ws.append(headers) # Row 6

    header_fill = PatternFill(start_color="1E40AF", end_color="1E40AF", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    for col_num in range(1, 8):
        cell = ws.cell(row=6, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    start_row = 7
    for row_idx, t in enumerate(txns, start=start_row):
        ws.cell(row=row_idx, column=1, value=t["date"]).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=2, value=t["narration"])
        ws.cell(row=row_idx, column=3, value=t["ref_no"]).alignment = Alignment(horizontal="center")
        
        c_deb = ws.cell(row=row_idx, column=4, value=t["debit"] if t["debit"] > 0 else "")
        c_deb.number_format = '₹#,##0.00'
        c_deb.font = Font(color="DC2626") if t["debit"] > 0 else Font()

        c_crd = ws.cell(row=row_idx, column=5, value=t["credit"] if t["credit"] > 0 else "")
        c_crd.number_format = '₹#,##0.00'
        c_crd.font = Font(color="16A34A") if t["credit"] > 0 else Font()

        c_bal = ws.cell(row=row_idx, column=6, value=t["balance"] if t["balance"] > 0 else "")
        c_bal.number_format = '₹#,##0.00'

        ws.cell(row=row_idx, column=7, value=t["ledger"])

        for col_num in range(1, 8):
            ws.cell(row=row_idx, column=col_num).border = thin_border

    # Grand Totals Row
    last_row = start_row + len(txns)
    ws.cell(row=last_row, column=2, value="TOTAL AMOUNT").font = Font(bold=True, size=12)
    ws.cell(row=last_row, column=2).alignment = Alignment(horizontal="right")
    
    tot_deb_cell = ws.cell(row=last_row, column=4, value=f"=SUM(D7:D{last_row-1})")
    tot_deb_cell.number_format = '₹#,##0.00'
    tot_deb_cell.font = Font(bold=True, color="DC2626", size=12)

    tot_crd_cell = ws.cell(row=last_row, column=5, value=f"=SUM(E7:E{last_row-1})")
    tot_crd_cell.number_format = '₹#,##0.00'
    tot_crd_cell.font = Font(bold=True, color="16A34A", size=12)

    for col_num in range(1, 8):
        c = ws.cell(row=last_row, column=col_num)
        c.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        c.border = thin_border

    # Auto-fit columns
    column_widths = {'A': 14, 'B': 45, 'C': 20, 'D': 18, 'E': 18, 'F': 18, 'G': 24}
    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"{meta.get('bank_name', 'Bank')}_Statement_{meta.get('account_no', 'Export')}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@router.get("/export/tally-xml")
def export_tally_xml():
    """Generates official Tally Prime Vouchers XML (Payment & Receipt)"""
    txns = CURRENT_STATEMENT_DATA["transactions"]
    meta = CURRENT_STATEMENT_DATA["metadata"]
    bank_ledger = meta.get("bank_name", "Bank Account")

    xml_lines = [
        '<ENVELOPE>',
        '  <HEADER>',
        '    <TALLYREQUEST>Import Data</TALLYREQUEST>',
        '  </HEADER>',
        '  <BODY>',
        '    <IMPORTDATA>',
        '      <REQUESTDESC>',
        '        <REPORTNAME>Vouchers</REPORTNAME>',
        '        <STATICVARIABLES>',
        '          <SVCURRENTCOMPANY>Company</SVCURRENTCOMPANY>',
        '        </STATICVARIABLES>',
        '      </REQUESTDESC>',
        '      <REQUESTDATA>'
    ]

    for t in txns:
        # Date conversion to YYYYMMDD for Tally
        clean_date = re.sub(r'[-/.]', '', str(t['date']))
        if len(clean_date) != 8:
            clean_date = "20260401" # Safe fallback
        else:
            # DDMMYYYY to YYYYMMDD
            clean_date = clean_date[4:] + clean_date[2:4] + clean_date[0:2]

        vch_type = "Payment" if t["debit"] > 0 else "Receipt"
        amt = t["debit"] if t["debit"] > 0 else t["credit"]
        party_ledger = t["ledger"]

        xml_lines.extend([
            '        <TALLYMESSAGE xmlns:UDF="TallyUDF">',
            f'          <VOUCHER VCHTYPE="{vch_type}" ACTION="Create">',
            f'            <DATE>{clean_date}</DATE>',
            f'            <NARRATION>{t["narration"]}</NARRATION>',
            f'            <VOUCHERTYPENAME>{vch_type}</VOUCHERTYPENAME>',
            f'            <REFERENCE>{t["ref_no"]}</REFERENCE>',
            '            <ALLLEDGERENTRIES.LIST>',
            f'              <LEDGERNAME>{party_ledger}</LEDGERNAME>',
            f'              <ISDEEMEDPOSITIVE>{"Yes" if vch_type == "Payment" else "No"}</ISDEEMEDPOSITIVE>',
            f'              <AMOUNT>{-amt if vch_type == "Payment" else amt}</AMOUNT>',
            '            </ALLLEDGERENTRIES.LIST>',
            '            <ALLLEDGERENTRIES.LIST>',
            f'              <LEDGERNAME>{bank_ledger}</LEDGERNAME>',
            f'              <ISDEEMEDPOSITIVE>{"No" if vch_type == "Payment" else "Yes"}</ISDEEMEDPOSITIVE>',
            f'              <AMOUNT>{amt if vch_type == "Payment" else -amt}</AMOUNT>',
            '            </ALLLEDGERENTRIES.LIST>',
            '          </VOUCHER>',
            '        </TALLYMESSAGE>'
        ])

    xml_lines.extend([
        '      </REQUESTDATA>',
        '    </IMPORTDATA>',
        '  </BODY>',
        '</ENVELOPE>'
    ])

    xml_content = "\n".join(xml_lines)
    return Response(
        content=xml_content,
        media_type="application/xml",
        headers={"Content-Disposition": "attachment; filename=Tally_Prime_Vouchers.xml"}
    )