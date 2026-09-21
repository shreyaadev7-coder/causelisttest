import os
import sys
import re
import time
import traceback
from datetime import datetime, timedelta
import pytz
from playwright.sync_api import sync_playwright
from google import genai
from google.genai import types
from pydantic import BaseModel, Field
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.cell.rich_text import TextBlock, CellRichText
from openpyxl.cell.text import InlineFont

class CauseListRow(BaseModel):
    sl_no: str = Field(description="First SL NO column")
    case_number: str = Field(description="CASE NUMBER (e.g. WP NO 102709/2026)")
    case_name: str = Field(description="Full raw case name string")
    ch: str = Field(description="Court Hall number under CH")
    list_num: str = Field(description="List number under LIST")
    list_sl_no: str = Field(description="Item number under second SL NO")
    status: str = Field(description="STATUS column")
    judges: str = Field(description="JUDGES column")

class CauseListDocument(BaseModel):
    date: str = Field(description="Date displayed at top of cause list")
    rows: list[CauseListRow]

def get_target_date_ist():
    manual = os.environ.get("TEST_DATE")
    if manual and manual.strip():
        return manual.strip()
    return "22/09/2026"

def clean_case_details(raw_name: str):
    """
    Cleans raw case name strictly into '[Petitioner] vs [Respondent]'
    and determines which party Mahesh Chowdhary represents.
    """
    text = re.sub(r'\s+', ' ', raw_name).strip()
    
    # Check if respondent side representation is indicated (e.g., '(RESPONDENT NO. 5 TO 7)')
    is_res_rep = bool(re.search(r'\b(RESPONDENT|RESPODNENT|RES)\b.*?\b(NO|NOS|R\d+|\d+)\b', text, re.I))
    
    parts = re.split(r'\s+V/?S\.?\s+', text, maxsplit=1, flags=re.IGNORECASE)
    if len(parts) == 2:
        pet_part, res_part = parts[0].strip(), parts[1].strip()
    else:
        pet_part, res_part = text, ""

    if re.search(r'\b(MAHESH|CHOWDHA?R[YI])\b', res_part, re.I):
        in_charge = "RES"
    elif re.search(r'\b(MAHESH|CHOWDHA?R[YI])\b', pet_part, re.I):
        in_charge = "PET"
    elif is_res_rep:
        in_charge = "RES"
    else:
        in_charge = "PET"

    def filter_noise(side_text: str) -> str:
        cleaned = re.sub(r'^(PET|RES|PETITIONER|RESPONDENT|APPELLANT|COMPLAINANT)\s*:\s*', '', side_text, flags=re.I)
        cleaned = re.sub(r'\([^\)]*\)', '', cleaned).strip()
        chunks = [c.strip() for c in re.split(r'[,;]', cleaned) if c.strip()]
        valid_parties = []
        for c in chunks:
            if not re.search(r'\b(MAHESH|CHOWDHA?R[YI]|AGA|HCGP|ADV|ADVOCATE|ADVOCATES|COUNSEL|FOR\s+RES|FOR\s+PET|GOVT|PLEADER)\b', c, re.I):
                valid_parties.append(c)
        if valid_parties:
            res = ", ".join(valid_parties).strip()
            return re.sub(r'[\s&,]+$', '', res).strip()
        return chunks[0] if chunks else side_text.strip()

    pet_clean = filter_noise(pet_part)
    res_clean = filter_noise(res_part)
    return pet_clean, res_clean, in_charge

def extract_rows_from_page(page) -> list[list[str]]:
    extracted_rows = []
    tables = page.locator("table:visible").all()
    for table in tables:
        rows = table.locator("tr").all()
        for row in rows:
            cells = [td.inner_text().strip() for td in row.locator("th, td").all()]
            if len(cells) >= 7 and not ("CASE NUMBER" in cells[1].upper() if len(cells) > 1 else False):
                while len(cells) < 8:
                    cells.append("")
                extracted_rows.append(cells[:8])
    return extracted_rows

def fetch_cause_list_pdf(pdf_path="causelist.pdf"):
    date_str = get_target_date_ist()
    print(f"[*] Target Causelist Date: {date_str}")

    scraperapi_key = os.environ.get("SCRAPERAPI_KEY")
    if not scraperapi_key:
        raise ValueError("SCRAPERAPI_KEY environment variable is missing from GitHub Secrets.")

    launch_args = {
        "headless": True,
        "proxy": {
            "server": "http://proxy-server.scraperapi.com:8001",
            "username": "scraperapi.country_code=in",
            "password": scraperapi_key
        },
        "args": [
            "--no-sandbox",
            "--disable-setuid-sandbox",
            "--disable-dev-shm-usage",
            "--ignore-certificate-errors"
        ]
    }

    raw_table_rows = []

    with sync_playwright() as p:
        browser = p.chromium.launch(**launch_args)
        context = browser.new_context(
            ignore_https_errors=True,
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
            viewport={"width": 1366, "height": 900}
        )
        page = context.new_page()
        page.add_init_script("window.print = () => { console.log('window.print intercepted'); };")

        print("[*] Navigating to High Court Portal via Indian Gateway...")
        page.goto("https://judiciary.karnataka.gov.in/causelistSearch.php", wait_until="domcontentloaded", timeout=90000)
        page.wait_for_timeout(3000)

        # 1. Bench Selection
        bench_select = page.locator("select[name='bench']:visible").first
        for opt in bench_select.locator("option").all():
            opt_text = opt.inner_text().strip()
            if "Bengaluru" in opt_text or "Bangalore" in opt_text or "Principal" in opt_text:
                bench_select.select_option(label=opt_text)
                break
        page.wait_for_timeout(2000)

        # 2. Search By: Advocate
        searchby_select = page.locator("select[name='searchby']:visible").first
        for opt in searchby_select.locator("option").all():
            if "Advocate" in opt.inner_text().strip():
                searchby_select.select_option(label=opt.inner_text().strip())
                break
        page.wait_for_timeout(3000)

        # 3. Enter Dates
        page.locator("#fromDt:visible").first.fill(date_str)
        page.locator("#toDt:visible").first.fill(date_str)
        page.wait_for_timeout(1000)

        # 4. Enter Advocate Name
        adv_input = page.locator("input[placeholder='Enter Advocate Name']:visible").first
        if not adv_input.is_visible():
            adv_input = page.locator("#advName:visible").first

        adv_input.click()
        adv_input.fill("")
        adv_input.fill("Mahesh Chowdhary")
        page.wait_for_timeout(1000)

        # 5. Click GET DETAILS
        page.locator("#getData:visible").first.click()
        print("[*] Waiting 7 seconds for court database query to complete...")
        page.wait_for_timeout(7000)

        raw_table_rows = extract_rows_from_page(page)

        # 6. Locate Print Button
        print_btn = None
        for btn in page.locator("input[type='button']:visible, button:visible, a:visible").all():
            btn_id = (btn.get_attribute("id") or "").lower()
            val = (btn.get_attribute("value") or "").lower()
            txt = (btn.inner_text() or "").lower()
            if btn_id == "getdata":
                continue
            if "print" in val or "print" in txt or "print" in btn_id:
                print_btn = btn
                break

        if print_btn:
            try:
                with context.expect_page(timeout=8000) as popup_info:
                    print_btn.click()
                print_page = popup_info.value
                print_page.wait_for_load_state("domcontentloaded")
                print_page.wait_for_timeout(3000)
                print_page.pdf(path=pdf_path, format="A4", print_background=True)
                popup_rows = extract_rows_from_page(print_page)
                if len(popup_rows) > len(raw_table_rows):
                    raw_table_rows = popup_rows
            except Exception:
                page.wait_for_timeout(2000)
                page.pdf(path=pdf_path, format="A4", print_background=True)
        else:
            page.pdf(path=pdf_path, format="A4", print_background=True)

        browser.close()

    return raw_table_rows

def write_to_excel(rows: list, excel_path="cause_list.xlsx"):
    """Writes clean cause list data to Excel without colors."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cause List"
    ws.views.sheetView[0].showGridLines = True

    headers = ["SL NO", "CASE NUMBER", "CASE NAME", "CH", "LIST", "SL NO", "STATUS", "JUDGES"]
    ws.append(headers)
    ws.row_dimensions[1].height = 24

    # 1. Column Headers in Bold (No background colors)
    header_font = Font(name="Calibri", size=10, bold=True)
    thin_border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000")
    )

    for col_idx in range(1, 9):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    bold_client_font = InlineFont(rFont="Calibri", sz=10, b=True)
    reg_font = InlineFont(rFont="Calibri", sz=10, b=False)

    for idx, row in enumerate(rows, start=1):
        if isinstance(row, CauseListRow):
            case_num = row.case_number.strip()
            raw_case_name = row.case_name.strip()
            ch = row.ch.strip()
            list_num = row.list_num.strip()
            list_sl_no = row.list_sl_no.strip()
            status = row.status.strip()
            judges = row.judges.strip()
        else:
            case_num = row[1].strip() if len(row) > 1 else ""
            raw_case_name = row[2].strip() if len(row) > 2 else ""
            ch = row[3].strip() if len(row) > 3 else ""
            list_num = row[4].strip() if len(row) > 4 else ""
            list_sl_no = row[5].strip() if len(row) > 5 else ""
            status = row[6].strip() if len(row) > 6 else ""
            judges = row[7].strip() if len(row) > 7 else ""

        pet_clean, res_clean, in_charge = clean_case_details(raw_case_name)
        plain_case_text = f"{pet_clean}\nvs\n{res_clean}"

        ws.append([
            idx,
            case_num,
            plain_case_text,
            ch,
            list_num,
            list_sl_no,
            status,
            judges
        ])

        curr_row = idx + 1
        ws.row_dimensions[curr_row].height = 45

        # 3. Bold the client represented by Mahesh Chowdhary
        case_cell = ws.cell(row=curr_row, column=3)
        try:
            if in_charge == "PET":
                case_cell.value = CellRichText(
                    TextBlock(bold_client_font, f"{pet_clean}\n"),
                    TextBlock(reg_font, f"vs\n{res_clean}")
                )
            else:
                case_cell.value = CellRichText(
                    TextBlock(reg_font, f"{pet_clean}\nvs\n"),
                    TextBlock(bold_client_font, res_clean)
                )
        except Exception:
            case_cell.value = plain_case_text

    # Format Data Rows
    for r in range(2, ws.max_row + 1):
        for c in range(1, 9):
            cell = ws.cell(row=r, column=c)
            cell.border = thin_border
            if c in [1, 4, 5, 6]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(name="Calibri", size=10, bold=False)
            elif c == 2:
                # 2. Case number in bold
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(name="Calibri", size=10, bold=True)
            elif c == 3:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                cell.font = Font(name="Calibri", size=10, bold=False)

    col_widths = {1: 8, 2: 22, 3: 45, 4: 8, 5: 8, 6: 10, 7: 25, 8: 32}
    for col_idx, width in col_widths.items():
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = width

    wb.save(excel_path)
    print(f"[+] Excel successfully generated at: {excel_path}")

def generate_pdf_summary(rows: list, date_str: str, pdf_path="cause_list_summary.pdf"):
    """Generates clean black-and-white landscape A4 PDF."""
    print("[*] Generating formatted PDF cause list...")

    rows_html = ""
    for idx, row in enumerate(rows, start=1):
        if isinstance(row, CauseListRow):
            case_num = row.case_number.strip()
            raw_case_name = row.case_name.strip()
            ch = row.ch.strip()
            list_num = row.list_num.strip()
            list_sl_no = row.list_sl_no.strip()
            status = row.status.strip()
            judges = row.judges.strip()
        else:
            case_num = row[1].strip() if len(row) > 1 else ""
            raw_case_name = row[2].strip() if len(row) > 2 else ""
            ch = row[3].strip() if len(row) > 3 else ""
            list_num = row[4].strip() if len(row) > 4 else ""
            list_sl_no = row[5].strip() if len(row) > 5 else ""
            status = row[6].strip() if len(row) > 6 else ""
            judges = row[7].strip() if len(row) > 7 else ""

        pet, res, in_charge = clean_case_details(raw_case_name)

        if in_charge == "PET":
            case_name_cell = f"<strong>{pet}</strong><br>vs<br>{res}"
        else:
            case_name_cell = f"{pet}<br>vs<br><strong>{res}</strong>"

        rows_html += f"""
        <tr>
            <td class="text-center">{idx}</td>
            <td class="text-center font-bold">{case_num}</td>
            <td>{case_name_cell}</td>
            <td class="text-center">{ch}</td>
            <td class="text-center">{list_num}</td>
            <td class="text-center">{list_sl_no}</td>
            <td>{status}</td>
            <td>{judges}</td>
        </tr>
        """

    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            @page {{
                size: A4 landscape;
                margin: 10mm;
            }}
            body {{
                font-family: Arial, Helvetica, sans-serif;
                color: #000000;
                margin: 0;
                padding: 0;
                font-size: 11px;
            }}
            .header {{
                display: flex;
                justify-content: space-between;
                align-items: center;
                border-bottom: 1px solid #000000;
                padding-bottom: 4px;
                margin-bottom: 8px;
            }}
            .header h1 {{
                margin: 0;
                font-size: 14px;
                text-transform: uppercase;
                color: #000000;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                table-layout: fixed;
            }}
            th {{
                background-color: #ffffff;
                color: #000000;
                font-size: 10px;
                font-weight: bold;
                padding: 6px 4px;
                border: 1px solid #000000;
            }}
            td {{
                padding: 6px 4px;
                border: 1px solid #000000;
                vertical-align: middle;
                word-wrap: break-word;
            }}
            .text-center {{ text-align: center; }}
            .font-bold {{ font-weight: bold; }}
        </style>
    </head>
    <body>
        <div class="header">
            <div>
                <h1>High Court of Karnataka &mdash; Bengaluru Bench</h1>
                <div>Daily Cause List &bull; Date: {date_str}</div>
            </div>
            <div style="text-align: right;">
                <div style="font-weight: bold;">ADV. MAHESH CHOWDHARY</div>
                <div>Total Matters: {len(rows)}</div>
            </div>
        </div>

        <table>
            <colgroup>
                <col style="width: 5%;">
                <col style="width: 15%;">
                <col style="width: 34%;">
                <col style="width: 5%;">
                <col style="width: 5%;">
                <col style="width: 6%;">
                <col style="width: 12%;">
                <col style="width: 18%;">
            </colgroup>
            <thead>
                <tr>
                    <th>SL NO</th>
                    <th>CASE NUMBER</th>
                    <th>CASE NAME</th>
                    <th>CH</th>
                    <th>LIST</th>
                    <th>SL NO</th>
                    <th>STATUS</th>
                    <th>JUDGES</th>
                </tr>
            </thead>
            <tbody>
                {rows_html}
            </tbody>
        </table>
    </body>
    </html>
    """

    with sync_playwright() as p:
        browser = p.chromium.launch(args=["--no-sandbox", "--disable-setuid-sandbox"])
        page = browser.new_page()
        page.set_content(html_content, wait_until="networkidle")
        page.pdf(
            path=pdf_path,
            format="A4",
            landscape=True,
            print_background=True,
            margin={"top": "10mm", "bottom": "10mm", "left": "10mm", "right": "10mm"}
        )
        browser.close()

    print(f"[+] Formatted PDF cause list generated at: {pdf_path}")

def parse_pdf_with_gemini(pdf_path="causelist.pdf", excel_path="cause_list.xlsx"):
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("GEMINI_API_KEY environment variable is missing.")

    client = genai.Client(api_key=api_key)
    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()

    prompt = (
        "Extract the complete cause list table from this PDF into the structured JSON schema. "
        "Strictly preserve exact cell contents, case numbers, party names, judge titles, and status text. "
        "Do not omit any row. If no cases are listed, return an empty rows array."
    )

    models_to_try = ["gemini-3.5-flash", "gemini-3.1-flash-lite", "gemini-3.6-flash"]
    response = None

    for model_name in models_to_try:
        try:
            print(f"[*] Submitting PDF to fallback AI model: {model_name}...")
            response = client.models.generate_content(
                model=model_name,
                contents=[types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"), prompt],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=CauseListDocument,
                    temperature=0.0
                )
            )
            break
        except Exception:
            time.sleep(3)

    if not response:
        raise RuntimeError("All Gemini models are temporarily unavailable.")

    parsed: CauseListDocument = CauseListDocument.model_validate_json(response.text)
    write_to_excel(parsed.rows, excel_path)
    return parsed.rows

if __name__ == "__main__":
    try:
        direct_rows = fetch_cause_list_pdf("causelist.pdf")
        target_date = get_target_date_ist()
        
        if direct_rows and len(direct_rows) > 0:
            write_to_excel(direct_rows, "cause_list.xlsx")
            generate_pdf_summary(direct_rows, target_date, "cause_list_summary.pdf")
        else:
            ai_rows = parse_pdf_with_gemini("causelist.pdf", "cause_list.xlsx")
            generate_pdf_summary(ai_rows, target_date, "cause_list_summary.pdf")
            
    except Exception as err:
        print(f"[FATAL] Process aborted: {err}", file=sys.stderr)
        traceback.print_exc()
        sys.exit(1)
