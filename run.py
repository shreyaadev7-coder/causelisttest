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
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.cell.rich_text import TextBlock, CellRichText
from openpyxl.cell.text import InlineFont

# Schema matching the exact 8 High Court PDF columns
class CauseListRow(BaseModel):
    sl_no: str = Field(description="First SL NO column (overall serial)")
    case_number: str = Field(description="CASE NUMBER (e.g. WP NO 102709/2026)")
    case_name: str = Field(description="Full CASE NAME with parties and respondent notes")
    ch: str = Field(description="Court Hall number under CH")
    list_num: str = Field(description="List number under LIST")
    list_sl_no: str = Field(description="Second SL NO column (Item number)")
    status: str = Field(description="STATUS column (e.g. ORDERS, PRELIMINARY HEARING)")
    judges: str = Field(description="JUDGES column with full bench text")

class CauseListDocument(BaseModel):
    date: str = Field(description="Date displayed at top of cause list")
    rows: list[CauseListRow]

def get_target_date_ist():
    manual = os.environ.get("TEST_DATE")
    if manual and manual.strip():
        print(f"[*] Overriding date with TEST_DATE: {manual.strip()}")
        return manual.strip()

    # Hardcoded test date for verification
    return "22/09/2026"

    # Production logic (uncomment when testing is complete):
    # ist = pytz.timezone('Asia/Kolkata')
    # next_day = datetime.now(ist) + timedelta(days=1)
    # return next_day.strftime("%d/%m/%Y")

def clean_case_details(raw_name: str):
    """
    Requirement 3:
    Cleans raw case name into '[Petitioner] vs [Respondent]'
    and determines which party Mahesh Chowdhary represents.
    """
    text = re.sub(r'\s+', ' ', raw_name).strip()
    
    # Check if respondent side representation is indicated (e.g. '(RESPONDENT NO. 5 TO 7)')
    is_res_rep = bool(re.search(r'\b(RESPONDENT|RESPODNENT|RES)\b.*?\b(NO|NOS|R\d+|\d+)\b', text, re.I))
    
    # Split across petitioner and respondent
    parts = re.split(r'\s+V/?S\.?\s+', text, maxsplit=1, flags=re.IGNORECASE)
    if len(parts) == 2:
        pet_part, res_part = parts[0].strip(), parts[1].strip()
    else:
        pet_part, res_part = text, ""

    # Detect which side Mahesh Chowdhary is in charge of
    if re.search(r'\b(MAHESH|CHOWDHA?R[YI])\b', res_part, re.I):
        in_charge = "RES"
    elif re.search(r'\b(MAHESH|CHOWDHA?R[YI])\b', pet_part, re.I):
        in_charge = "PET"
    elif is_res_rep:
        in_charge = "RES"
    else:
        in_charge = "PET"

    def filter_noise(side_text: str) -> str:
        # Strip PET:, RES:, etc.
        cleaned = re.sub(r'^(PET|RES|PETITIONER|RESPONDENT|APPELLANT|COMPLAINANT)\s*:\s*', '', side_text, flags=re.I)
        # Strip parenthetical notes like (MA NOT FILED), (RESPONDENT NO. 5 TO 7)
        cleaned = re.sub(r'\([^\)]*\)', '', cleaned).strip()
        
        # Split by comma to remove advocate/counsel names
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
    """Directly extracts the rendered table rows from the browser DOM in 0.1s."""
    print("[*] Extracting cause list table directly from page DOM...")
    extracted_rows = []
    
    tables = page.locator("table:visible").all()
    for table in tables:
        rows = table.locator("tr").all()
        for row in rows:
            cells = [td.inner_text().strip() for td in row.locator("th, td").all()]
            # Filter out empty rows or pure header duplicates
            if len(cells) >= 7 and not ("CASE NUMBER" in cells[1].upper() if len(cells) > 1 else False):
                while len(cells) < 8:
                    cells.append("")
                extracted_rows.append(cells[:8])
                
    print(f"[+] Direct DOM extraction found {len(extracted_rows)} case rows.")
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
        page.screenshot(path="00_initial_page.png")

        # 1. Bench Selection
        print("[*] Step 1: Selecting Bench -> Bengaluru Bench")
        bench_select = page.locator("select[name='bench']:visible").first
        selected_bench = False
        for opt in bench_select.locator("option").all():
            opt_text = opt.inner_text().strip()
            if "Bengaluru" in opt_text or "Bangalore" in opt_text or "Principal" in opt_text:
                bench_select.select_option(label=opt_text)
                print(f"[+] Bench selected: '{opt_text}'")
                selected_bench = True
                break
        if not selected_bench:
            bench_select.select_option(index=1)
        page.wait_for_timeout(2000)

        # 2. Search By: Advocate
        print("[*] Step 2: Selecting Search By -> Advocate")
        searchby_select = page.locator("select[name='searchby']:visible").first
        selected_search_by = False
        for opt in searchby_select.locator("option").all():
            opt_text = opt.inner_text().strip()
            if "Advocate" in opt_text:
                searchby_select.select_option(label=opt_text)
                print(f"[+] Search By selected: '{opt_text}'")
                selected_search_by = True
                break
        if not selected_search_by:
            searchby_select.select_option(index=1)
        page.wait_for_timeout(3000)

        # 3. Enter Dates
        print(f"[*] Step 3: Entering Date -> {date_str}")
        from_dt = page.locator("#fromDt:visible").first
        if from_dt.is_visible():
            from_dt.fill(date_str)
            print(f"[+] #fromDt set to: '{date_str}'")

        to_dt = page.locator("#toDt:visible").first
        if to_dt.is_visible():
            to_dt.fill(date_str)
            print(f"[+] #toDt set to: '{date_str}'")

        page.wait_for_timeout(1000)

        # 4. Enter Advocate Name
        print("[*] Step 4: Entering Advocate Name -> Mahesh Chowdhary")
        adv_input = page.locator("input[placeholder='Enter Advocate Name']:visible").first
        if not adv_input.is_visible():
            adv_input = page.locator("#advName:visible").first

        adv_input.click()
        adv_input.fill("")
        adv_input.fill("Mahesh Chowdhary")
        print(f"[+] Advocate Name verified in field: '{adv_input.input_value()}'")

        page.wait_for_timeout(1000)
        page.screenshot(path="01_form_filled.png")

        # 5. Click GET DETAILS
        print("[*] Step 5: Submitting search via '#getData'...")
        get_data_btn = page.locator("#getData:visible").first
        get_data_btn.click()

        print("[*] Waiting 7 seconds for court database query to complete...")
        page.wait_for_timeout(7000)
        page.screenshot(path="02_search_results.png")

        # Extract table directly from the live DOM (Primary Fast Path)
        raw_table_rows = extract_rows_from_page(page)

        # 6. Locate Print Button to save the PDF artifact
        print("[*] Step 6: Triggering Print List to capture causelist.pdf...")
        print_btn = None
        for btn in page.locator("input[type='button']:visible, button:visible, a:visible").all():
            btn_id = (btn.get_attribute("id") or "").lower()
            val = (btn.get_attribute("value") or "").lower()
            txt = (btn.inner_text() or "").lower()
            if btn_id == "getdata":
                continue
            if "print" in val or "print" in txt or "print" in btn_id:
                print_btn = btn
                print(f"[+] Identified Print Button: id='{btn_id}', val='{val}'")
                break

        if print_btn:
            try:
                with context.expect_page(timeout=8000) as popup_info:
                    print_btn.click()
                print_page = popup_info.value
                print_page.wait_for_load_state("domcontentloaded")
                print_page.wait_for_timeout(3000)
                print_page.pdf(path=pdf_path, format="A4", print_background=True)
                print_page.screenshot(path="03_print_view.png")
                # Also check print page for table records
                popup_rows = extract_rows_from_page(print_page)
                if len(popup_rows) > len(raw_table_rows):
                    raw_table_rows = popup_rows
                print(f"[+] Captured cause list PDF via print window: {pdf_path}")
            except Exception as e:
                print(f"[*] Print opened in-page or no popup ({e}). Rendering page to PDF...")
                page.wait_for_timeout(2000)
                page.pdf(path=pdf_path, format="A4", print_background=True)
                page.screenshot(path="03_print_view.png")
                print(f"[+] Saved cause list PDF from page: {pdf_path}")
        else:
            print("[*] Rendering table directly to PDF...")
            page.pdf(path=pdf_path, format="A4", print_background=True)
            page.screenshot(path="03_print_view.png")

        browser.close()

    return raw_table_rows

def write_to_excel(rows: list, excel_path="cause_list.xlsx"):
    """Writes standard 8-column cause list data to Excel with styling."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cause List"
    ws.views.sheetView[0].showGridLines = True

    headers = ["SL NO", "CASE NUMBER", "CASE NAME", "CH", "LIST", "SL NO", "STATUS", "JUDGES"]
    ws.append(headers)
    ws.row_dimensions[1].height = 26

    header_fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    regular_font = Font(name="Calibri", size=10)
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )

    for col_idx in range(1, 9):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    # Rich Text Fonts for Highlighting Mahesh Chowdhary's Client
    bold_client_font = InlineFont(rFont="Calibri", sz=10, b=True, color="001E40AF") # Bold Navy Blue
    reg_case_font = InlineFont(rFont="Calibri", sz=10, color="00334155")
    vs_font = InlineFont(rFont="Calibri", sz=9, i=True, color="0064748B")

    # Insert Rows
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

        # Requirement 3: Clean party names & identify in-charge
        pet_clean, res_clean, in_charge = clean_case_details(raw_case_name)
        plain_case_text = f"★ {pet_clean} (In-Charge)\nvs\n{res_clean}" if in_charge == "PET" else f"{pet_clean}\nvs\n★ {res_clean} (In-Charge)"

        # Requirement 2: Clean 1, 2, 3 sequential numbers
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
        # Requirement 1: Explicit row height to completely prevent cell collapse
        ws.row_dimensions[curr_row].height = 48

        # Apply rich text styling with highlight to the Case Name cell
        case_cell = ws.cell(row=curr_row, column=3)
        try:
            if in_charge == "PET":
                case_cell.value = CellRichText(
                    TextBlock(bold_client_font, f"★ {pet_clean} (In-Charge)\n"),
                    TextBlock(vs_font, "vs\n"),
                    TextBlock(reg_case_font, res_clean)
                )
            else:
                case_cell.value = CellRichText(
                    TextBlock(reg_case_font, f"{pet_clean}\n"),
                    TextBlock(vs_font, "vs\n"),
                    TextBlock(bold_client_font, f"★ {res_clean} (In-Charge)")
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
                cell.font = Font(name="Calibri", size=10, bold=(c in [1, 4, 6]))
            elif c == 2:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(name="Calibri", size=10, bold=True)
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                if c != 3:
                    cell.font = regular_font

    # Requirement 1: Generous column widths so text never overflows
    col_widths = {1: 8, 2: 24, 3: 48, 4: 8, 5: 8, 6: 10, 7: 28, 8: 36}
    for col_idx, width in col_widths.items():
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = width

    wb.save(excel_path)
    print(f"[+] Excel successfully generated at: {excel_path}")

def generate_pdf_summary(rows: list, date_str: str, pdf_path="cause_list_summary.pdf"):
    """Requirement 1: Generates a landscape A4 PDF cause list with visual party highlighting."""
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
            case_name_cell = f"""
                <div class="party-box client-box">
                    <span class="badge-tag">IN-CHARGE</span>
                    <strong>{pet}</strong>
                </div>
                <div class="vs-text">vs</div>
                <div class="party-box opp-box">{res}</div>
            """
        else:
            case_name_cell = f"""
                <div class="party-box opp-box">{pet}</div>
                <div class="vs-text">vs</div>
                <div class="party-box client-box">
                    <span class="badge-tag">IN-CHARGE</span>
                    <strong>{res}</strong>
                </div>
            """

        status_class = "status-ia" if "IA" in status.upper() else ("status-orders" if "ORDERS" in status.upper() else "status-general")

        rows_html += f"""
        <tr>
            <td class="text-center font-bold">{idx}</td>
            <td class="text-center case-num">{case_num}</td>
            <td>{case_name_cell}</td>
            <td class="text-center ch-cell">{ch}</td>
            <td class="text-center">{list_num}</td>
            <td class="text-center font-bold">{list_sl_no}</td>
            <td><span class="status-badge {status_class}">{status}</span></td>
            <td class="judges-text">{judges}</td>
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
                margin: 8mm;
            }}
            body {{
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
                color: #1e293b;
                margin: 0;
                padding: 0;
                font-size: 11px;
            }}
            .header {{
                display: flex;
                justify-content: space-between;
                align-items: center;
                border-bottom: 2px solid #0f172a;
                padding-bottom: 6px;
                margin-bottom: 8px;
            }}
            .header h1 {{
                margin: 0;
                font-size: 15px;
                text-transform: uppercase;
                color: #0f172a;
            }}
            .header-meta {{
                font-size: 11px;
                color: #475569;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                table-layout: fixed;
            }}
            th {{
                background-color: #1e293b;
                color: #ffffff;
                text-transform: uppercase;
                font-size: 10px;
                font-weight: 700;
                padding: 6px 4px;
                border: 1px solid #334155;
            }}
            td {{
                padding: 5px 4px;
                border: 1px solid #cbd5e1;
                vertical-align: middle;
                word-wrap: break-word;
            }}
            tr:nth-child(even) {{ background-color: #f8fafc; }}
            .text-center {{ text-align: center; }}
            .font-bold {{ font-weight: 700; }}
            .case-num {{ font-weight: 700; color: #0f172a; font-size: 11px; }}
            .ch-cell {{ font-size: 13px; font-weight: 800; color: #0284c7; }}
            .party-box {{ line-height: 1.3; }}
            .client-box {{
                color: #1e40af;
                background-color: #eff6ff;
                padding: 3px 6px;
                border-left: 3px solid #2563eb;
                border-radius: 2px;
            }}
            .badge-tag {{
                font-size: 8px;
                background: #2563eb;
                color: #fff;
                padding: 1px 4px;
                border-radius: 2px;
                font-weight: 800;
                margin-right: 4px;
            }}
            .opp-box {{ color: #475569; padding: 1px 6px; }}
            .vs-text {{ font-size: 9px; font-style: italic; color: #94a3b8; margin: 1px 0 1px 6px; }}
            .status-badge {{
                display: inline-block;
                padding: 2px 6px;
                border-radius: 3px;
                font-size: 9px;
                font-weight: 700;
                text-transform: uppercase;
            }}
            .status-ia {{ background: #fef3c7; color: #92400e; }}
            .status-orders {{ background: #fee2e2; color: #991b1b; }}
            .status-general {{ background: #e2e8f0; color: #334155; }}
            .judges-text {{ font-size: 10px; line-height: 1.25; color: #1e293b; }}
        </style>
    </head>
    <body>
        <div class="header">
            <div>
                <h1>High Court of Karnataka &mdash; Bengaluru Bench</h1>
                <div class="header-meta">Daily Cause List &bull; <strong>Date: {date_str}</strong></div>
            </div>
            <div style="text-align: right;">
                <div style="font-size: 12px; font-weight: 700; color: #0f172a;">ADV. MAHESH CHOWDHARY</div>
                <div class="header-meta">Total Matters: <strong>{len(rows)}</strong></div>
            </div>
        </div>

        <table>
            <colgroup>
                <col style="width: 5%;">
                <col style="width: 15%;">
                <col style="width: 32%;">
                <col style="width: 5%;">
                <col style="width: 5%;">
                <col style="width: 6%;">
                <col style="width: 14%;">
                <col style="width: 18%;">
            </colgroup>
            <thead>
                <tr>
                    <th>SL</th>
                    <th>Case Number</th>
                    <th>Case Name (Parties)</th>
                    <th>CH</th>
                    <th>List</th>
                    <th>Item</th>
                    <th>Status</th>
                    <th>Judges</th>
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
            margin={"top": "8mm", "bottom": "8mm", "left": "8mm", "right": "8mm"}
        )
        browser.close()

    print(f"[+] Formatted PDF cause list generated at: {pdf_path}")

def parse_pdf_with_gemini(pdf_path="causelist.pdf", excel_path="cause_list.xlsx"):
    """Multi-model fallback AI parser (used if direct DOM table was empty)."""
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
                contents=[
                    types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
                    prompt
                ],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=CauseListDocument,
                    temperature=0.0
                )
            )
            print(f"[+] AI extraction succeeded using {model_name}.")
            break
        except Exception as e:
            print(f"[!] {model_name} failed ({e}). Trying next model...")
            time.sleep(3)

    if not response:
        raise RuntimeError("All Gemini models are temporarily unavailable.")

    parsed: CauseListDocument = CauseListDocument.model_validate_json(response.text)
    print(f"[+] Extracted {len(parsed.rows)} rows via AI.")
    write_to_excel(parsed.rows, excel_path)
    return parsed.rows

if __name__ == "__main__":
    try:
        # Step 1: Run browser automation & grab direct DOM rows
        direct_rows = fetch_cause_list_pdf("causelist.pdf")
        target_date = get_target_date_ist()
        
        # Step 2: Prefer direct DOM rows
        if direct_rows and len(direct_rows) > 0:
            print(f"[+] Writing {len(direct_rows)} direct DOM rows to Excel...")
            write_to_excel(direct_rows, "cause_list.xlsx")
            generate_pdf_summary(direct_rows, target_date, "cause_list_summary.pdf")
        else:
            print("[*] Direct DOM table was empty; triggering multi-model AI parsing on PDF...")
            ai_rows = parse_pdf_with_gemini("causelist.pdf", "cause_list.xlsx")
            generate_pdf_summary(ai_rows, target_date, "cause_list_summary.pdf")
            
    except Exception as err:
        print(f"[FATAL] Process aborted: {err}", file=sys.stderr)
        traceback.print_exc()
        sys.exit(1)
