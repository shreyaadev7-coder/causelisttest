import os
import sys
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
                # Ensure exactly 8 elements
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

    headers = ["SL NO", "CASE NUMBER", "CASE NAME", "CH", "LIST", "SL NO", "STATUS", "JUDGES"]
    ws.append(headers)

    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True)
    regular_font = Font(name="Calibri", size=10)
    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC")
    )

    for col_idx in range(1, 9):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    for row in rows:
        if isinstance(row, CauseListRow):
            ws.append([row.sl_no, row.case_number, row.case_name, row.ch, row.list_num, row.list_sl_no, row.status, row.judges])
        else:
            ws.append(row[:8])

    for r in range(2, ws.max_row + 1):
        for c in range(1, 9):
            cell = ws.cell(row=r, column=c)
            cell.font = regular_font
            cell.border = thin_border
            if c in [1, 4, 5, 6]:
                cell.alignment = Alignment(horizontal="center", vertical="top")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

    col_widths = {1: 8, 2: 20, 3: 35, 4: 8, 5: 8, 6: 8, 7: 25, 8: 30}
    for col_idx, width in col_widths.items():
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = width

    wb.save(excel_path)
    print(f"[+] Excel successfully generated at: {excel_path}")

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

    # Multi-model rotation list to prevent 503 deadlocks
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

if __name__ == "__main__":
    try:
        # Step 1: Run browser automation & grab direct DOM rows
        direct_rows = fetch_cause_list_pdf("causelist.pdf")
        
        # Step 2: Prefer direct DOM rows (0ms latency, immune to 503 errors)
        if direct_rows and len(direct_rows) > 0:
            print(f"[+] Writing {len(direct_rows)} direct DOM rows to Excel...")
            write_to_excel(direct_rows, "cause_list.xlsx")
        else:
            print("[*] Direct DOM table was empty; triggering multi-model AI parsing on PDF...")
            parse_pdf_with_gemini("causelist.pdf", "cause_list.xlsx")
            
    except Exception as err:
        print(f"[FATAL] Process aborted: {err}", file=sys.stderr)
        traceback.print_exc()
        sys.exit(1)
