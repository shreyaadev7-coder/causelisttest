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

# Schema strictly matching the High Court PDF columns
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
    # Check if manual date passed via GitHub Actions input
    manual = os.environ.get("TEST_DATE")
    if manual and manual.strip():
        print(f"[*] Overriding date with TEST_DATE: {manual.strip()}")
        return manual.strip()

    # Hardcoded test date (set to 22/09/2026)
    return "22/09/2026"

    # Production logic:
    # ist = pytz.timezone('Asia/Kolkata')
    # next_day = datetime.now(ist) + timedelta(days=1)
    # return next_day.strftime("%d/%m/%Y")

def fetch_cause_list_pdf(pdf_path="causelist.pdf"):
    date_str = get_target_date_ist()
    print(f"[*] Target Causelist Date: {date_str}")

    scraperapi_key = os.environ.get("SCRAPERAPI_KEY")
    if not scraperapi_key:
        raise ValueError("SCRAPERAPI_KEY environment variable is missing.")

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

    with sync_playwright() as p:
        browser = p.chromium.launch(**launch_args)
        context = browser.new_context(
            ignore_https_errors=True,
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
            viewport={"width": 1366, "height": 900}
        )
        page = context.new_page()

        print("[*] Connecting to High Court Portal via Indian Gateway...")
        page.goto("https://judiciary.karnataka.gov.in/causelistSearch.php", wait_until="domcontentloaded", timeout=90000)
        page.wait_for_timeout(3000)

        # 1. Bench Selection
        print("[*] Selecting Bench: Bengaluru Bench")
        bench_select = page.locator("select").first
        bench_select.select_option(label="Bengaluru Bench")
        page.wait_for_timeout(2000)  # Allow AJAX update

        # 2. Search By: Advocate
        print("[*] Selecting Search By: Advocate")
        # Search dropdowns for one containing 'Advocate'
        search_by_select = None
        for sel in page.locator("select").all():
            options_text = sel.inner_text()
            if "Advocate" in options_text:
                search_by_select = sel
                break
        
        if search_by_select:
            search_by_select.select_option(label="Advocate")
        else:
            page.locator("select").nth(1).select_option(label="Advocate")
        
        # Give page time to load the dynamic Advocate input field
        page.wait_for_timeout(3000)

        # 3. Enter Dates (From / To or single Causelist Date)
        print(f"[*] Entering Causelist Date: {date_str}")
        date_fields = page.locator("input[placeholder*='DD/MM/YYYY'], input[name*='date' i], input[id*='date' i], input.hasDatepicker").all()
        if date_fields:
            for d_field in date_fields:
                d_field.click()
                d_field.fill("")
                d_field.fill(date_str)
        else:
            print("[!] Warning: Specific date fields not found by attribute, searching visible text inputs...")

        # 4. Enter Advocate Name
        print("[*] Locating Advocate Name input...")
        adv_field = None
        # Check by name or id containing 'adv'
        adv_candidates = page.locator("input[name*='adv' i], input[id*='adv' i], input[placeholder*='adv' i]").all()
        if adv_candidates:
            adv_field = adv_candidates[0]
        else:
            # Fallback: find input preceded by 'Advocate' text in table/form
            adv_field = page.locator("//tr[contains(., 'Advocate')]//input[@type='text'] | //div[contains(., 'Advocate')]//input[@type='text']").first

        if not adv_field or not adv_field.is_visible():
            # Final fallback: find the text input that does not contain the date
            all_text_inputs = page.locator("input[type='text']:visible").all()
            for inp in all_text_inputs:
                val = inp.input_value()
                if val != date_str:
                    adv_field = inp
                    break

        if adv_field:
            adv_field.click()
            adv_field.fill("")
            adv_field.fill("Mahesh Chowdhary")
            print(f"[+] Advocate field filled. Verified Value: '{adv_field.input_value()}'")
        else:
            print("[FATAL] Unable to isolate Advocate input box!")

        # Take screenshot of filled form
        page.screenshot(path="01_form_filled.png")
        print("[*] Saved screenshot: 01_form_filled.png")

        # 5. CLICK 'GET DETAILS' FIRST
        print("[*] Submitting search via 'GET DETAILS'...")
        get_details_btn = page.locator("input[value*='GET DETAILS' i], input[value*='DETAILS' i], button:has-text('GET DETAILS'), button:has-text('DETAILS')").first
        
        if get_details_btn.is_visible():
            get_details_btn.click()
            print("[*] 'GET DETAILS' clicked. Waiting for records to load...")
            page.wait_for_timeout(5000)
            page.wait_for_load_state("networkidle")
        else:
            print("[!] 'GET DETAILS' button not explicitly found, attempting generic search/submit...")
            page.locator("input[type='submit']:visible, button[type='submit']:visible").first.click()
            page.wait_for_timeout(5000)

        # Screenshot after search
        page.screenshot(path="02_search_results.png")
        print("[*] Saved screenshot: 02_search_results.png")

        # Check if table appeared on page
        tables = page.locator("table").all()
        print(f"[*] Detected {len(tables)} tables on page after search.")

        # 6. CLICK 'PRINT LIST' TO GET CLEAN PDF
        print("[*] Locating 'PRINT LIST' button...")
        print_btn = page.locator("input[value*='PRINT' i], button:has-text('PRINT' i)").first
        
        if print_btn.is_visible():
            try:
                with context.expect_page(timeout=10000) as popup_info:
                    print_btn.click()
                target_page = popup_info.value
                target_page.wait_for_load_state("domcontentloaded")
                page.wait_for_timeout(2000)
                target_page.pdf(path=pdf_path, format="A4", print_background=True)
                print(f"[+] Successfully captured cause list PDF via print popup to {pdf_path}")
            except Exception as e:
                print(f"[!] Popup did not open ({e}), capturing PDF from current page...")
                page.pdf(path=pdf_path, format="A4", print_background=True)
        else:
            print("[*] 'PRINT LIST' button not present; rendering page directly to PDF...")
            page.pdf(path=pdf_path, format="A4", print_background=True)

        browser.close()

def parse_pdf_to_excel(pdf_path="causelist.pdf", excel_path="cause_list.xlsx"):
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

    model_name = os.environ.get("GEMINI_MODEL", "gemini-3.6-flash")
    max_retries = 5
    response = None

    for attempt in range(1, max_retries + 1):
        try:
            print(f"[*] Submitting cause list PDF to {model_name} (Attempt {attempt}/{max_retries})...")
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
            break
        except Exception as e:
            err_str = str(e)
            if "503" in err_str or "UNAVAILABLE" in err_str or "ResourceExhausted" in err_str:
                if attempt < max_retries:
                    wait_time = attempt * 8
                    print(f"[!] Demand spike (503). Pausing {wait_time}s before retry...")
                    time.sleep(wait_time)
                    continue
            raise

    parsed: CauseListDocument = CauseListDocument.model_validate_json(response.text)
    print(f"[+] Successfully extracted {len(parsed.rows)} rows for date: {parsed.date}")

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

    for row in parsed.rows:
        ws.append([
            row.sl_no,
            row.case_number,
            row.case_name,
            row.ch,
            row.list_num,
            row.list_sl_no,
            row.status,
            row.judges
        ])

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
    print(f"[+] Excel written to {excel_path}")

if __name__ == "__main__":
    try:
        fetch_cause_list_pdf("causelist.pdf")
        parse_pdf_to_excel("causelist.pdf", "cause_list.xlsx")
    except Exception as err:
        print(f"[FATAL] Process aborted: {err}", file=sys.stderr)
        traceback.print_exc()
        sys.exit(1)
