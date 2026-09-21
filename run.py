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
    # Priority 1: GitHub Actions manual input if provided
    manual = os.environ.get("TEST_DATE")
    if manual and manual.strip():
        print(f"[*] Overriding date with TEST_DATE: {manual.strip()}")
        return manual.strip()

    # Priority 2: Hardcoded test date for verification
    return "22/09/2026"

    # Production logic (uncomment when testing is complete):
    # ist = pytz.timezone('Asia/Kolkata')
    # next_day = datetime.now(ist) + timedelta(days=1)
    # return next_day.strftime("%d/%m/%Y")

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

    with sync_playwright() as p:
        browser = p.chromium.launch(**launch_args)
        context = browser.new_context(
            ignore_https_errors=True,
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
            viewport={"width": 1366, "height": 900}
        )
        page = context.new_page()

        print("[*] Navigating to High Court Portal via Indian Gateway...")
        page.goto("https://judiciary.karnataka.gov.in/causelistSearch.php", wait_until="domcontentloaded", timeout=90000)
        page.wait_for_timeout(3000)
        page.screenshot(path="00_initial_page.png")

        # 1. Bench Selection
        print("[*] Step 1: Selecting Bench -> Bengaluru Bench")
        bench_select = page.locator("select:visible").first
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
        selected_search_by = False
        for sel in page.locator("select:visible").all():
            for opt in sel.locator("option").all():
                opt_text = opt.inner_text().strip()
                if "Advocate" in opt_text:
                    sel.select_option(label=opt_text)
                    print(f"[+] Dropdown selected: '{opt_text}'")
                    selected_search_by = True
                    break
            if selected_search_by:
                break

        # Check for radio button if not a dropdown
        if not selected_search_by:
            for r in page.locator("input[type='radio']:visible, label:has-text('Advocate')").all():
                val = r.get_attribute("value") or ""
                text = r.inner_text().strip() if hasattr(r, 'inner_text') else ""
                if "adv" in val.lower() or "advocate" in text.lower():
                    r.click()
                    print(f"[+] Clicked Advocate radio button (value='{val}')")
                    selected_search_by = True
                    break

        page.wait_for_timeout(3000)

        # Print all visible inputs to the console for debugging
        print("\n[*] --- Detected Visible Form Controls ---")
        for inp in page.locator("input:visible, select:visible").all():
            tag = inp.evaluate("el => el.tagName.toLowerCase()")
            el_type = inp.get_attribute("type") or ""
            el_id = inp.get_attribute("id") or ""
            name = inp.get_attribute("name") or ""
            ph = inp.get_attribute("placeholder") or ""
            print(f"    <{tag} type='{el_type}' id='{el_id}' name='{name}' placeholder='{ph}'>")
        print("[*] ----------------------------------------\n")

        # 3. Enter Date into visible date inputs
        print(f"[*] Step 3: Entering Date -> {date_str}")
        for inp in page.locator("input:visible").all():
            el_type = (inp.get_attribute("type") or "").lower()
            if el_type in ["hidden", "submit", "button", "radio", "checkbox"]:
                continue
            ph = (inp.get_attribute("placeholder") or "").lower()
            name = (inp.get_attribute("name") or "").lower()
            el_id = (inp.get_attribute("id") or "").lower()
            cls = (inp.get_attribute("class") or "").lower()
            
            # Identify date input by placeholder, class, or name
            if "dd/mm/yyyy" in ph or "datepicker" in cls or "date" in name or "date" in el_id or "dt" in name or "dt" in el_id:
                inp.fill(date_str)
                print(f"[+] Date set to '{date_str}' in <input id='{el_id}' name='{name}'>")

        page.wait_for_timeout(1000)

        # 4. Enter Advocate Name
        print("[*] Step 4: Entering Advocate Name -> Mahesh Chowdhary")
        adv_input = None
        
        # Priority A: Check for input explicitly referencing advocate
        for inp in page.locator("input:visible").all():
            el_type = (inp.get_attribute("type") or "").lower()
            if el_type in ["hidden", "submit", "button", "radio", "checkbox"]:
                continue
            ph = (inp.get_attribute("placeholder") or "").lower()
            name = (inp.get_attribute("name") or "").lower()
            el_id = (inp.get_attribute("id") or "").lower()
            
            if ("adv" in name or "adv" in el_id or "adv" in ph) and ("dt" not in name and "date" not in name and "dd/mm/yyyy" not in ph):
                adv_input = inp
                break

        # Priority B: First visible text input that is not a date field
        if not adv_input:
            for inp in page.locator("input[type='text']:visible, input:not([type]):visible").all():
                ph = (inp.get_attribute("placeholder") or "").lower()
                name = (inp.get_attribute("name") or "").lower()
                el_id = (inp.get_attribute("id") or "").lower()
                if "dd/mm/yyyy" not in ph and "date" not in name and "dt" not in name and "dt" not in el_id:
                    adv_input = inp
                    break

        if adv_input:
            adv_input.fill("Mahesh Chowdhary")
            print(f"[+] Advocate Name verified in field: '{adv_input.input_value()}'")
        else:
            print("[FATAL] Could not isolate Advocate Name input box!")

        page.wait_for_timeout(1000)
        page.screenshot(path="01_form_filled.png")
        print("[*] Saved screenshot: 01_form_filled.png")

        # 5. Click GET DETAILS
        print("[*] Step 5: Submitting search via 'GET DETAILS'...")
        get_btn = page.locator("input[value*='GET DETAILS' i]:visible, button:has-text('GET DETAILS'):visible, input[value*='DETAILS' i]:visible, button:has-text('DETAILS'):visible").first
        if get_btn.is_visible():
            get_btn.click()
        else:
            page.locator("input[type='submit']:visible, button[type='submit']:visible").first.click()

        print("[*] Waiting 6 seconds for High Court database results...")
        page.wait_for_timeout(6000)
        page.screenshot(path="02_search_results.png")
        print("[*] Saved screenshot: 02_search_results.png")

        # 6. Click PRINT LIST to generate the official printable view
        print("[*] Step 6: Triggering 'PRINT LIST'...")
        print_btn = page.locator("input[value*='PRINT' i]:visible, button:has-text('PRINT' i):visible").first
        
        if print_btn.is_visible():
            try:
                with context.expect_page(timeout=12000) as popup_info:
                    print_btn.click()
                print_page = popup_info.value
                print_page.wait_for_load_state("domcontentloaded")
                print_page.wait_for_timeout(3000)
                print_page.pdf(path=pdf_path, format="A4", print_background=True)
                print_page.screenshot(path="03_print_view.png")
                print(f"[+] Captured cause list PDF via print window: {pdf_path}")
            except Exception as e:
                print(f"[*] Popup wait finished ({e}). Rendering main page to PDF...")
                page.wait_for_timeout(2000)
                page.pdf(path=pdf_path, format="A4", print_background=True)
                page.screenshot(path="03_print_view.png")
                print(f"[+] Saved cause list PDF: {pdf_path}")
        else:
            print("[*] 'PRINT LIST' button not present; rendering results directly to PDF...")
            page.pdf(path=pdf_path, format="A4", print_background=True)
            page.screenshot(path="03_print_view.png")

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
        "Do not omit any row. If there are no cases listed, return an empty rows array."
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
                    print(f"[!] Server busy (503). Pausing {wait_time}s before retry...")
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
