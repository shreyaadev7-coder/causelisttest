import os
import sys
import traceback
from datetime import datetime, timedelta
import pytz
from playwright.sync_api import sync_playwright
from google import genai
from google.genai import types
from pydantic import BaseModel, Field
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Define schema matching exact PDF columns
class CauseListRow(BaseModel):
    sl_no: str = Field(description="First SL NO column")
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

def get_next_day_ist():
    ist = pytz.timezone('Asia/Kolkata')
    next_day = datetime.now(ist) + timedelta(days=1)
    #return next_day.strftime("%d/%m/%Y")
    return "22/09/2026"

def fetch_cause_list_pdf(pdf_path="causelist.pdf"):
    date_str = get_next_day_ist()
    print(f"[*] Target Causelist Date (IST): {date_str}")

    scraperapi_key = os.environ.get("SCRAPERAPI_KEY")
    if not scraperapi_key:
        raise ValueError("SCRAPERAPI_KEY environment variable is missing from GitHub Secrets.")

    # Route Playwright through ScraperAPI's Indian gateway
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
            viewport={"width": 1280, "height": 900}
        )
        page = context.new_page()

        print("[*] Navigating to High Court Portal via Indian Gateway...")
        page.goto(
            "https://judiciary.karnataka.gov.in/causelistSearch.php", 
            wait_until="domcontentloaded", 
            timeout=90000
        )
        print("[+] Portal successfully connected.")
        page.wait_for_timeout(2000)

        # 1. Bench Selection: Bengaluru Bench
        bench = page.locator("select").first
        bench.select_option(label="Bengaluru Bench")
        page.wait_for_timeout(1000)

        # 2. Search By: Advocate
        search_by = page.locator("select").nth(1)
        try:
            search_by.select_option(label="Advocate")
        except Exception:
            page.select_option("select:has-text('Advocate')", label="Advocate")
        page.wait_for_timeout(1000)

        # 3. Advocate Name
        adv_input = page.locator("input[type='text']:visible").first
        adv_input.fill("Mahesh Chowdhary")

        # 4. Dates (From and To)
        date_inputs = page.locator("input[placeholder*='DD/MM/YYYY'], input[name*='date'], input[id*='date']").all()
        if len(date_inputs) >= 2:
            date_inputs[0].fill(date_str)
            date_inputs[1].fill(date_str)
        else:
            for inp in page.locator("input[type='text']:visible").all()[1:]:
                inp.fill(date_str)

        page.wait_for_timeout(1000)

        # 5. Click PRINT LIST to generate printable cause list view
        print_btn = page.locator("input[value*='PRINT'], button:has-text('PRINT')").first
        
        try:
            with context.expect_page(timeout=15000) as popup_info:
                print_btn.click()
            target_page = popup_info.value
            target_page.wait_for_load_state("domcontentloaded")
            page.wait_for_timeout(2000)
            target_page.pdf(path=pdf_path, format="A4", print_background=True)
            print(f"[+] Saved cause list PDF via popup to {pdf_path}")
        except Exception:
            page.wait_for_timeout(2000)
            page.pdf(path=pdf_path, format="A4", print_background=True)
            print(f"[+] Saved cause list PDF to {pdf_path}")

        browser.close()

import time

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
        "If there are no cases listed or the list is empty, return an empty rows array."
    )

    model_name = os.environ.get("GEMINI_MODEL", "gemini-3.6-flash")

    # Retry loop with exponential backoff for 503 temporary demand spikes
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
            # If successful, break out of retry loop
            break
        except Exception as e:
            err_str = str(e)
            if "503" in err_str or "UNAVAILABLE" in err_str or "ResourceExhausted" in err_str:
                if attempt < max_retries:
                    wait_time = attempt * 8  # 8s, 16s, 24s, 32s
                    print(f"[!] Server busy (503/Spike). Pausing {wait_time}s before retry...")
                    time.sleep(wait_time)
                    continue
            # If not a temporary demand error, or out of retries, raise
            raise

    parsed: CauseListDocument = CauseListDocument.model_validate_json(response.text)
    print(f"[+] Extracted {len(parsed.rows)} rows for date: {parsed.date}")

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
