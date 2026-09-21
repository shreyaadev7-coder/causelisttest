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
    return next_day.strftime("%d/%m/%Y")

def fetch_cause_list_pdf(pdf_path="causelist.pdf"):
    date_str = get_next_day_ist()
    print(f"[*] Targeting Causelist Date (IST): {date_str}")

    proxy_server = os.environ.get("PROXY_SERVER") # Optional: "http://ip:port"
    launch_args = {
        "headless": True,
        "args": ["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"]
    }
    if proxy_server:
        launch_args["proxy"] = {"server": proxy_server}

    with sync_playwright() as p:
        browser = p.chromium.launch(**launch_args)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
            viewport={"width": 1280, "height": 900},
            extra_http_headers={
                "Accept-Language": "en-US,en;q=0.9,kn;q=0.8",
                "Referer": "https://judiciary.karnataka.gov.in/"
            }
        )
        page = context.new_page()

        print("[*] Navigating to High Court Portal...")
        try:
            # Use 'commit' to prevent hangs on stalled scripts
            response = page.goto(
                "https://judiciary.karnataka.gov.in/causelistSearch.php", 
                wait_until="commit", 
                timeout=45000
            )
            print(f"[*] Initial HTTP Response Status: {response.status if response else 'None'}")
            page.wait_for_timeout(4000)
        except Exception as e:
            print(f"[FATAL] Could not establish connection to the court website: {e}")
            page.screenshot(path="failure.png")
            browser.close()
            raise

        # Check if page actually loaded form elements
        if not page.locator("select").first.is_visible():
            print("[FATAL] Page loaded but cause list form elements were not found (Possible IP Block/WAF).")
            page.screenshot(path="failure.png")
            with open("page_dump.html", "w", encoding="utf-8") as f:
                f.write(page.content())
            browser.close()
            raise RuntimeError("Court portal blocked access or failed to display the search form.")

        # 1. Bench Selection
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

        # 4. Dates
        date_inputs = page.locator("input[placeholder*='DD/MM/YYYY'], input[name*='date'], input[id*='date']").all()
        if len(date_inputs) >= 2:
            date_inputs[0].fill(date_str)
            date_inputs[1].fill(date_str)
        else:
            for inp in page.locator("input[type='text']:visible").all()[1:]:
                inp.fill(date_str)

        page.wait_for_timeout(1000)

        # 5. Print List Trigger
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
            # Fallback if opened in same tab
            page.wait_for_timeout(2000)
            page.pdf(path=pdf_path, format="A4", print_background=True)
            print(f"[+] Saved cause list PDF via current page to {pdf_path}")

        browser.close()

def parse_pdf_to_excel(pdf_path="causelist.pdf", excel_path="cause_list.xlsx"):
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("GEMINI_API_KEY environment variable is missing.")

    client = genai.Client(api_key=api_key)

    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()

    print("[*] Submitting cause list PDF to Gemini 3.6 Flash...")
    prompt = (
        "Extract the complete cause list table from this PDF into the structured JSON schema. "
        "Strictly preserve exact cell contents, case numbers, party names, judge titles, and status text. "
        "If there are no cases listed or the list is empty, return an empty rows array."
    )

    response = client.models.generate_content(
        model="gemini-3.6-flash",
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

    parsed: CauseListDocument = CauseListDocument.model_validate_json(response.text)
    print(f"[+] Extracted {len(parsed.rows)} rows for date: {parsed.date}")

    # Build Excel
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
