import os
import re
from datetime import datetime, timedelta
import pytz
from playwright.sync_api import sync_playwright
from google import genai
from google.genai import types
from pydantic import BaseModel, Field
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Define strict schema matching the exact PDF columns
class CauseListRow(BaseModel):
    sl_no: str = Field(description="First SL NO column")
    case_number: str = Field(description="CASE NUMBER (e.g., WP NO 102709/2026)")
    case_name: str = Field(description="Full CASE NAME with party details and respondent notes")
    ch: str = Field(description="Court Hall number under CH")
    list_num: str = Field(description="List number under LIST")
    list_sl_no: str = Field(description="Second SL NO column (item serial number)")
    status: str = Field(description="STATUS column (e.g., ORDERS, HEARING - IA)")
    judges: str = Field(description="JUDGES column with full bench / judge names")

class CauseListDocument(BaseModel):
    date: str = Field(description="Date displayed at the top of the cause list")
    rows: list[CauseListRow]

def get_next_day_ist():
    ist = pytz.timezone('Asia/Kolkata')
    next_day = datetime.now(ist) + timedelta(days=1)
    # Court portal format is typically DD-MM-YYYY or DD/MM/YYYY
    return next_day.strftime("%d-%m-%Y"), next_day.strftime("%d/%m/%Y")

def fetch_cause_list_pdf(pdf_path="causelist.pdf"):
    date_dash, date_slash = get_next_day_ist()
    print(f"Target Causelist Date (IST): {date_dash}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        page.goto("https://judiciary.karnataka.gov.in/causelistSearch.php", timeout=60000)
        page.wait_for_load_state("networkidle")

        # 1. Bench Selection
        bench_select = page.locator("select[name*='bench'], #bench, select:has-text('Bengaluru')").first
        if bench_select.is_visible():
            bench_select.select_option(label="Bengaluru Bench")

        # 2. Date Input
        date_input = page.locator("input[name*='date'], input[id*='date']").first
        if date_input.is_visible():
            date_input.fill(date_slash)

        # 3. Search By Advocate
        adv_radio = page.locator("input[type='radio'][value*='adv'], input[value*='A']").first
        if adv_radio.is_visible():
            adv_radio.check()

        # 4. Advocate Name
        adv_name_input = page.locator("input[name*='adv'], input[id*='adv']").first
        if adv_name_input.is_visible():
            adv_name_input.fill("Mahesh Chowdhary")

        # Handle Captcha if present on the form
        captcha_img = page.locator("img[src*='captcha'], #captcha_image")
        if captcha_img.is_visible():
            captcha_box = page.locator("input[name*='captcha'], #captcha")
            # If a simple text captcha exists, screenshot and read it via Gemini
            captcha_img.screenshot(path="captcha.png")
            ai_client = genai.Client()
            with open("captcha.png", "rb") as f:
                c_bytes = f.read()
            cap_res = ai_client.models.generate_content(
                model="gemini-2.5-flash",
                contents=[
                    types.Part.from_bytes(data=c_bytes, mime_type="image/png"),
                    "Extract only the alphanumeric characters in this captcha image. No whitespace."
                ]
            )
            captcha_val = re.sub(r'[^A-Za-z0-9]', '', cap_res.text.strip())
            captcha_box.fill(captcha_val)

        # 5. Submit Form and catch download
        submit_btn = page.locator("input[type='submit'], button:has-text('Submit'), button:has-text('Search')").first
        with page.expect_download(timeout=60000) as download_info:
            submit_btn.click()

        download = download_info.value
        download.save_as(pdf_path)
        print(f"Downloaded cause list PDF to {pdf_path}")
        browser.close()

def parse_pdf_to_excel(pdf_path="causelist.pdf", excel_path="cause_list.xlsx"):
    client = genai.Client()

    # Read the downloaded PDF directly
    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()

    print("Submitting PDF to Gemini for structured extraction...")
    prompt = (
        "Extract the exact cause list table from this PDF document into the structured schema. "
        "Preserve exact cell contents, party names, judge titles, case numbers, and status text. "
        "Do not omit any row or column."
    )

    response = client.models.generate_content(
        model="gemini-2.5-flash",
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
    print(f"Extracted {len(parsed.rows)} rows for date: {parsed.date}")

    # Build Excel with exact matching headers
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cause List"

    # Exact headers from the document
    headers = ["SL NO", "CASE NUMBER", "CASE NAME", "CH", "LIST", "SL NO", "STATUS", "JUDGES"]
    ws.append(headers)

    # Style definitions
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True)
    regular_font = Font(name="Calibri", size=10)
    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC")
    )

    # Apply header formatting
    for col_idx in range(1, 9):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    # Insert rows
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

    # Format data rows
    for r in range(2, ws.max_row + 1):
        for c in range(1, 9):
            cell = ws.cell(row=r, column=c)
            cell.font = regular_font
            cell.border = thin_border
            # Left align text columns; center short codes and numbers
            if c in [1, 4, 5, 6]:
                cell.alignment = Alignment(horizontal="center", vertical="top")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

    # Set column widths for readability
    col_widths = {1: 8, 2: 20, 3: 35, 4: 8, 5: 8, 6: 8, 7: 25, 8: 30}
    for col_idx, width in col_widths.items():
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = width

    wb.save(excel_path)
    print(f"Excel workbook created at {excel_path}")

if __name__ == "__main__":
    fetch_cause_list_pdf("causelist.pdf")
    parse_pdf_to_excel("causelist.pdf", "cause_list.xlsx")
