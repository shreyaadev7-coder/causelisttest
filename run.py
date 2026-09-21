import os
import sys
import re
from datetime import datetime, timedelta
import pytz
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from google import genai
from google.genai import types
from pydantic import BaseModel, Field
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

class CauseListRow(BaseModel):
    sl_no: str = Field(description="First SL NO column")
    case_number: str = Field(description="CASE NUMBER")
    case_name: str = Field(description="Full CASE NAME")
    ch: str = Field(description="Court Hall number")
    list_num: str = Field(description="List number")
    list_sl_no: str = Field(description="Second SL NO column")
    status: str = Field(description="STATUS")
    judges: str = Field(description="JUDGES")

class CauseListDocument(BaseModel):
    date: str = Field(description="Date of the cause list")
    rows: list[CauseListRow]

def get_next_day_ist():
    ist = pytz.timezone('Asia/Kolkata')
    next_day = datetime.now(ist) + timedelta(days=1)
    return next_day.strftime("%d-%m-%Y"), next_day.strftime("%d/%m/%Y")

def fetch_cause_list_pdf(pdf_path="causelist.pdf"):
    date_dash, date_slash = get_next_day_ist()
    print(f"[*] Target Causelist Date (IST): {date_dash} ({date_slash})")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"]
        )
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            accept_downloads=True
        )
        page = context.new_page()

        try:
            print("[*] Navigating to Karnataka High Court cause list search...")
            response = page.goto("https://judiciary.karnataka.gov.in/causelistSearch.php", timeout=60000)
            print(f"[*] Page loaded with HTTP status: {response.status if response else 'Unknown'}")

            # 1. Select Bench
            print("[*] Selecting Bengaluru Bench...")
            bench_select = page.locator("select").first
            bench_select.select_option(label="Bengaluru Bench")

            # 2. Causelist Date
            print(f"[*] Setting date: {date_slash}...")
            # Target any date input field or text input
            date_input = page.locator("input[type='text'], input[name*='date']").first
            date_input.fill(date_slash)

            # 3. Select Advocate Search
            print("[*] Selecting Advocate filter...")
            adv_radio = page.locator("input[type='radio'][value*='adv'], input[type='radio'][value*='A'], input[type='radio']").nth(0)
            adv_radio.check()

            # 4. Advocate Name
            print("[*] Entering Advocate Name: Mahesh Chowdhary...")
            adv_name_input = page.locator("input[name*='adv'], input[id*='adv'], input[type='text']").nth(1)
            adv_name_input.fill("Mahesh Chowdhary")

            # 5. Handle Captcha if present
            captcha_img = page.locator("img[src*='captcha'], #captcha_image, img")
            if captcha_img.count() > 0 and "captcha" in (captcha_img.first.get_attribute("src") or "").lower():
                print("[*] Captcha detected. Solving via Gemini...")
                captcha_img.first.screenshot(path="captcha.png")
                ai_client = genai.Client()
                with open("captcha.png", "rb") as f:
                    c_bytes = f.read()
                cap_res = ai_client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=[
                        types.Part.from_bytes(data=c_bytes, mime_type="image/png"),
                        "Output only the alphanumeric characters in this captcha image. No explanation."
                    ]
                )
                cap_val = re.sub(r'[^A-Za-z0-9]', '', cap_res.text.strip())
                print(f"[*] Solved Captcha: {cap_val}")
                page.locator("input[name*='captcha'], #captcha").first.fill(cap_val)

            # 6. Click Submit & Wait for Download
            print("[*] Submitting search...")
            submit_btn = page.locator("input[type='submit'], button[type='submit'], input[value*='Search'], button:has-text('Search')").first
            
            with page.expect_download(timeout=20000) as download_info:
                submit_btn.click()
            
            download = download_info.value
            download.save_as(pdf_path)
            print(f"[✓] Successfully downloaded PDF: {pdf_path}")
            return True

        except PlaywrightTimeoutError:
            # Check if an on-screen alert or message appeared
            page_text = page.locator("body").inner_text()
            page.screenshot(path="failure.png")
            print("[!] Download timed out. Saved screenshot to failure.png.")

            if "no record" in page_text.lower() or "not found" in page_text.lower():
                print("[*] Court portal returned: No records / cause list not yet published for tomorrow.")
                return False
            else:
                print(f"[!] Current page snippet: {page_text[:300]}")
                raise

        except Exception as e:
            page.screenshot(path="failure.png")
            print(f"[!] Error during scraping: {e}")
            raise

        finally:
            browser.close()

def create_empty_excel(excel_path="cause_list.xlsx", message="No matters listed for tomorrow."):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cause List"
    headers = ["SL NO", "CASE NUMBER", "CASE NAME", "CH", "LIST", "SL NO", "STATUS", "JUDGES"]
    ws.append(headers)
    ws.append([message] + [""] * 7)
    wb.save(excel_path)
    print(f"[*] Empty cause list placeholder written to {excel_path}")

def parse_pdf_to_excel(pdf_path="causelist.pdf", excel_path="cause_list.xlsx"):
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("GEMINI_API_KEY secret is not set in repository environment!")

    client = genai.Client()
    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()

    print("[*] Sending PDF to Gemini 2.5 Flash for table extraction...")
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
    print(f"[✓] Extracted {len(parsed.rows)} rows for date: {parsed.date}")

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
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    for row in parsed.rows:
        ws.append([
            row.sl_no, row.case_number, row.case_name,
            row.ch, row.list_num, row.list_sl_no,
            row.status, row.judges
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
    print(f"[✓] Excel workbook created at {excel_path}")

if __name__ == "__main__":
    has_pdf = fetch_cause_list_pdf("causelist.pdf")
    if has_pdf and os.path.exists("causelist.pdf"):
        parse_pdf_to_excel("causelist.pdf", "cause_list.xlsx")
    else:
        create_empty_excel("cause_list.xlsx", "No cases listed or cause list not yet released.")
