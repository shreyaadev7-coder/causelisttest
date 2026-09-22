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

ADVOCATES = [
    "Mahesh Chowdhary",
    "Nagaraja Naidu",
    "Abhimanyu",
    "Krishika Vaishnav",
    "Shahbaaz Hussain"
]

ADVOCATE_REGEX = r'\b(MAHESH|CHOWDHA?R[YI]|NAGARAJA?|NAIDU|ABHIMANYU|KRISHIKA|VAISHNAV|SHAHBAAZ|HUSSAIN)\b'

class CauseListRow(BaseModel):
    sl_no: str = Field(description="Serial number")
    case_number: str = Field(description="Case number like WP NO 102709/2026")
    case_name: str = Field(description="Party names and representation details")
    ch: str = Field(description="Court Hall number")
    list_num: str = Field(description="List number")
    list_sl_no: str = Field(description="Item number")
    status: str = Field(description="Case status like ORDERS")
    judges: str = Field(description="Hon'ble Judge name(s)")

class CauseListDocument(BaseModel):
    date: str = Field(description="Date of the cause list")
    rows: list[CauseListRow]

def get_target_date_ist():
    manual = os.environ.get("TEST_DATE")
    if manual and manual.strip():
        return manual.strip()
    return "22/09/2026"

def clean_party_string(s: str) -> str:
    """Strips out prefixes, advocate names, and procedural notes."""
    s = re.sub(r'^(?:PET|RES|PETITIONER|RESPONDENT|APPELLANT|COMPLAINANT)\s*:\s*', '', s, flags=re.I)
    s = re.sub(r'\([^\)]*\)', '', s)
    # Cut off at ADV: or ADVOCATE:
    adv_split = re.split(r'\b(?:ADV|ADVOCATE|ADVOCATES|COUNSEL)\s*[:.]?\s*', s, flags=re.I)
    s = adv_split[0]
    # Filter out comma-separated lawyer tokens
    chunks = [c.strip() for c in re.split(r'[,;]', s) if c.strip()]
    valid = []
    for c in chunks:
        if not re.search(rf'({ADVOCATE_REGEX}|AGA|HCGP|ADV|ADVOCATE|ADVOCATES|COUNSEL|FOR\s+RES|FOR\s+PET|GOVT|PLEADER)', c, re.I):
            valid.append(c)
    res = ", ".join(valid) if valid else (chunks[0] if chunks else s)
    return res.strip(" ,;:-")

def clean_case_details(raw_name: str):
    """
    Cleans raw case name into:
    [Petitioner] vs [Respondent]
    Identifies which party our advocate team represents.
    """
    text = re.sub(r'[\r\n]+', ' ', raw_name)
    text = re.sub(r'\s+', ' ', text).strip()

    vs_match = re.search(r'\s+(?:-?\s*V/?S\.?\s*-?|VERSUS)\s+', text, re.I)
    res_match = re.search(r'[\s\-]+(?:RES|RESPONDENT|RESPODNENT)\s*:\s*', text, re.I)

    if vs_match:
        pet_block = text[:vs_match.start()].strip()
        res_block = text[vs_match.end():].strip()
    elif res_match:
        pet_block = text[:res_match.start()].strip()
        res_block = text[res_match.start():].strip()
    else:
        pet_block = text
        res_block = ""

    # Detect which side our advocates represent
    if re.search(ADVOCATE_REGEX, res_block, re.I):
        in_charge = "RES"
    elif re.search(ADVOCATE_REGEX, pet_block, re.I):
        in_charge = "PET"
    elif bool(re.search(r'\b(RESPONDENT|RES)\b.*?\b(NO|NOS|R\d+|\d+)\b', text, re.I)):
        in_charge = "RES"
    else:
        in_charge = "PET"

    pet_clean = clean_party_string(pet_block)
    res_clean = clean_party_string(res_block)

    return pet_clean, res_clean, in_charge

def parse_pdf_with_gemini(pdf_path: str) -> list[CauseListRow]:
    """Uses Gemini with retry logic to parse actual cause list rows."""
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("GEMINI_API_KEY is missing from environment secrets.")

    client = genai.Client(api_key=api_key)
    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()

    prompt = (
        "Extract the complete cause list table from this PDF into the structured JSON schema. "
        "Strictly preserve exact cell contents, case numbers, party names, judge titles, and status text. "
        "Do not omit any row. If no cases are listed, return an empty rows array."
    )

    models = ["gemini-3.5-flash", "gemini-3.1-flash-lite", "gemini-3.6-flash"]

    for model_name in models:
        for attempt in range(1, 4):
            try:
                print(f"[*] Submitting {pdf_path} to {model_name} (Attempt {attempt})...")
                response = client.models.generate_content(
                    model=model_name,
                    contents=[types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"), prompt],
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        response_schema=CauseListDocument,
                        temperature=0.0
                    )
                )
                parsed = CauseListDocument.model_validate_json(response.text)
                print(f"[+] Extracted {len(parsed.rows)} case records via {model_name}.")
                return parsed.rows
            except Exception as e:
                print(f"[!] {model_name} attempt {attempt} failed ({e}), retrying in {attempt * 3}s...")
                time.sleep(attempt * 3)

    return []

def write_to_excel(rows: list, excel_path="cause_list.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cause List"
    ws.views.sheetView[0].showGridLines = True

    # 1. Column headers in bold (no colors)
    headers = ["SL NO", "CASE NUMBER", "CASE NAME", "CH", "LIST", "SL NO", "STATUS", "JUDGES"]
    ws.append(headers)
    ws.row_dimensions[1].height = 25

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
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    bold_font = InlineFont(rFont="Calibri", sz=10, b=True)
    regular_font = InlineFont(rFont="Calibri", sz=10, b=False)

    for idx, row in enumerate(rows, start=1):
        pet_clean, res_clean, in_charge = clean_case_details(row.case_name)
        plain_text = f"{pet_clean} vs {res_clean}"

        # 2. Sequential SL NO (1, 2, 3...)
        ws.append([
            idx,
            row.case_number.strip(),
            plain_text,
            row.ch.strip(),
            row.list_num.strip(),
            row.list_sl_no.strip(),
            row.status.strip(),
            row.judges.strip()
        ])

        curr_row = idx + 1
        ws.row_dimensions[curr_row].height = 45

        # 3. Bold only the represented party
        case_cell = ws.cell(row=curr_row, column=3)
        try:
            if in_charge == "PET":
                case_cell.value = CellRichText(
                    TextBlock(bold_font, pet_clean),
                    TextBlock(regular_font, f" vs {res_clean}")
                )
            else:
                case_cell.value = CellRichText(
                    TextBlock(regular_font, f"{pet_clean} vs "),
                    TextBlock(bold_font, res_clean)
                )
        except Exception:
            case_cell.value = plain_text

    # Apply borders, text wrapping, and case number bold
    for r in range(2, ws.max_row + 1):
        for c in range(1, 9):
            cell = ws.cell(row=r, column=c)
            cell.border = thin_border
            if c in [1, 4, 5, 6]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(name="Calibri", size=10, bold=False)
            elif c == 2:
                # Case number in bold
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(name="Calibri", size=10, bold=True)
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                if c != 3:
                    cell.font = Font(name="Calibri", size=10, bold=False)

    col_widths = {1: 8, 2: 24, 3: 45, 4: 8, 5: 8, 6: 10, 7: 25, 8: 35}
    for col_idx, width in col_widths.items():
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = width

    wb.save(excel_path)
    print(f"[+] Excel written to {excel_path}")

def generate_pdf(rows: list, date_str: str, pdf_path="cause_list.pdf"):
    """Renders landscape black-and-white table PDF."""
    rows_html = ""
    for idx, row in enumerate(rows, start=1):
        pet_clean, res_clean, in_charge = clean_case_details(row.case_name)
        if in_charge == "PET":
            case_name_cell = f"<strong>{pet_clean}</strong> vs {res_clean}"
        else:
            case_name_cell = f"{pet_clean} vs <strong>{res_clean}</strong>"

        rows_html += f"""
        <tr>
            <td class="text-center">{idx}</td>
            <td class="text-center font-bold">{row.case_number}</td>
            <td>{case_name_cell}</td>
            <td class="text-center">{row.ch}</td>
            <td class="text-center">{row.list_num}</td>
            <td class="text-center">{row.list_sl_no}</td>
            <td>{row.status}</td>
            <td>{row.judges}</td>
        </tr>
        """

    adv_title = ", ".join(ADVOCATES)

    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            @page {{ size: A4 landscape; margin: 10mm; }}
            body {{ font-family: Arial, sans-serif; color: #000; margin: 0; font-size: 11px; }}
            .header {{ display: flex; justify-content: space-between; border-bottom: 1px solid #000; padding-bottom: 4px; margin-bottom: 8px; }}
            .header h1 {{ margin: 0; font-size: 14px; text-transform: uppercase; }}
            table {{ width: 100%; border-collapse: collapse; table-layout: fixed; }}
            th {{ border: 1px solid #000; padding: 6px; font-weight: bold; text-align: center; background: #fff; }}
            td {{ border: 1px solid #000; padding: 6px; vertical-align: middle; word-wrap: break-word; }}
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
                <div style="font-weight: bold; font-size: 10px;">ADVOCATES: {adv_title}</div>
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
    print(f"[+] PDF cause list written to {pdf_path}")

def setup_form_fields(page, date_str: str):
    """Configures Bench, Search By Advocate, and target Date."""
    bench_select = page.locator("select[name='bench']:visible").first
    for opt in bench_select.locator("option").all():
        opt_text = opt.inner_text().strip()
        if "Bengaluru" in opt_text or "Bangalore" in opt_text or "Principal" in opt_text:
            bench_select.select_option(label=opt_text)
            break
    page.wait_for_timeout(1500)

    searchby_select = page.locator("select[name='searchby']:visible").first
    for opt in searchby_select.locator("option").all():
        if "Advocate" in opt.inner_text().strip():
            searchby_select.select_option(label=opt.inner_text().strip())
            break
    page.wait_for_timeout(1500)

    page.locator("#fromDt:visible").first.fill(date_str)
    page.locator("#toDt:visible").first.fill(date_str)
    page.wait_for_timeout(1000)

if __name__ == "__main__":
    try:
        date_str = get_target_date_ist()
        print(f"[*] Target Causelist Date: {date_str}")

        scraperapi_key = os.environ.get("SCRAPERAPI_KEY")
        if not scraperapi_key:
            raise ValueError("SCRAPERAPI_KEY is missing from GitHub Secrets.")

        launch_args = {
            "headless": True,
            "proxy": {
                "server": "http://proxy-server.scraperapi.com:8001",
                "username": "scraperapi.country_code=in",
                "password": scraperapi_key
            },
            "args": ["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage", "--ignore-certificate-errors"]
        }

        all_cases = []
        seen_case_numbers = set()

        # Run all advocates inside one stable browser session
        with sync_playwright() as p:
            browser = p.chromium.launch(**launch_args)
            context = browser.new_context(
                ignore_https_errors=True,
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
                viewport={"width": 1366, "height": 900}
            )
            page = context.new_page()
            page.add_init_script("window.print = () => { console.log('print intercepted'); };")

            print("[*] Navigating to High Court Portal via Indian Gateway...")
            page.goto("https://judiciary.karnataka.gov.in/causelistSearch.php", wait_until="domcontentloaded", timeout=90000)
            page.wait_for_timeout(3000)

            setup_form_fields(page, date_str)

            for idx, advocate_name in enumerate(ADVOCATES, start=1):
                print(f"\n[*] Querying listings for: '{advocate_name}'...")

                # Ensure page hasn't navigated away
                if "causelistSearch.php" not in page.url:
                    page.goto("https://judiciary.karnataka.gov.in/causelistSearch.php", wait_until="domcontentloaded", timeout=60000)
                    page.wait_for_timeout(2000)
                    setup_form_fields(page, date_str)

                adv_input = page.locator("input[placeholder='Enter Advocate Name']:visible").first
                if not adv_input.is_visible():
                    adv_input = page.locator("#advName:visible").first
                adv_input.click()
                adv_input.fill("")
                adv_input.fill(advocate_name)
                page.wait_for_timeout(1000)

                page.locator("#getData:visible").first.click()
                print(f"[*] Waiting for search results for {advocate_name}...")
                page.wait_for_timeout(6000)

                # Check if print button appeared (only appears if cases exist)
                print_btn = None
                for btn in page.locator("input[type='button']:visible, button:visible").all():
                    val = (btn.get_attribute("value") or "").lower()
                    txt = (btn.inner_text() or "").lower()
                    if btn.get_attribute("id") == "getData":
                        continue
                    if "print" in val or "print" in txt:
                        print_btn = btn
                        break

                if print_btn:
                    temp_pdf = f"causelist_{idx}.pdf"
                    target_page = page
                    try:
                        with context.expect_page(timeout=5000) as popup_info:
                            print_btn.click()
                        target_page = popup_info.value
                        target_page.wait_for_load_state("domcontentloaded")
                        target_page.wait_for_timeout(2000)
                        target_page.pdf(path=temp_pdf, format="A4", print_background=True)
                        target_page.close()
                    except Exception:
                        page.pdf(path=temp_pdf, format="A4", print_background=True)

                    print(f"[+] Downloaded cause list PDF for {advocate_name}")
                    parsed_rows = parse_pdf_with_gemini(temp_pdf)
                    for r in parsed_rows:
                        c_num = r.case_number.strip().upper()
                        if c_num and c_num not in seen_case_numbers:
                            seen_case_numbers.add(c_num)
                            all_cases.append(r)
                        elif not c_num:
                            all_cases.append(r)
                else:
                    print(f"[-] No listings found on board for {advocate_name}.")

            browser.close()

        print(f"\n[+] Total unique cases compiled: {len(all_cases)}")
        write_to_excel(all_cases, "cause_list.xlsx")
        generate_pdf(all_cases, date_str, "cause_list.pdf")
        print("[+] Done! Both cause_list.xlsx and cause_list.pdf generated successfully.")

    except Exception as err:
        print(f"[FATAL] Process aborted: {err}", file=sys.stderr)
        traceback.print_exc()
        sys.exit(1)
