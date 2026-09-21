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

class CauseListRow(BaseModel):
    sl_no: str = Field(description="First SL NO column")
    case_number: str = Field(description="CASE NUMBER (e.g. WP NO 102709/2026)")
    case_name: str = Field(description="Full raw case name string")
    ch: str = Field(description="Court Hall number under CH")
    list_num: str = Field(description="List number under LIST")
    list_sl_no: str = Field(description="Item number under second SL NO")
    status: str = Field(description="STATUS column (e.g. ORDERS, HEARING - IA)")
    judges: str = Field(description="JUDGES column")

class CauseListDocument(BaseModel):
    date: str = Field(description="Date displayed at top of cause list")
    rows: list[CauseListRow]

def get_target_date_ist():
    manual = os.environ.get("TEST_DATE")
    if manual and manual.strip():
        return manual.strip()
    return "22/09/2026"  # Revert to next_day.strftime("%d/%m/%Y") for live runs

def parse_case_parties(raw_name: str, advocate_keyword="CHOWDHARY"):
    """
    Cleans raw cause list text into '[Petitioner] vs [Respondent]'
    and detects which side Mahesh Chowdhary represents.
    """
    parts = re.split(r'\s+V/?S\s+', raw_name, maxsplit=1, flags=re.IGNORECASE)
    if len(parts) != 2:
        return raw_name.strip(), "", "PETITIONER"

    pet_raw, res_raw = parts[0], parts[1]

    # Detect which party Mahesh Chowdhary represents
    pet_has_adv = bool(re.search(r'\b(MAHESH|CHOWDHARY)\b', pet_raw, re.I))
    res_has_adv = bool(re.search(r'\b(MAHESH|CHOWDHARY)\b', res_raw, re.I))

    if res_has_adv and not pet_has_adv:
        client_side = "RESPONDENT"
    else:
        client_side = "PETITIONER"

    def clean_side(text: str) -> str:
        # Strip prefixes like PET:, RES:, etc.
        t = re.sub(r'^(PET|RES|PETITIONER|RESPONDENT|APPELLANT|COMPLAINANT)\s*:\s*', '', text, flags=re.I)
        # Strip notes in parentheses like (MA NOT FILED), (VK NOT FILED)
        t = re.sub(r'\((MA|VK|V/K|VAKALATH|MEMO|NOC|POWER)[^\)]*\)', '', t, flags=re.I)
        
        chunks = [c.strip() for c in re.split(r'[,;]', t) if c.strip()]
        core_parties = []
        for chunk in chunks:
            # Drop tokens that represent advocates, government pleaders, or appearances
            if not re.search(r'\b(MAHESH|CHOWDHARY|AGA|HCGP|ADV|ADVOCATE|ADVOCATES|COUNSEL|FOR\s+RES|FOR\s+PET|FOR\s+R\d+|FOR\s+P\d+|GOVT\s+PLEADER)\b', chunk, re.I):
                core_parties.append(chunk)

        if core_parties:
            return ", ".join(core_parties).strip()
        return chunks[0] if chunks else text.strip()

    pet_clean = clean_side(pet_raw)
    res_clean = clean_side(res_raw)

    return pet_clean, res_clean, client_side

def extract_rows_from_page(page) -> list[list[str]]:
    print("[*] Extracting cause list table directly from page DOM...")
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
                
    print(f"[+] Direct DOM extraction found {len(extracted_rows)} case rows.")
    return extracted_rows

def fetch_cause_list_data():
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
        "args": ["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage", "--ignore-certificate-errors"]
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
        page.add_init_script("window.print = () => { console.log('print intercepted'); };")

        print("[*] Connecting to High Court Portal via Indian Gateway...")
        page.goto("https://judiciary.karnataka.gov.in/causelistSearch.php", wait_until="domcontentloaded", timeout=90000)
        page.wait_for_timeout(3000)

        # 1. Bench Selection
        bench_select = page.locator("select[name='bench']:visible").first
        for opt in bench_select.locator("option").all():
            opt_text = opt.inner_text().strip()
            if "Bengaluru" in opt_text or "Bangalore" in opt_text:
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
        adv_input.fill("Mahesh Chowdhary")
        page.wait_for_timeout(1000)

        # 5. Click GET DETAILS
        page.locator("#getData:visible").first.click()
        print("[*] Waiting 7 seconds for court database query...")
        page.wait_for_timeout(7000)

        # Extract table rows
        raw_table_rows = extract_rows_from_page(page)

        # Also capture raw site PDF as a secondary record
        print_btn = page.locator("input[value*='print list' i]:visible, button:has-text('print' i):visible").first
        if print_btn.is_visible():
            try:
                with context.expect_page(timeout=8000) as popup_info:
                    print_btn.click()
                print_page = popup_info.value
                print_page.wait_for_load_state("domcontentloaded")
                print_page.pdf(path="causelist.pdf", format="A4", print_background=True)
            except Exception:
                page.pdf(path="causelist.pdf", format="A4", print_background=True)
        else:
            page.pdf(path="causelist.pdf", format="A4", print_background=True)

        browser.close()

    return date_str, raw_table_rows

def write_to_excel(date_str: str, raw_rows: list, excel_path="cause_list.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cause List"
    ws.views.sheetView[0].showGridLines = True

    # 1. Sheet Title Banner
    ws.merge_cells("A1:H1")
    title_cell = ws["A1"]
    title_cell.value = f"HIGH COURT OF KARNATAKA (BENGALURU) - CAUSE LIST FOR {date_str} | ADVOCATE: MAHESH CHOWDHARY"
    title_cell.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    # 2. Table Headers
    headers = ["SL NO", "CASE NUMBER", "CASE NAME", "CH", "LIST", "SL NO", "STATUS", "JUDGES"]
    ws.append(headers)
    ws.row_dimensions[2].height = 26

    header_fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )

    for col_idx in range(1, 9):
        cell = ws.cell(row=2, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    # Rich Text Fonts for Case Name
    bold_client_font = InlineFont(rFont="Calibri", sz=10, b=True, color="001E40AF") # Bold Navy Blue
    regular_font = InlineFont(rFont="Calibri", sz=10, color="00334155")
    vs_font = InlineFont(rFont="Calibri", sz=9, i=True, color="0064748B")

    # 3. Insert and Format Data Rows
    row_idx = 3
    for seq, r in enumerate(raw_rows, start=1):
        pet, res, client_side = parse_case_parties(r[2])

        # Rich text with Mahesh Chowdhary's client highlighted
        if client_side == "PETITIONER":
            case_name_rich = CellRichText(
                TextBlock(bold_client_font, f"★ {pet} (In-Charge)\n"),
                TextBlock(vs_font, "vs\n"),
                TextBlock(regular_font, res)
            )
        else:
            case_name_rich = CellRichText(
                TextBlock(regular_font, f"{pet}\n"),
                TextBlock(vs_font, "vs\n"),
                TextBlock(bold_client_font, f"★ {res} (In-Charge)")
            )

        row_data = [
            seq,           # Requirement 2: Clean 1, 2, 3...
            r[1].strip(),  # CASE NUMBER
            case_name_rich,# Requirement 3: Cleaned & Highlighted
            r[3].strip(),  # CH
            r[4].strip(),  # LIST
            r[5].strip(),  # BOARD ITEM SL NO
            r[6].strip(),  # STATUS
            r[7].strip()   # JUDGES
        ]
        ws.append(row_data)

        # Style data row
        is_even = (seq % 2 == 0)
        row_bg = "F8FAFC" if is_even else "FFFFFF"
        cell_fill = PatternFill(start_color=row_bg, end_color=row_bg, fill_type="solid")

        for c in range(1, 9):
            cell = ws.cell(row=row_idx, column=c)
            cell.fill = cell_fill
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
                    cell.font = Font(name="Calibri", size=10)

        # Requirement 1: Prevent cell collapse by providing explicit row height
        ws.row_dimensions[row_idx].height = 58
        row_idx += 1

    # Set generous column widths
    col_widths = {1: 10, 2: 24, 3: 48, 4: 8, 5: 8, 6: 10, 7: 28, 8: 36}
    for col_idx, width in col_widths.items():
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = width

    wb.save(excel_path)
    print(f"[+] Formatted Excel saved to: {excel_path}")

def generate_styled_pdf(date_str: str, raw_rows: list, pdf_path="cause_list_summary.pdf"):
    """Renders a landscape A4 cause list PDF with high visual clarity."""
    print("[*] Generating formatted PDF cause list...")

    rows_html = ""
    for seq, r in enumerate(raw_rows, start=1):
        pet, res, client_side = parse_case_parties(r[2])

        if client_side == "PETITIONER":
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

        status_class = "status-ia" if "IA" in r[6].upper() else ("status-orders" if "ORDERS" in r[6].upper() else "status-general")

        rows_html += f"""
        <tr>
            <td class="text-center font-bold">{seq}</td>
            <td class="text-center case-num">{r[1]}</td>
            <td>{case_name_cell}</td>
            <td class="text-center ch-cell">{r[3]}</td>
            <td class="text-center">{r[4]}</td>
            <td class="text-center font-bold">{r[5]}</td>
            <td><span class="status-badge {status_class}">{r[6]}</span></td>
            <td class="judges-text">{r[7]}</td>
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
                margin: 8mm 10mm 8mm 10mm;
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
                margin-bottom: 10px;
            }}
            .header h1 {{
                margin: 0;
                font-size: 16px;
                text-transform: uppercase;
                letter-spacing: 0.5px;
                color: #0f172a;
            }}
            .header-meta {{
                font-size: 11px;
                color: #475569;
                font-weight: 500;
            }}
            .summary-bar {{
                background: #f1f5f9;
                padding: 6px 12px;
                border-radius: 4px;
                margin-bottom: 10px;
                display: flex;
                gap: 20px;
                font-weight: 600;
                font-size: 11px;
                color: #334155;
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
                padding: 8px 5px;
                letter-spacing: 0.5px;
                border: 1px solid #334155;
            }}
            td {{
                padding: 6px 5px;
                border: 1px solid #cbd5e1;
                vertical-align: middle;
                word-wrap: break-word;
            }}
            tr:nth-child(even) {{
                background-color: #f8fafc;
            }}
            .text-center {{ text-align: center; }}
            .font-bold {{ font-weight: 700; }}
            .case-num {{
                font-weight: 700;
                color: #0f172a;
                font-size: 11px;
            }}
            .ch-cell {{
                font-size: 13px;
                font-weight: 800;
                color: #0284c7;
            }}
            .party-box {{
                line-height: 1.3;
            }}
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
                vertical-align: middle;
            }}
            .opp-box {{
                color: #475569;
                padding: 1px 6px;
            }}
            .vs-text {{
                font-size: 9px;
                font-style: italic;
                color: #94a3b8;
                margin: 2px 0 2px 8px;
            }}
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
            .judges-text {{
                font-size: 10px;
                line-height: 1.25;
                color: #1e293b;
            }}
        </style>
    </head>
    <body>
        <div class="header">
            <div>
                <h1>High Court of Karnataka &mdash; Bengaluru Bench</h1>
                <div class="header-meta">Daily Cause List &bull; <strong>Date: {date_str}</strong></div>
            </div>
            <div style="text-align: right;">
                <div style="font-size: 13px; font-weight: 700; color: #0f172a;">ADV. MAHESH CHOWDHARY</div>
                <div class="header-meta">Generated: {datetime.now(pytz.timezone('Asia/Kolkata')).strftime('%d-%b-%Y %I:%M %p')} IST</div>
            </div>
        </div>

        <div class="summary-bar">
            <div>Total Matters Listed: <strong>{len(raw_rows)}</strong></div>
            <div>Bench: <strong>Principal Bench (Bengaluru)</strong></div>
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

    print(f"[+] Custom formatted PDF generated at: {pdf_path}")

if __name__ == "__main__":
    try:
        date_str, rows = fetch_cause_list_data()
        if not rows:
            print("[!] No records found for this date. Generating empty templates.")
            rows = []

        # Generate both polished formats
        write_to_excel(date_str, rows, "cause_list.xlsx")
        generate_styled_pdf(date_str, rows, "cause_list_summary.pdf")

    except Exception as err:
        print(f"[FATAL] Process aborted: {err}", file=sys.stderr)
        traceback.print_exc()
        sys.exit(1)
