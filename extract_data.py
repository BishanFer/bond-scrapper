#!/usr/bin/env python3
"""
Sri Lanka Treasury Bond Yield Scraper
Extracts daily bond yields from treasury.gov.lk and creates formatted Excel report.
"""

import os
import sys
import subprocess
import json
import re
from datetime import datetime, timedelta
from collections import defaultdict
from urllib.request import urlopen
from urllib.error import URLError

try:
    import xlrd
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "xlrd", "-q"])
    import xlrd

try:
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "-q"])
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter


TREASURY_URL = "https://www.treasury.gov.lk/web/report-daily-report/section/{year}"
REPORTS_DIR = os.environ.get("REPORTS_DIR", "treasury_reports")
OUTPUT_FILE = os.environ.get("OUTPUT_FILE", "treasury_bond_yields.xlsx")


def excel_date_to_datetime(excel_date):
    """Convert Excel date serial to datetime"""
    if isinstance(excel_date, float):
        return datetime(1899, 12, 30) + timedelta(days=excel_date)
    return None


def parse_bond_number(bond_str):
    """Clean bond number string"""
    if not bond_str:
        return None
    bond_str = str(bond_str).strip().replace('%', '')
    return bond_str


def get_report_urls(year=2025):
    """Scrape treasury.gov.lk to get Daily Summary Report URLs"""
    print(f"Fetching report list from treasury.gov.lk for {year}...")
    
    url = TREASURY_URL.format(year=year)
    try:
        response = urlopen(url, timeout=30)
        html = response.read().decode('utf-8')
    except URLError as e:
        print(f"Error fetching URL: {e}")
        return {}
    
    # Pattern to find Daily Summary Report links
    # Format: /api/file/{uuid}
    pattern = r'/api/file/([a-f0-9-]+)[^>]*>.*?(\d{2}\.\d{2}\.\d{4}).*?Daily Summary Report'
    
    # Alternative pattern for the HTML structure we observed
    alt_pattern = r'>(\d{2}\.\d{2}\.\d{4})</.*?href="(https://www\.treasury\.gov\.lk/api/file/[a-f0-9-]+)"[^>]*>Daily Summary Report'
    
    reports = {}
    
    # Try to extract dates and URLs
    date_url_pattern = r'(\d{2})\.(\d{2})\.(\d{4}).*?api/file/([a-f0-9-]+).*?Daily Summary Report'
    matches = re.findall(date_url_pattern, html, re.DOTALL)
    
    for match in matches:
        day, month, year, uuid = match
        date_str = f"{day}.{month}.{year}"
        url = f"https://www.treasury.gov.lk/api/file/{uuid}"
        
        # Parse date
        try:
            report_date = datetime.strptime(date_str, '%d.%m.%Y')
            reports[report_date] = url
        except ValueError:
            continue
    
    print(f"Found {len(reports)} reports for {year}")
    return reports


def download_report(url, date_str, output_dir):
    """Download a single report Excel file"""
    filename = f"report_{date_str}.xlsx"
    filepath = os.path.join(output_dir, filename)
    
    # Skip if already downloaded
    if os.path.exists(filepath) and os.path.getsize(filepath) > 100000:
        return filepath
    
    print(f"Downloading: {filename}")
    try:
        subprocess.run(
            ["curl", "-L", "-o", filepath, url],
            capture_output=True,
            check=True,
            timeout=60
        )
        return filepath if os.path.exists(filepath) else None
    except subprocess.CalledProcessError as e:
        print(f"Error downloading {url}: {e}")
        return None


def extract_data_from_report(filepath, date):
    """Extract QuotesTBond data from a single report"""
    data_twoway = {}
    data_ddo = {}
    
    try:
        wb = xlrd.open_workbook(filepath, ragged_rows=True)
        ws = wb.sheet_by_name("QuotesTBond")
        
        # TWO WAY QUOTES section (rows 8-72 approximately)
        for row_idx in range(8, 73):
            try:
                bond_number = ws.cell_value(row_idx, 2)
                maturity_date_val = ws.cell_value(row_idx, 4)
                buy_yield = ws.cell_value(row_idx, 7)
                
                if not bond_number:
                    continue
                    
                bond_number = parse_bond_number(bond_number)
                maturity_date = excel_date_to_datetime(maturity_date_val)
                
                if maturity_date and bond_number and buy_yield and buy_yield != 0:
                    key = (bond_number, maturity_date)
                    data_twoway[key] = buy_yield
            except:
                pass
        
        # DDO section (rows 75 onwards)
        for row_idx in range(75, ws.nrows):
            try:
                bond_number = ws.cell_value(row_idx, 2)
                maturity_date_val = ws.cell_value(row_idx, 4)
                buy_yield = ws.cell_value(row_idx, 7)
                
                if not bond_number:
                    continue
                    
                bond_number = parse_bond_number(bond_number)
                maturity_date = excel_date_to_datetime(maturity_date_val)
                
                if maturity_date and bond_number and buy_yield and buy_yield != 0:
                    key = (bond_number, maturity_date)
                    data_ddo[key] = buy_yield
            except:
                pass
                
    except Exception as e:
        print(f"Error reading {filepath}: {e}")
    
    return data_twoway, data_ddo


def create_excel(data_twoway, data_ddo, all_dates, output_path):
    """Create formatted Excel output"""
    wb_out = Workbook()
    ws_twoway = wb_out.active
    ws_twoway.title = "TWO_WAY_QUOTES"
    ws_ddo = wb_out.create_sheet("DDO_EDR_BONDS")
    
    # Styling
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    bond_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    sorted_dates = sorted(all_dates)
    date_formats = [d.strftime('%d-%b-%y') for d in sorted_dates]
    
    # Write header
    for ws, data_dict in [(ws_twoway, data_twoway), (ws_ddo, data_ddo)]:
        ws.cell(1, 1, "Bond Number")
        ws.cell(1, 2, "Maturity Date")
        
        # Style header cells
        for col in [1, 2]:
            ws.cell(1, col).fill = header_fill
            ws.cell(1, col).font = header_font
            ws.cell(1, col).border = thin_border
            ws.cell(1, col).alignment = Alignment(horizontal='center')
        
        # Date columns
        for col, date_str in enumerate(date_formats, 3):
            cell = ws.cell(1, col, date_str)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')
            cell.border = thin_border
        
        # Get all unique bonds
        all_bonds = set(data_dict.keys())
        sorted_bonds = sorted(all_bonds, key=lambda x: x[1])
        
        # Write data
        for row_idx, (bond_number, maturity_date) in enumerate(sorted_bonds, 2):
            ws.cell(row_idx, 1, bond_number)
            ws.cell(row_idx, 2, maturity_date.strftime('%d-%b-%Y'))
            
            # Style bond/maturity cells
            ws.cell(row_idx, 1).fill = bond_fill
            ws.cell(row_idx, 2).fill = bond_fill
            
            for col in range(1, len(sorted_dates) + 3):
                ws.cell(row_idx, col).border = thin_border
            
            # Get yields for this bond
            yields_dict = data_dict.get(bond_number, {})
            
            for col, date in enumerate(sorted_dates, 3):
                yield_val = yields_dict.get(date)
                if yield_val is not None and yield_val != 0:
                    ws.cell(row_idx, col, yield_val)
                    ws.cell(row_idx, col).number_format = '0.00%'
                    ws.cell(row_idx, col).alignment = Alignment(horizontal='right')
        
        # Column widths
        ws.column_dimensions['A'].width = 16
        ws.column_dimensions['B'].width = 14
        for col in range(3, len(sorted_dates) + 3):
            ws.column_dimensions[get_column_letter(col)].width = 12
    
    wb_out.save(output_path)
    print(f"Saved: {output_path}")
    return wb_out


def update_existing_excel(existing_path, new_data_twoway, new_data_ddo, new_date):
    """Add new date column to existing Excel"""
    # For now, recreate from scratch
    # TODO: Implement incremental update for efficiency
    pass


def run_full_export(year=2025, month=None):
    """Run full export for a given year/month"""
    os.makedirs(REPORTS_DIR, exist_ok=True)
    
    # Get all report URLs
    reports = get_report_urls(year)
    
    if not reports:
        print("No reports found!")
        return
    
    # Filter by month if specified
    if month:
        reports = {k: v for k, v in reports.items() if k.month == month}
        print(f"Filtered to {len(reports)} reports for month {month}")
    
    # Download all reports
    for date, url in sorted(reports.items()):
        date_str = date.strftime('%d-%m-%Y')
        download_report(url, date_str, REPORTS_DIR)
    
    # Process all reports
    data_twoway = defaultdict(dict)
    data_ddo = defaultdict(dict)
    all_dates = []
    
    report_files = sorted([f for f in os.listdir(REPORTS_DIR) if f.endswith('.xlsx')])
    
    for report_file in report_files:
        date_str = report_file.replace('report_', '').replace('.xlsx', '')
        try:
            trade_date = datetime.strptime(date_str, '%d-%m-%Y')
        except ValueError:
            continue
        
        filepath = os.path.join(REPORTS_DIR, report_file)
        
        # Skip if file too small (download failed)
        if os.path.getsize(filepath) < 100000:
            print(f"Skipping incomplete file: {report_file}")
            continue
        
        all_dates.append(trade_date)
        
        twoday, ddo = extract_data_from_report(filepath, trade_date)
        
        for key, yield_val in twoday.items():
            data_twoway[key][trade_date] = yield_val
        
        for key, yield_val in ddo.items():
            data_ddo[key][trade_date] = yield_val
    
    # Create Excel
    create_excel(dict(data_twoway), dict(data_ddo), all_dates, OUTPUT_FILE)
    
    print(f"\nComplete!")
    print(f"Total dates: {len(set(all_dates))}")
    print(f"TWO WAY bonds: {len(data_twoway)}")
    print(f"DDO bonds: {len(data_ddo)}")


def run_today_update():
    """Fetch only today's report and update Excel"""
    today = datetime.now()
    year = today.year
    
    # Get today's date string as used on treasury.gov.lk
    date_str = today.strftime('%d.%m.%Y')
    
    # Get report URLs
    reports = get_report_urls(year)
    
    # Find today's report
    today_report = None
    for date, url in reports.items():
        if date.strftime('%d.%m.%Y') == date_str:
            today_report = (date, url)
            break
    
    if not today_report:
        print(f"No report found for today ({date_str})")
        # Try yesterday
        yesterday = today - timedelta(days=1)
        date_str_yest = yesterday.strftime('%d.%m.%Y')
        for date, url in reports.items():
            if date.strftime('%d.%m.%Y') == date_str_yest:
                today_report = (date, url)
                print(f"Using yesterday's report: {date_str_yest}")
                break
    
    if not today_report:
        print("No recent report found")
        return
    
    date, url = today_report
    date_fmt = date.strftime('%d-%m-%Y')
    
    # Download
    filepath = download_report(url, date_fmt, REPORTS_DIR)
    if not filepath:
        print("Failed to download report")
        return
    
    # Extract data
    twoday, ddo = extract_data_from_report(filepath, date)
    
    # For now, just show what we got
    print(f"Extracted from {date_fmt}:")
    print(f"  TWO WAY: {len(twoday)} bonds")
    print(f"  DDO: {len(ddo)} bonds")
    
    return twoday, ddo, date


if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="Sri Lanka Treasury Bond Yield Scraper")
    parser.add_argument("--year", type=int, default=2025, help="Year to scrape")
    parser.add_argument("--month", type=int, help="Month to filter (1-12)")
    parser.add_argument("--today", action="store_true", help="Only fetch today's report")
    
    args = parser.parse_args()
    
    if args.today:
        run_today_update()
    else:
        run_full_export(args.year, args.month)