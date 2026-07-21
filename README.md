# Workflow-Automation-using-HTML-Report-Parser-to-Excel-
## HTML Report Parser to Excel
This project automates the process of extracting structured data from a ZIP archive of HTML files and exporting the results into an Excel spreadsheet.  The repository recreates a real-world workflow using synthetic sample files. No employer, customer, or proprietary data is included.
## Overview
Manually reviewing downloaded HTML pages and copying values into a spreadsheet is repetitive and time-consuming. I recently was tasked with this and this project demonstrates how I automated that workflow by:
- Reading a ZIP archive containing multiple HTML files
- Using Python to parse specific fields from each page
- Normalizing the extracted data
- Exporting the results into a structured Excel file
## Development Notes
The parser was initially generated with assistance from Perplexity.ai and then manually reviewed, verified, and tested against sample data to confirm tge code is functional with no critical bugs, and reusable for future use.

This project is an example of AI-augmented development with human validation and review. The final implementation, testing, and documentation were reviewed and confirmed before being added to this repository
## About the "Acuity_lab/" Directory
The Acuity_lab directory created for this project contains SYNTHETIC HTML examples that simulate exports from a scheduling workflow similar to the original real-world use-case. All data and page content in this directory and project is fabricated for lab and portfolio purposes only, and does NOT include any real customer, organizational, or propreitary information to protect the Confidentiality and Integrity of data.
## Features
- Parses multiple HTML files from a single ZIP archive
- Extracts selected fields such as title, date, status, and table values
- Writes structured output into Excel
- Uses sample input files to safely demonstrate the workflow
- Can be adapted to similar reporting or log-processing tasks
## Sample Workflow
- Place the sample ZIP archive in the Acutiy_lab/ folder
- Run the parser script
- The script reads each HTML file in the archive
- Selected fields are extracted and written into an Excel spreadsheet.
- The finished spreadsheet is saved in the output/ folder
## Requirements
- Python 3.10+
- pip
## Dependencies
```text
pip install -r requirements.txt
```
## Usage
```text
python src/parse_html_zip.py
```
## Python Script Used
```text
from __future__ import annotations

import sys
import zipfile
from pathlib import Path
from html.parser import HTMLParser
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

HEADERS = [
    'Row Number',
    'Purchase ID',
    'Coupon Code',
    'Assigned Package Holder',
    'Date Redeemed',
    'Initial Balance',
    'Current Balance',
    'Uses Redeemed',
    'Status',
    'Source HTML File',
]

HDR_FILL = PatternFill('solid', fgColor='0F6D73')
HDR_FONT = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
BODY_FONT = Font(name='Calibri', size=11, color='000000')
TITLE_FONT = Font(name='Calibri', size=14, bold=True, color='000000')
CENTER = Alignment(horizontal='center', vertical='center')
LEFT = Alignment(horizontal='left', vertical='center')
RIGHT = Alignment(horizontal='right', vertical='center')
CUR_FMT = '$#,##0'
INT_FMT = '#,##0'


class ReportTableParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.in_table = False
        self.in_tr = False
        self.in_td = False
        self.current_cell = []
        self.current_row = []
        self.rows = []

    def handle_starttag(self, tag, attrs):
        if tag == 'table':
            self.in_table = True
        elif tag == 'tr' and self.in_table:
            self.in_tr = True
            self.current_row = []
        elif tag in ('td', 'th') and self.in_tr:
            self.in_td = True
            self.current_cell = []

    def handle_endtag(self, tag):
        if tag in ('td', 'th') and self.in_td:
            text = ''.join(self.current_cell).strip()
            self.current_row.append(' '.join(text.split()))
            self.in_td = False
        elif tag == 'tr' and self.in_tr:
            if len(self.current_row) == 9 and self.current_row[0].isdigit():
                self.rows.append(self.current_row)
            self.current_row = []
            self.in_tr = False
        elif tag == 'table':
            self.in_table = False

    def handle_data(self, data):
        if self.in_td:
            self.current_cell.append(data)


def money_to_int(value: str) -> int:
    return int(value.replace('$', '').replace(',', '').strip())


def parse_html_text(html_text: str, source_name: str) -> list[list]:
    parser = ReportTableParser()
    parser.feed(html_text)
    parsed = []
    for row in parser.rows:
        parsed.append([
            int(row[0]),
            row[1],
            row[2],
            row[3],
            row[4],
            money_to_int(row[5]),
            money_to_int(row[6]),
            int(row[7]),
            row[8],
            source_name,
        ])
    return parsed


def autosize(ws):
    widths = {1: 12, 2: 18, 3: 16, 4: 24, 5: 14, 6: 14, 7: 14, 8: 14, 9: 14, 10: 28}
    for col_idx, width in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def build_workbook(rows: list[list], zip_name: str, out_path: Path):
    wb = Workbook()
    summary = wb.active
    summary.title = 'Summary'
    data_ws = wb.create_sheet('ParsedData')

    summary['B2'] = 'Acuity HTML ZIP Parsed Export'
    summary['B2'].font = TITLE_FONT
    summary['B3'] = f'Source ZIP: {zip_name}'
    summary['B4'] = 'Purpose: Parse multi-page HTML reports and export to XLSX.'
    summary['B6'] = 'Total records'
    summary['C6'] = len(rows)
    summary['B7'] = 'Unique HTML files'
    summary['C7'] = len(sorted({r[9] for r in rows}))
    summary['B8'] = 'Initial balance total'
    summary['C8'] = '=SUM(ParsedData[Initial Balance])'
    summary['B9'] = 'Current balance total'
    summary['C9'] = '=SUM(ParsedData[Current Balance])'
    summary['B10'] = 'Total uses redeemed'
    summary['C10'] = '=SUM(ParsedData[Uses Redeemed])'

    for cell in ['B6', 'B7', 'B8', 'B9', 'B10']:
        summary[cell].font = Font(name='Calibri', size=11, bold=True)

    for cell in ['C8', 'C9']:
        summary[cell].number_format = CUR_FMT

    summary.column_dimensions['A'].width = 3
    summary.column_dimensions['B'].width = 24
    summary.column_dimensions['C'].width = 18

    data_ws.append(HEADERS)
    for row in rows:
        data_ws.append(row)

    for cell in data_ws[1]:
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = CENTER

    for row in data_ws.iter_rows(min_row=2, max_row=data_ws.max_row):
        row[0].font = BODY_FONT
        row[0].alignment = CENTER
        row[0].number_format = INT_FMT
        row[1].font = BODY_FONT
        row[1].alignment = LEFT
        row[2].font = BODY_FONT
        row[2].alignment = CENTER
        row[3].font = BODY_FONT
        row[3].alignment = LEFT
        row[4].font = BODY_FONT
        row[4].alignment = CENTER
        row[5].font = BODY_FONT
        row[5].alignment = RIGHT
        row[5].number_format = CUR_FMT
        row[6].font = BODY_FONT
        row[6].alignment = RIGHT
        row[6].number_format = CUR_FMT
        row[7].font = BODY_FONT
        row[7].alignment = CENTER
        row[7].number_format = INT_FMT
        row[8].font = BODY_FONT
        row[8].alignment = CENTER
        row[9].font = BODY_FONT
        row[9].alignment = LEFT

    table = Table(displayName='ParsedData', ref=f'A1:J{data_ws.max_row}')
    table.tableStyleInfo = TableStyleInfo(
        name='TableStyleMedium2',
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False
    )
    data_ws.add_table(table)
    data_ws.freeze_panes = 'A2'
    autosize(data_ws)
    wb.save(out_path)


def main():
    if len(sys.argv) < 2:
        print('Usage: python parse_acuity_html_zip_to_xlsx.py <input_zip> [output_xlsx]')
        raise SystemExit(1)

    input_zip = Path(sys.argv[1]).expanduser().resolve()
    output_xlsx = Path(sys.argv[2]).expanduser().resolve() if len(sys.argv) > 2 else input_zip.with_name(input_zip.stem + '_parsed.xlsx')

    if not input_zip.exists():
        print(f'Input ZIP not found: {input_zip}')
        raise SystemExit(1)

    rows = []
    with zipfile.ZipFile(input_zip, 'r') as zf:
        html_files = sorted([name for name in zf.namelist() if name.lower().endswith('.html')])

        for name in html_files:
            if 'index' in Path(name).name.lower():
                continue
            html_text = zf.read(name).decode('utf-8', errors='ignore')
            rows.extend(parse_html_text(html_text, Path(name).name))

    if not rows:
        print('No parseable report rows found in ZIP.')
        raise SystemExit(1)

    build_workbook(rows, input_zip.name, output_xlsx)
    print(f'Parsed {len(rows)} rows into: {output_xlsx}')


if __name__ == '__main__':
    main()
```
## Example Fields Extracted
Depending on the HTML structure, the script can extract fields such as:
- Report name
- Page title
- Generated date
- Status
- Record ID
- Table-based values
- Notes or summary text
## Skills Demonstrated
This project highlights practical skills in:
- Python automation
- File handling and ZIP processing
- HTML parsing
- Data transformation
- AI-assisted development with manual verification and testing
## Future Improvments
- Add command-line arguments for flexible input and output paths
- Add logging and error handling
- Add unit tests for parsing logic
- Support CSV outputin addition to Excel
- Package the script for easier reuse
## License
This project is shared for portfolio educational purposes
