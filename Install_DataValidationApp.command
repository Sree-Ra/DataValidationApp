#!/bin/bash

echo "Installing required packages..."
pip3 install pandas openpyxl pillow requests --quiet --break-system-packages 2>/dev/null || pip3 install pandas openpyxl pillow requests --quiet

APP_DIR="$HOME/DataValidationApp"
APP_FILE="$APP_DIR/app.py"
mkdir -p "$APP_DIR"

cat > "$APP_FILE" << 'ENDOFPYTHON'
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import os
import subprocess
import re
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# ─────────────────────────────────────────────
#  MACHINIFY COLOR PALETTE
# ─────────────────────────────────────────────
BG_PAGE     = '#FFFFFF'
BG_NAVBAR   = '#0B0B0F'
BG_CARD     = '#F7F8FC'
BG_INPUT    = '#FFFFFF'
ACCENT_1    = '#6C3FF5'
BTN_HOVER   = '#5533CC'
TEXT_DARK   = '#0B0B0F'
TEXT_MID    = '#4A4A68'
TEXT_LIGHT  = '#9898B0'
BORDER_CLR  = '#E0E0EF'

# ─────────────────────────────────────────────
#  VALID US STATES
# ─────────────────────────────────────────────
VALID_US_STATES = {
    'AL','AK','AZ','AR','CA','CO','CT','DE','FL','GA',
    'HI','ID','IL','IN','IA','KS','KY','LA','ME','MD',
    'MA','MI','MN','MS','MO','MT','NE','NV','NH','NJ',
    'NM','NY','NC','ND','OH','OK','OR','PA','RI','SC',
    'SD','TN','TX','UT','VT','VA','WA','WV','WI','WY',
    'DC','PR','GU','VI','AS','MP',
    'ALABAMA','ALASKA','ARIZONA','ARKANSAS','CALIFORNIA',
    'COLORADO','CONNECTICUT','DELAWARE','FLORIDA','GEORGIA',
    'HAWAII','IDAHO','ILLINOIS','INDIANA','IOWA','KANSAS',
    'KENTUCKY','LOUISIANA','MAINE','MARYLAND','MASSACHUSETTS',
    'MICHIGAN','MINNESOTA','MISSISSIPPI','MISSOURI','MONTANA',
    'NEBRASKA','NEVADA','NEW HAMPSHIRE','NEW JERSEY','NEW MEXICO',
    'NEW YORK','NORTH CAROLINA','NORTH DAKOTA','OHIO','OKLAHOMA',
    'OREGON','PENNSYLVANIA','RHODE ISLAND','SOUTH CAROLINA',
    'SOUTH DAKOTA','TENNESSEE','TEXAS','UTAH','VERMONT',
    'VIRGINIA','WASHINGTON','WEST VIRGINIA','WISCONSIN','WYOMING'
}

VALID_US_CITIES = {
    'NEW YORK','LOS ANGELES','CHICAGO','HOUSTON','PHOENIX','PHILADELPHIA',
    'SAN ANTONIO','SAN DIEGO','DALLAS','SAN JOSE','AUSTIN','JACKSONVILLE',
    'FORT WORTH','COLUMBUS','CHARLOTTE','INDIANAPOLIS','SAN FRANCISCO',
    'SEATTLE','DENVER','NASHVILLE','OKLAHOMA CITY','EL PASO','WASHINGTON',
    'BOSTON','LAS VEGAS','MEMPHIS','LOUISVILLE','PORTLAND','BALTIMORE',
    'MILWAUKEE','ALBUQUERQUE','TUCSON','FRESNO','SACRAMENTO','MESA',
    'KANSAS CITY','ATLANTA','OMAHA','COLORADO SPRINGS','RALEIGH',
    'LONG BEACH','VIRGINIA BEACH','MINNEAPOLIS','TAMPA','NEW ORLEANS',
    'ARLINGTON','BAKERSFIELD','HONOLULU','ANAHEIM','AURORA','SANTA ANA',
    'CORPUS CHRISTI','RIVERSIDE','ST LOUIS','PITTSBURGH','LEXINGTON',
    'ANCHORAGE','STOCKTON','CINCINNATI','ST PAUL','TOLEDO','GREENSBORO',
    'NEWARK','PLANO','HENDERSON','LINCOLN','BUFFALO','FORT WAYNE',
    'JERSEY CITY','CHULA VISTA','ORLANDO','ST PETERSBURG','NORFOLK',
    'CHANDLER','LAREDO','MADISON','DURHAM','LUBBOCK','WINSTON SALEM',
    'GARLAND','GLENDALE','HIALEAH','RENO','BATON ROUGE','IRVINE',
    'CHESAPEAKE','SCOTTSDALE','NORTH LAS VEGAS','FREMONT','GILBERT',
    'SAN BERNARDINO','BIRMINGHAM','ROCHESTER','RICHMOND','SPOKANE',
    'DES MOINES','MONTGOMERY','LITTLE ROCK','SALT LAKE CITY','TALLAHASSEE',
    'HARTFORD','JACKSON','AUGUSTA','COLUMBIA','PROVIDENCE','CONCORD',
    'DOVER','HELENA','BOISE','PIERRE','BISMARCK','CHEYENNE','JUNEAU',
    'FRANKFORT','ANNAPOLIS','TRENTON','SANTA FE','ALBANY','TOPEKA',
    'OLYMPIA','SPRINGFIELD','LANSING','ST. LOUIS','ST. PAUL',
    'ST. PETERSBURG','WINSTON-SALEM'
}

INVALID_ZIPS = {
    '00000','99999','11111','22222','33333','44444',
    '55555','66666','77777','88888','12345'
}

def is_valid_zip(val):
    s = str(val).strip().split('-')[0]
    if not re.match(r'^\d{5}$', s):
        return False
    if s in INVALID_ZIPS:
        return False
    zip_int = int(s)
    valid_ranges = [
        (1001,2791),(2801,2940),(3031,3897),(3901,4992),(5001,5495),
        (5501,5544),(6001,6928),(7001,8989),(10001,14975),(15001,19640),
        (19701,19980),(20001,20599),(20601,21930),(22001,24658),(24701,26886),
        (27006,28909),(29001,29948),(30001,31999),(32004,34997),(35004,36925),
        (37010,38589),(38601,39776),(40003,42788),(43001,45999),(46001,47997),
        (48001,49971),(50001,52809),(53001,54990),(55001,56763),(57001,57799),
        (58001,58856),(59001,59937),(60001,62999),(63001,65899),(66002,67954),
        (68001,69367),(70001,71497),(71601,72959),(73001,74966),(75001,79999),
        (80001,81658),(82001,83128),(83201,83876),(84001,84784),(85001,86556),
        (87001,88441),(88901,89883),(90001,96162),(96701,96898),(97001,97920),
        (98001,99403),(99501,99950),
    ]
    for lo, hi in valid_ranges:
        if lo <= zip_int <= hi:
            return True
    return False

RESULTS_TAB = 'DataAnalysisChecksApp_Results'

def detect_col(columns, keywords):
    for kw in keywords:
        for col in columns:
            if kw.lower() in col.lower():
                return col
    return None

def looks_like_id_or_text_field(series):
    non_null = [str(v).strip() for v in series
                if pd.notna(v) and str(v).strip() != '']
    if not non_null:
        return False
    mixed = sum(1 for v in non_null
                if re.search(r'[A-Za-z]', v) and re.search(r'\d', v))
    return (mixed / len(non_null)) > 0.20

def looks_like_date_field(series):
    non_null = [str(v).strip() for v in series
                if pd.notna(v) and str(v).strip() != '']
    if not non_null:
        return False
    parsed = sum(1 for v in non_null[:50]
                 if _try_parse_date(v))
    return (parsed / len(non_null[:50])) > 0.6

def _try_parse_date(v):
    try:
        pd.to_datetime(v, dayfirst=False)
        return True
    except:
        return False

def get_data_sheet(filepath):
    xl = pd.ExcelFile(filepath)
    all_sheets = xl.sheet_names
    print(f"[INFO] All sheets: {all_sheets}")
    for name in all_sheets:
        if name.strip() != RESULTS_TAB:
            print(f"[INFO] Using sheet: '{name}'")
            return name
    return all_sheets[0]

# ─────────────────────────────────────────────
#  VALIDATION HELPERS
# ─────────────────────────────────────────────

def pct(count, total):
    return round((count / total) * 100, 2) if total > 0 else 0.0

def null_check(series):
    c = int(series.isna().sum())
    return pct(c, len(series)), c, ('NULL' if c > 0 else '')

def blank_check(series):
    c = int(series.apply(
        lambda x: (not pd.isna(x)) and str(x).strip() == ''
    ).sum())
    return pct(c, len(series)), c, ('BLANK (empty cell)' if c > 0 else '')

def invalid_date_check(series):
    cutoff = datetime(1920, 1, 1)
    bad = []
    for v in series:
        if pd.isna(v) or str(v).strip() == '':
            continue
        raw = str(v).strip()
        try:
            d = pd.to_datetime(raw, dayfirst=False)
            if d <= pd.Timestamp(cutoff):
                bad.append(raw)
        except:
            bad.append(raw)
    unique_bad = ', '.join(sorted(set(bad)))
    return pct(len(bad), len(series)), len(bad), unique_bad

def invalid_city_check(series):
    bad = []
    for v in series:
        if pd.isna(v) or str(v).strip() == '':
            continue
        raw   = str(v).strip()
        upper = raw.upper()
        if re.match(r'^\d+$', upper):
            bad.append(raw)
            continue
        norm = upper.replace('.','').replace('-',' ').replace('  ',' ')
        if norm not in VALID_US_CITIES and upper not in VALID_US_CITIES:
            bad.append(raw)
    return pct(len(bad), len(series)), len(bad), ', '.join(sorted(set(bad)))

def invalid_state_check(series):
    bad = []
    for v in series:
        if pd.isna(v) or str(v).strip() == '':
            continue
        raw = str(v).strip()
        if raw.upper() not in VALID_US_STATES:
            bad.append(raw)
    return pct(len(bad), len(series)), len(bad), ', '.join(sorted(set(bad)))

def invalid_zip_check(series):
    bad = []
    for v in series:
        if pd.isna(v) or str(v).strip() == '':
            continue
        raw = str(v).strip()
        if not is_valid_zip(raw):
            bad.append(raw)
    return pct(len(bad), len(series)), len(bad), ', '.join(sorted(set(bad)))

def invalid_numeric_check(series):
    bad = []
    for v in series:
        if pd.isna(v) or str(v).strip() == '':
            continue
        raw = str(v).strip()
        try:
            float(raw.replace(',', ''))
        except:
            bad.append(raw)
    return pct(len(bad), len(series)), len(bad), ', '.join(sorted(set(bad)))

def salesorder_amount_check(df, amount_col, qty_col, rate_col):
    bad     = []
    checked = 0
    for _, row in df.iterrows():
        try:
            amt      = float(str(row[amount_col]).replace(',','').strip())
            qty      = float(str(row[qty_col]).replace(',','').strip())
            rate     = float(str(row[rate_col]).replace(',','').strip())
            expected = round(qty * rate, 2)
            if abs(amt - expected) > max(0.02, abs(expected) * 0.001):
                bad.append(
                    f'Amt={amt} Qty={qty} Rate={rate} Expected={expected}')
            checked += 1
        except:
            pass
    summary = ' | '.join(bad[:20])
    if len(bad) > 20:
        summary += f' ... and {len(bad)-20} more'
    return pct(len(bad), checked) if checked else 0.0, len(bad), summary

# ─────────────────────────────────────────────
#  VALIDATION ENGINE
# ─────────────────────────────────────────────

NON_NUMERIC_HINTS = [
    'id','number','no','num','order','code','date','dt',
    'name','desc','city','state','zip','postal','country',
    'address','addr','phone','email','status','type','category'
]

def is_numeric_col(col_name, series):
    name_lower = col_name.lower()
    for hint in NON_NUMERIC_HINTS:
        if hint in name_lower:
            return False
    if looks_like_id_or_text_field(series):
        return False
    if looks_like_date_field(series):
        return False
    return True

def run_validations(filepath, selected_checks):
    sheet_name = get_data_sheet(filepath)
    df   = pd.read_excel(filepath, sheet_name=sheet_name, dtype=str)
    cols = df.columns.tolist()
    print(f"[INFO] Columns : {cols}")
    print(f"[INFO] Rows    : {len(df)}")

    results = {'total_rows': len(df), 'sheet_name': sheet_name}

    date_col   = detect_col(cols, ['date','dt','dob','created','updated','time'])
    city_col   = detect_col(cols, ['city','town','municipality'])
    state_col  = detect_col(cols, ['state','province'])
    zip_col    = detect_col(cols, ['zip','postal','zipcode','zip_code'])
    amount_col = detect_col(cols, ['salesorderamount','saleorderamount',
                                   'amount','total','revenue'])
    qty_col    = detect_col(cols, ['quantity','qty','units'])
    rate_col   = detect_col(cols, ['rate','unitprice','unit_price','unitrate'])

    if 'Check for Null %s' in selected_checks:
        results['nulls'] = [
            (c,) + null_check(df[c]) for c in cols]

    if 'Check for Blank %s' in selected_checks:
        results['blanks'] = [
            (c,) + blank_check(df[c]) for c in cols]

    if 'Check for Invalid dates' in selected_checks:
        if date_col:
            results['dates'] = [(date_col,) + invalid_date_check(df[date_col])]
        else:
            results['dates'] = [('(No date column detected)', 0.0, 0, '')]

    if 'Check for Invalid Amounts' in selected_checks:
        rows = []
        numeric_kw = ['amount','qty','quantity','rate','price',
                      'total','sum','revenue','sales','units','cost']
        for col in cols:
            if any(k in col.lower() for k in numeric_kw):
                if is_numeric_col(col, df[col]):
                    rows.append((col,) + invalid_numeric_check(df[col]))
                else:
                    print(f"[SKIP] '{col}' — not a numeric field")
        if amount_col and qty_col and rate_col:
            rows.append(
                ('SalesOrderAmount vs Qty x Rate',)
                + salesorder_amount_check(df, amount_col, qty_col, rate_col))
        if not rows:
            rows.append(('(No numeric columns detected)', 0.0, 0, ''))
        results['numeric'] = rows

    if 'Check for other Invalid values' in selected_checks:
        rows = []
        if city_col:
            rows.append((f'{city_col} (City)',)
                        + invalid_city_check(df[city_col]))
        if state_col:
            rows.append((f'{state_col} (State)',)
                        + invalid_state_check(df[state_col]))
        if zip_col:
            rows.append((f'{zip_col} (Zip)',)
                        + invalid_zip_check(df[zip_col]))
        if not rows:
            rows.append(('(No city/state/zip columns detected)', 0.0, 0, ''))
        results['geo'] = rows

    return results

# ─────────────────────────────────────────────
#  WRITE RESULTS TO EXCEL
# ─────────────────────────────────────────────

def write_results(filepath, results):
    wb = openpyxl.load_workbook(filepath)
    if RESULTS_TAB in wb.sheetnames:
        del wb[RESULTS_TAB]
    ws = wb.create_sheet(RESULTS_TAB)

    title_fill   = PatternFill("solid", fgColor="0B0B0F")
    title_font   = Font(bold=True, color="A259FF", size=14)
    header_fill  = PatternFill("solid", fgColor="6C3FF5")
    header_font  = Font(bold=True, color="FFFFFF", size=11)
    section_fill = PatternFill("solid", fgColor="F0EEFF")
    section_font = Font(bold=True, color="6C3FF5", size=11)
    green_fill   = PatternFill("solid", fgColor="E6FAF5")
    red_fill     = PatternFill("solid", fgColor="FFF1F0")
    yellow_fill  = PatternFill("solid", fgColor="FFFBE6")
    alt_fill     = PatternFill("solid", fgColor="F7F8FC")
    border = Border(
        left=Side(style='thin',   color='E0E0EF'),
        right=Side(style='thin',  color='E0E0EF'),
        top=Side(style='thin',    color='E0E0EF'),
        bottom=Side(style='thin', color='E0E0EF')
    )
    center = Alignment(horizontal='center', vertical='center')
    left   = Alignment(horizontal='left',   vertical='center', wrap_text=True)

    row = 1
    ws.merge_cells(f'A{row}:G{row}')
    cell = ws.cell(row=row, column=1,
                   value='Machinify  |  Data Analysis & Validation Report')
    cell.fill = title_fill; cell.font = title_font; cell.alignment = center
    ws.row_dimensions[row].height = 34
    row += 1

    ws.merge_cells(f'A{row}:G{row}')
    ts = ws.cell(row=row, column=1,
                 value=(f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}'
                        f'   |   Source Sheet: {results.get("sheet_name","?")}'))
    ts.font = Font(italic=True, color="9898B0", size=10)
    ts.fill = title_fill; ts.alignment = center
    row += 2

    def section_header(title):
        nonlocal row
        ws.merge_cells(f'A{row}:G{row}')
        c = ws.cell(row=row, column=1, value=title)
        c.fill = section_fill; c.font = section_font; c.alignment = left
        ws.row_dimensions[row].height = 24
        row += 1

    def col_header():
        nonlocal row
        for ci, h in enumerate(
            ['Check Type','Field / Column','Total Rows',
             'Invalid Count','Result (%)','Status',
             'Invalid Values (unique samples)'], 1):
            c = ws.cell(row=row, column=ci, value=h)
            c.fill = header_fill; c.font = header_font
            c.alignment = center; c.border = border
        ws.row_dimensions[row].height = 22
        row += 1

    def data_row(check, field, total, inv_count, value, bad_vals, alt=False):
        nonlocal row
        for ci, v in enumerate([check, field, total, inv_count, f'{value}%'], 1):
            c = ws.cell(row=row, column=ci, value=v)
            c.alignment = left if ci <= 2 else center
            c.border = border
            if alt: c.fill = alt_fill
        sc = ws.cell(row=row, column=6)
        sc.alignment = center; sc.border = border
        if value == 0.0:
            sc.value = 'PASS'; sc.fill = green_fill
            sc.font = Font(bold=True, color="00C48C")
        elif value <= 5.0:
            sc.value = 'LOW';  sc.fill = yellow_fill
            sc.font = Font(bold=True, color="FAAD14")
        elif value <= 20.0:
            sc.value = 'MEDIUM'; sc.fill = yellow_fill
            sc.font = Font(bold=True, color="FAAD14")
        else:
            sc.value = 'HIGH'; sc.fill = red_fill
            sc.font = Font(bold=True, color="FF4D4F")
        iv = ws.cell(row=row, column=7, value=bad_vals)
        iv.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        iv.border = border
        if alt: iv.fill = alt_fill
        ws.row_dimensions[row].height = max(
            30, min(120, 15 * (1 + bad_vals.count(',') // 3)))
        row += 1

    def write_section(label, data_list, total_rows):
        section_header(label)
        col_header()
        for i, (field, val, inv_count, bad_vals) in enumerate(data_list):
            data_row(label.split()[0], field,
                     total_rows, inv_count, val, bad_vals, alt=i % 2 == 1)
        nonlocal row
        row += 1

    total_rows = results.get('total_rows', 0)
    if 'nulls'   in results: write_section('NULL % Checks',           results['nulls'],   total_rows)
    if 'blanks'  in results: write_section('BLANK % Checks',          results['blanks'],  total_rows)
    if 'dates'   in results: write_section('Date Validation Checks',  results['dates'],   total_rows)
    if 'numeric' in results: write_section('Numeric & Amount Checks', results['numeric'], total_rows)
    if 'geo'     in results: write_section('Geographic Value Checks', results['geo'],     total_rows)

    ws.column_dimensions['A'].width = 24
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 12
    ws.column_dimensions['F'].width = 10
    ws.column_dimensions['G'].width = 60
    wb.move_sheet(RESULTS_TAB, offset=-len(wb.sheetnames) + 1)
    wb.save(filepath)
    return RESULTS_TAB

# ─────────────────────────────────────────────
#  MACHINIFY LOGO
# ─────────────────────────────────────────────

def make_logo(parent):
    c = tk.Canvas(parent, width=200, height=44,
                  bg=BG_NAVBAR, highlightthickness=0)
    c.pack(side='left', padx=(16, 0), pady=0)
    c.create_line(4,  34, 13, 12, fill=ACCENT_1, width=4,
                  capstyle='round', joinstyle='round')
    c.create_line(13, 12, 21, 26, fill=ACCENT_1, width=4,
                  capstyle='round', joinstyle='round')
    c.create_line(21, 26, 29, 12, fill=ACCENT_1, width=4,
                  capstyle='round', joinstyle='round')
    c.create_line(29, 12, 38, 34, fill=ACCENT_1, width=4,
                  capstyle='round', joinstyle='round')
    c.create_text(46, 22, text='machinify', anchor='w',
                  font=('Helvetica Neue', 16, 'bold'),
                  fill='#FFFFFF')

# ─────────────────────────────────────────────
#  BUTTON HELPER
#  On macOS tk.Button ignores fg when relief=flat
#  Fix: use a Canvas-drawn button that we fully
#  control — background colour + text colour
# ─────────────────────────────────────────────

class MacButton(tk.Canvas):
    """
    A fully custom button drawn on a Canvas so that
    background AND foreground colours always work on macOS.
    """
    def __init__(self, parent, text, command,
                 bg='#6C3FF5', fg='#FFFFFF',
                 hover_bg='#5533CC', hover_fg='#FFFFFF',
                 font_spec=('Helvetica Neue', 11, 'bold'),
                 pad_x=20, pad_y=10, **kwargs):
        # measure text size to auto-size canvas
        tmp = tk.Label(font=font_spec, text=text)
        tw  = tmp.winfo_reqwidth()  + pad_x * 2
        th  = tmp.winfo_reqheight() + pad_y * 2
        tmp.destroy()

        super().__init__(parent, width=tw, height=th,
                         bg=parent['bg'],
                         highlightthickness=0,
                         cursor='hand2', **kwargs)

        self._text      = text
        self._command   = command
        self._bg        = bg
        self._fg        = fg
        self._hover_bg  = hover_bg
        self._hover_fg  = hover_fg
        self._font      = font_spec
        self._tw        = tw
        self._th        = th

        self._render(bg, fg)

        self.bind('<Enter>',    lambda e: self._render(hover_bg, hover_fg))
        self.bind('<Leave>',    lambda e: self._render(bg, fg))
        self.bind('<Button-1>', lambda e: command())

    def _render(self, bg, fg):
        self.delete('all')
        r = 6   # corner radius
        w, h = self._tw, self._th
        # draw rounded rectangle
        self.create_polygon(
            r, 0,  w-r, 0,
            w, r,  w, h-r,
            w-r, h, r, h,
            0, h-r, 0, r,
            fill=bg, outline=bg, smooth=True
        )
        self.create_text(
            w // 2, h // 2,
            text=self._text,
            fill=fg,
            font=self._font
        )

# ─────────────────────────────────────────────
#  GUI
# ─────────────────────────────────────────────

class DataValidationApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Machinify — Data Validation')
        self.geometry('820x700')
        self.resizable(False, False)
        self.configure(bg=BG_PAGE)
        self.filepath = None
        self._build_ui()

    def _build_ui(self):

        # ── NAV BAR ──
        nav = tk.Frame(self, bg=BG_NAVBAR, height=52)
        nav.pack(fill='x')
        nav.pack_propagate(False)
        make_logo(nav)
        tk.Frame(nav, bg='#2A2A3A', width=1,
                 height=24).pack(side='left', padx=14, pady=14)
        tk.Label(nav, text='Data Validation',
                 font=('Helvetica Neue', 12),
                 fg='#8888AA', bg=BG_NAVBAR).pack(side='left')

        # ── ACCENT LINE ──
        tk.Frame(self, bg=ACCENT_1, height=3).pack(fill='x')

        # ── HERO ──
        hero = tk.Frame(self, bg=BG_PAGE, pady=20)
        hero.pack(fill='x')
        tk.Label(hero,
                 text='Data Quality Validation',
                 font=('Helvetica Neue', 22, 'bold'),
                 fg=TEXT_DARK, bg=BG_PAGE).pack()
        tk.Label(hero,
                 text='Upload an Excel file, select checks, and run automated data quality analysis',
                 font=('Helvetica Neue', 11),
                 fg=TEXT_MID, bg=BG_PAGE).pack(pady=(2, 0))

        content = tk.Frame(self, bg=BG_PAGE)
        content.pack(fill='both', expand=True, padx=32)

        # ── CARD 1 : File ──
        c1 = self._card(content)
        c1.pack(fill='x', pady=(0, 14))
        self._tag_label(c1, 'STEP 1')
        tk.Label(c1, text='Choose Input Excel File',
                 font=('Helvetica Neue', 13, 'bold'),
                 fg=TEXT_DARK, bg=BG_CARD).pack(anchor='w', padx=20)
        tk.Label(c1,
                 text='Select the .xlsx or .xls file — validation runs on the first data sheet',
                 font=('Helvetica Neue', 10),
                 fg=TEXT_LIGHT, bg=BG_CARD).pack(anchor='w', padx=20, pady=(2, 10))

        file_row = tk.Frame(c1, bg=BG_CARD)
        file_row.pack(fill='x', padx=20, pady=(0, 16))

        self.file_pill = tk.Frame(file_row, bg='#EDEDF8',
                                  highlightbackground=BORDER_CLR,
                                  highlightthickness=1)
        self.file_pill.pack(side='left', fill='x', expand=True,
                            ipady=7, ipadx=10)
        self.file_label = tk.Label(
            self.file_pill,
            text='No file selected …',
            font=('Helvetica Neue', 10),
            fg=TEXT_LIGHT, bg='#EDEDF8', anchor='w')
        self.file_label.pack(fill='x', padx=8)

        # ── Browse File button — MacButton so fg is always visible ──
        MacButton(
            file_row,
            text='  Browse File  ',
            command=self.browse_file,
            bg=ACCENT_1, fg='#FFFFFF',
            hover_bg=BTN_HOVER, hover_fg='#FFFFFF',
            font_spec=('Helvetica Neue', 11, 'bold'),
            pad_x=18, pad_y=9
        ).pack(side='right', padx=(12, 0))

        # ── CARD 2 : Checks ──
        c2 = self._card(content)
        c2.pack(fill='x', pady=(0, 14))
        self._tag_label(c2, 'STEP 2')
        tk.Label(c2, text='Select Validation Checks',
                 font=('Helvetica Neue', 13, 'bold'),
                 fg=TEXT_DARK, bg=BG_CARD).pack(anchor='w', padx=20)
        tk.Label(c2,
                 text='Click one item  |  Hold \u2318 Cmd to select multiple',
                 font=('Helvetica Neue', 10),
                 fg=TEXT_LIGHT, bg=BG_CARD).pack(anchor='w', padx=20, pady=(2, 8))

        lb_frame = tk.Frame(c2, bg=BG_CARD)
        lb_frame.pack(fill='x', padx=20, pady=(0, 8))
        sb = tk.Scrollbar(lb_frame, orient='vertical')
        self.listbox = tk.Listbox(
            lb_frame,
            selectmode='multiple',
            height=5,
            font=('Helvetica Neue', 11),
            bg=BG_INPUT,
            fg=TEXT_DARK,
            selectbackground=ACCENT_1,
            selectforeground='#FFFFFF',
            relief='flat', bd=0,
            highlightthickness=1,
            highlightbackground=BORDER_CLR,
            activestyle='none',
            yscrollcommand=sb.set
        )
        sb.config(command=self.listbox.yview)
        sb.pack(side='right', fill='y')
        self.listbox.pack(fill='x')

        for item in ['Check for Null %s',
                     'Check for Blank %s',
                     'Check for Invalid dates',
                     'Check for Invalid Amounts',
                     'Check for other Invalid values']:
            self.listbox.insert('end', f'   {item}')

        tog = tk.Frame(c2, bg=BG_CARD)
        tog.pack(fill='x', padx=20, pady=(0, 14))
        for lbl, cmd in [('Select All', self.select_all),
                          ('Clear All',  self.clear_all)]:
            MacButton(
                tog,
                text=f'  {lbl}  ',
                command=cmd,
                bg='#EDEDF8', fg=ACCENT_1,
                hover_bg='#E0D8FF', hover_fg=BTN_HOVER,
                font_spec=('Helvetica Neue', 9, 'bold'),
                pad_x=10, pad_y=5
            ).pack(side='left', padx=(0, 8))

        # ── RUN VALIDATION button ──
        run_row = tk.Frame(content, bg=BG_PAGE)
        run_row.pack(pady=12)
        MacButton(
            run_row,
            text='   Run Validation   ',
            command=self.do_validation,
            bg=ACCENT_1, fg='#FFFFFF',
            hover_bg=BTN_HOVER, hover_fg='#FFFFFF',
            font_spec=('Helvetica Neue', 14, 'bold'),
            pad_x=28, pad_y=13
        ).pack()

        # ── STATUS ──
        self.status_var = tk.StringVar(
            value='Ready — select a file and choose your checks.')
        tk.Label(content,
                 textvariable=self.status_var,
                 font=('Helvetica Neue', 10, 'italic'),
                 fg=ACCENT_1, bg=BG_PAGE,
                 wraplength=740, justify='center'
                 ).pack(pady=(8, 4))

        # ── PROGRESS BAR ──
        style = ttk.Style()
        style.theme_use('default')
        style.configure('M.Horizontal.TProgressbar',
                        troughcolor='#EDEDF8',
                        background=ACCENT_1,
                        thickness=5)
        self.progress = ttk.Progressbar(
            content, mode='indeterminate', length=600,
            style='M.Horizontal.TProgressbar')
        self.progress.pack(pady=(0, 16))

    def _card(self, parent):
        return tk.Frame(parent, bg=BG_CARD,
                        highlightbackground=BORDER_CLR,
                        highlightthickness=1)

    def _tag_label(self, parent, text):
        row = tk.Frame(parent, bg=BG_CARD)
        row.pack(fill='x', padx=20, pady=(14, 4))
        tk.Label(row, text=f' {text} ',
                 font=('Helvetica Neue', 8, 'bold'),
                 fg='#FFFFFF', bg=ACCENT_1,
                 padx=6, pady=2).pack(side='left')

    def browse_file(self):
        path = filedialog.askopenfilename(
            title='Select Excel File',
            filetypes=[('Excel files', '*.xlsx *.xls'),
                       ('All files', '*.*')]
        )
        if path:
            self.filepath = path
            self.file_label.config(
                text=f'  {os.path.basename(path)}',
                fg=ACCENT_1,
                font=('Helvetica Neue', 10, 'bold'))
            self.file_pill.config(bg='#EDE8FF')
            self.file_label.config(bg='#EDE8FF')
            self.status_var.set(
                f'File loaded: {os.path.basename(path)}')

    def select_all(self):
        self.listbox.select_set(0, 'end')

    def clear_all(self):
        self.listbox.selection_clear(0, 'end')

    def do_validation(self):
        if not self.filepath:
            messagebox.showwarning('No File',
                'Please select an Excel file first.')
            return
        sel = self.listbox.curselection()
        if not sel:
            messagebox.showwarning('No Checks',
                'Please select at least one validation check.')
            return
        selected_checks = [self.listbox.get(i).strip() for i in sel]
        self.progress.start(10)
        self.status_var.set('Running validations … please wait …')
        self.update()
        try:
            results = run_validations(self.filepath, selected_checks)
            sheet   = write_results(self.filepath, results)
            self.progress.stop()
            self.status_var.set(
                f'Done!  Results written to "{sheet}"  |  Opening file …')
            self.update()
            subprocess.Popen(['open', self.filepath])
            messagebox.showinfo('Validation Complete',
                f'All selected validations completed!\n\n'
                f'Results saved to tab:\n"{sheet}"\n\n'
                f'Source data sheet:\n"{results["sheet_name"]}"\n\n'
                f'The file has been opened for you.')
        except Exception as e:
            self.progress.stop()
            self.status_var.set(f'Error: {e}')
            messagebox.showerror('Error',
                f'Something went wrong:\n\n{e}')

if __name__ == '__main__':
    app = DataValidationApp()
    app.mainloop()
ENDOFPYTHON

echo "Launching Machinify Data Validation App..."
python3 "$APP_FILE"