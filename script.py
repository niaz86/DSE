# -*- coding: utf-8 -*-
"""DSE"""

import requests
from bs4 import BeautifulSoup
import os
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ===================== ✅ NEW JSON IMPORT =====================
import json

# ── Config ────────────────────────────────────────────────────────────────────
EXCEL_FILE = "dse_stocks.xlsx"
JSON_FILE  = "dse_stocks.json"   # ✅ NEW
URL        = "https://www.dsebd.org/latest_share_price_scroll_group.php"

# ── Styles ────────────────────────────────────────────────────────────────────
HEADER_FILL = PatternFill("solid", start_color="006B6B")
HEADER_FONT = Font(bold=True, color="FFFFFF", name="Arial", size=10)
ALT_FILL    = PatternFill("solid", start_color="F2FAFA")
UP_FILL     = PatternFill("solid", start_color="E6F9EC")
DOWN_FILL   = PatternFill("solid", start_color="FDE8E8")
UP_FONT     = Font(name="Arial", size=10, color="1A7A1A", bold=True)
DOWN_FONT   = Font(name="Arial", size=10, color="CC0000", bold=True)
NORMAL_FONT = Font(name="Arial", size=10)
BORDER = Border(
    left=Side(style="thin",   color="CCCCCC"),
    right=Side(style="thin",  color="CCCCCC"),
    top=Side(style="thin",    color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)

DATA_COLS  = ["DATE", "#", "TRADING CODE", "GROUP", "LTP*", "HIGH", "LOW",
              "CLOSEP*", "YCP*", "CHANGE", "TRADE", "VALUE (mn)", "VOLUME"]
COL_WIDTHS = [12,5,16,7,9,9,9,10,9,9,9,12,14]

# ── Session ───────────────────────────────────────────────────────────────────
session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0"
})

# ── Fetch ─────────────────────────────────────────────────────────────────────
def fetch_group(group="A"):
    resp = session.get(URL, params={"group": group}, timeout=15)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    rows = []
    for a_tag in soup.select("a.ab1"):
        tr = a_tag.find_parent("tr")
        if not tr:
            continue
        tds = tr.find_all("td")
        if len(tds) < 11:
            continue
        rows.append({
            "#": tds[0].get_text(strip=True),
            "TRADING CODE": a_tag.get_text(strip=True),
            "LTP*": tds[2].get_text(strip=True),
            "HIGH": tds[3].get_text(strip=True),
            "LOW": tds[4].get_text(strip=True),
            "CLOSEP*": tds[5].get_text(strip=True),
            "YCP*": tds[6].get_text(strip=True),
            "CHANGE": tds[7].get_text(strip=True),
            "TRADE": tds[8].get_text(strip=True),
            "VALUE (mn)": tds[9].get_text(strip=True),
            "VOLUME": tds[10].get_text(strip=True),
        })
    return rows

# ── Excel helpers ─────────────────────────────────────────────────────────────
def style_header(ws):
    for col_idx, (col, width) in enumerate(zip(DATA_COLS, COL_WIDTHS), 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.value = col
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(col_idx)].width = width

def get_or_create_sheet(wb, name):
    safe = name[:31]
    if safe in wb.sheetnames:
        return wb[safe], False
    ws = wb.create_sheet(title=safe)
    style_header(ws)
    return ws, True

def build_row(today, stock, group):
    return [
        today,
        stock.get("#", ""),
        stock.get("TRADING CODE", ""),
        group,
        stock.get("LTP*", ""),
        stock.get("HIGH", ""),
        stock.get("LOW", ""),
        stock.get("CLOSEP*", ""),
        stock.get("YCP*", ""),
        stock.get("CHANGE", ""),
        stock.get("TRADE", ""),
        stock.get("VALUE (mn)", ""),
        stock.get("VOLUME", ""),
    ]

# ===================== ✅ JSON HELPERS =====================
def load_json():
    if os.path.isfile(JSON_FILE):
        with open(JSON_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_json(data):
    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

def build_json_row(today, stock, group):
    return {
        "DATE": today,
        "#": stock.get("#", ""),
        "TRADING CODE": stock.get("TRADING CODE", ""),
        "GROUP": group,
        "LTP*": stock.get("LTP*", ""),
        "HIGH": stock.get("HIGH", ""),
        "LOW": stock.get("LOW", ""),
        "CLOSEP*": stock.get("CLOSEP*", ""),
        "YCP*": stock.get("YCP*", ""),
        "CHANGE": stock.get("CHANGE", ""),
        "TRADE": stock.get("TRADE", ""),
        "VALUE (mn)": stock.get("VALUE (mn)", ""),
        "VOLUME": stock.get("VOLUME", ""),
    }

def json_is_duplicate(stock_list, new_row):
    if not stock_list:
        return False
    last = stock_list[-1]
    keys = ["GROUP","LTP*","HIGH","LOW","CLOSEP*","YCP*","CHANGE","TRADE","VALUE (mn)","VOLUME"]
    return all(str(last.get(k,"")) == str(new_row.get(k,"")) for k in keys)

# ── Main ──────────────────────────────────────────────────────────────────────
def run(groups=("A","B","G","N","Z")):

    wb = load_workbook(EXCEL_FILE) if os.path.isfile(EXCEL_FILE) else Workbook()

    # ===================== ✅ LOAD JSON =====================
    json_data = load_json()
    json_created = json_appended = json_duplicate = 0

    today = datetime.now().strftime("%Y-%m-%d")

    for group in groups:
        stocks = fetch_group(group)

        for stock in stocks:
            code = stock.get("TRADING CODE", "").strip()
            if not code:
                continue

            # ===== EXISTING EXCEL LOGIC (UNCHANGED) =====
            ws, is_new = get_or_create_sheet(wb, code)
            ws.append(build_row(today, stock, group))

            # ===================== ✅ JSON LOGIC =====================
            json_row = build_json_row(today, stock, group)

            if code not in json_data:
                json_data[code] = [json_row]
                json_created += 1
            else:
                if json_is_duplicate(json_data[code], json_row):
                    json_duplicate += 1
                else:
                    json_data[code].append(json_row)
                    json_appended += 1

    wb.save(EXCEL_FILE)

    # ===================== ✅ SAVE JSON =====================
    save_json(json_data)

    print("\nJSON SUMMARY")
    print("New:", json_created)
    print("Appended:", json_appended)
    print("Duplicate:", json_duplicate)

if __name__ == "__main__":
    run()

# keep your last_run.txt
from datetime import datetime
with open("last_run.txt", "w") as f:
    f.write(datetime.now().strftime("%Y-%m-%d"))
