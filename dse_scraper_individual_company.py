import requests
import json
import time
import os
import sys
import re
import threading
from datetime import datetime, timezone, timedelta
from urllib.parse import quote
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from concurrent.futures import ThreadPoolExecutor, as_completed

# ── Configuration ─────────────────────────────────────────────────────────────
BASE_URL      = "https://www.dse.com.bd"
# Code discovery now uses the JSON price feed (same one script.py uses).
PRICES_API    = BASE_URL + "/api/live/prices"
# Company page on the new site, e.g. https://www.dse.com.bd/company/ACI
# (server-rendered HTML, so plain requests + BeautifulSoup works).
COMPANY_URL   = BASE_URL + "/company/{code}"
MIN_SUCCESS_RATIO = 0.8   # below this share of companies OK, abort and save nothing
BDT           = timezone(timedelta(hours=6))   # runners are UTC; DSE is Bangladesh time
def today_bdt():
    return datetime.now(BDT).strftime("%Y-%m-%d")
CATEGORIES    = ["A", "B", "G", "N", "Z"]
OUTPUT_JSON   = os.environ.get("DSE_OUTPUT_JSON",  "dse_companies.json")
OUTPUT_EXCEL  = os.environ.get("DSE_OUTPUT_EXCEL", "dse_companies.xlsx")
REQUEST_DELAY = 1.0   # seconds each worker waits before a request
MAX_WORKERS   = 4     # ~3 req/s overall; be gentle with the exchange's servers

HEADERS = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-GB,en-US;q=0.9,en;q=0.8",
    "Cache-Control": "max-age=0",
    "Connection": "keep-alive",
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/146.0.0.0 Safari/537.36"
    ),
    "Referer": BASE_URL,
}

# ── Column layout for each stock sheet ────────────────────────────────────────
DAILY_COLUMNS = [
    ("Date",                    "last_updated"),
    ("Opening Price",           "opening_price"),
    ("Market Cap (mn)",         "market_cap_mn"),
    ("52W High",                "week52_high"),
    ("52W Low",                 "week52_low"),
    ("Trailing P/E",            "trailing_pe"),
    ("EPS Basic",               "latest_eps_basic"),
    ("NAV Per Share",           "nav_per_share"),
    ("Cash Dividend",           "cash_dividend"),
    ("Bonus Dividend",          "bonus_stock_dividend"),
    ("Paid-up Capital (mn)",    "paid_up_capital_mn"),
    ("Authorized Capital (mn)", "authorized_capital_mn"),
    ("Sponsor/Dir %",           "sponsor_director"),
    ("Institute %",             "institute"),
    ("Public %",                "public"),
]


# ── Scraping helpers ──────────────────────────────────────────────────────────
PRICE_META = {}   # code -> {"sector":..., "category":..., "open":...} from the price feed

LABELS = [
    "Market cap", "Free float market cap", "NAV per share", "P/E (audited)",
    "Day's value", "Day's volume", "Day's trades", "Opening price",
    "Adjusted Opening Price", "Last Update", "Day's range", "52-week range",
    "Basics", "Listing year", "Authorized Capital (mn)", "Paid-up capital",
    "Face value", "Total outstanding securities", "Market Lot",
    "Shareholding pattern", "Shareholding percentages as reported",
    "Corporate & contact", "Present Operational Status",
    "Head Office / Registered Office", "Factory Address", "Contact Phone", "Fax",
    "E-mail", "Company Secretary Name", "Company Secretary Cell No.",
    "Company Secretary Telephone No.", "Company Secretary E-mail", "Website",
    "Loan status", "Present Loan Status", "Credit rating", "Last AGM",
    "Year-end", "Details of Financial Statement", "Price Sensitive Information",
    "(i) Dynamic information",
]
HOLDING_LABELS = ["Sponsor / Director", "Government", "Institution", "Foreign", "Public"]

def fetch_page(trading_code):
    url = COMPANY_URL.format(code=quote(trading_code, safe="()-"))
    last = None
    for attempt in range(3):
        try:
            resp = requests.get(url, headers=HEADERS, timeout=20)
            if resp.status_code >= 500:
                raise requests.HTTPError(f"{resp.status_code} from server")
            resp.raise_for_status()
            return BeautifulSoup(resp.text, "html.parser")
        except (requests.Timeout, requests.ConnectionError, requests.HTTPError) as e:
            last = e
            if isinstance(e, requests.HTTPError) and e.response is not None \
                    and e.response.status_code < 500:
                raise                      # 404/410 etc: no point retrying
            time.sleep(2 ** attempt)
    raise last

def split_by_labels(text, labels):
    """Return {label: text up to the next label}, first occurrence wins.
    Works on flattened text, so it doesn't matter how the HTML nests the nodes."""
    hits = []
    for lab in labels:
        for m in re.finditer(r"(?<!\w)" + re.escape(lab) + r"(?!\w)", text):   # case-sensitive
            hits.append((m.start(), m.end(), lab))
    # drop hits that sit inside a longer label (e.g. "E-mail" inside "Company Secretary E-mail")
    hits = [h for h in hits
            if not any(o is not h and o[0] <= h[0] and h[1] <= o[1] and (o[1] - o[0]) > (h[1] - h[0])
                       for o in hits)]
    hits.sort()
    out = {}
    for i, (st, en, lab) in enumerate(hits):
        nxt = hits[i + 1][0] if i + 1 < len(hits) else len(text)
        out.setdefault(lab, text[en:nxt].strip(" :·-"))
    return out

def clean_num(v):
    """'BDT 878.32 mn' -> '878.32'; '—' -> ''."""
    v = re.sub(r"\bBDT\b|\bmn\b", "", v or "").strip()
    return "" if v in ("—", "-", "–") else v

def numbers(v):
    return re.findall(r"\d[\d,]*\.?\d*", (v or "").replace("BDT", " "))

def page_text(soup):
    for t in soup(["script", "style", "noscript"]):
        t.decompose()
    anchor = soup.find(string=lambda s: s and "Back to market" in s)
    h1 = anchor.find_next("h1") if anchor else soup.find("h1")
    name = h1.get_text(" ", strip=True) if h1 else ""
    flat = re.sub(r"\s+", " ", soup.get_text(" ", strip=True))
    i = flat.find("Back to market")           # drop the live ticker / nav above it
    return name, (flat[i:] if i >= 0 else flat)

def scrape_company(trading_code):
    soup = fetch_page(trading_code)
    name, flat = page_text(soup)
    vals = split_by_labels(flat, LABELS)
    if not name or len(vals) < 5:
        raise ValueError(f"unrecognised page layout (name={name!r}, labels found={len(vals)}); "
                         f"text starts: {flat[:200]!r}")

    rng = numbers(vals.get("52-week range"))          # [current, low, high]
    week52_low, week52_high = (rng[1], rng[2]) if len(rng) >= 3 else ("", "")

    hold = split_by_labels(vals.get("Shareholding pattern", ""), HOLDING_LABELS)
    def latest_pct(label):                             # last column = most recent date
        p = re.findall(r"(\d+(?:\.\d+)?)\s*%", hold.get(label, ""))
        return p[-1] if p else ""

    debut = re.search(r"Debut\s*·\s*(\d{1,2}[-/ ][A-Za-z0-9]+[-/ ]\d{2,4}|[A-Za-z]{3,9} \d{1,2},? \d{4})", flat)
    meta = PRICE_META.get(trading_code, {})

    return {
        "trading_code":          trading_code,
        "company_name":          name,
        "sector":                meta.get("sector", ""),
        "face_par_value":        clean_num(vals.get("Face value")),
        "debut_trading_date":    debut.group(1) if debut else "",
        "listing_year":          clean_num(vals.get("Listing year")),
        "market_category":       meta.get("category", ""),
        "address": {
            "head_office":   vals.get("Head Office / Registered Office", ""),
            "contact_phone": vals.get("Contact Phone", ""),
            "email":         vals.get("E-mail", ""),
            "web_address":   vals.get("Website", ""),
        },
        "authorized_capital_mn": clean_num(vals.get("Authorized Capital (mn)")),
        "paid_up_capital_mn":    clean_num(vals.get("Paid-up capital")),
        "nav_per_share":         clean_num(vals.get("NAV per share")),
        # Not on the overview page (Financials / Corporate actions tabs): left blank.
        "latest_eps_basic":      "",
        "cash_dividend":         "",
        "bonus_stock_dividend":  "",
        "sponsor_director":      latest_pct("Sponsor / Director"),
        "institute":             latest_pct("Institution"),
        "public":                latest_pct("Public"),
        "week52_high":           week52_high,
        "week52_low":            week52_low,
        "trailing_pe":           "",   # page only shows "P/E (audited)", a different measure
        "opening_price":         clean_num(vals.get("Opening price")) or str(meta.get("open", "")),
        "market_cap_mn":         clean_num(vals.get("Market cap")),
        "last_updated":          today_bdt(),
    }


# ── Excel writer ──────────────────────────────────────────────────────────────
def safe_sheet_name(code):
    for ch in [":", "\\", "/", "?", "*", "[", "]"]:
        code = code.replace(ch, "_")
    return code[:31]

def style_header_cell(cell, text):
    cell.value = cell.value if text is None else text
    cell.font      = Font(bold=True, color="FFFFFF")
    cell.fill      = PatternFill("solid", fgColor="004B5A")
    cell.alignment = Alignment(horizontal="center")

def write_static_block(ws, data):
    static_fields = [
        ("Company Name",    data["company_name"]),
        ("Trading Code",    data["trading_code"]),
        ("Sector",          data["sector"]),
        ("Market Category", data["market_category"]),
        ("Face Value",      data["face_par_value"]),
        ("Listing Year",    data["listing_year"]),
        ("Debut Date",      data["debut_trading_date"]),
        ("Address",         data["address"]["head_office"]),
        ("Phone",           data["address"]["contact_phone"]),
        ("Email",           data["address"]["email"]),
    ]
    for i, (label, value) in enumerate(static_fields, start=1):
        label_cell = ws.cell(row=i, column=1, value=label)
        label_cell.font = Font(bold=True)
        ws.cell(row=i, column=2, value=value)

    for col_idx, (header_label, _) in enumerate(DAILY_COLUMNS, start=1):
        cell = ws.cell(row=12, column=col_idx, value=header_label)
        style_header_cell(cell, header_label)

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 30
    for col_idx in range(1, len(DAILY_COLUMNS) + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 18

def append_daily_row(ws, data):
    next_row = 13
    while ws.cell(row=next_row, column=1).value is not None:
        next_row += 1
    for col_idx, (_, key) in enumerate(DAILY_COLUMNS, start=1):
        ws.cell(row=next_row, column=col_idx, value=data.get(key, ""))

def update_excel(filepath, all_data):
    if os.path.exists(filepath):
        wb = load_workbook(filepath)
    else:
        wb = Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

    for data in all_data:
        sheet_name = safe_sheet_name(data["trading_code"])
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.create_sheet(title=sheet_name)
            write_static_block(ws, data)
        append_daily_row(ws, data)

    wb.save(filepath)
    print(f"Excel workbook saved: '{filepath}' ({len(wb.sheetnames)} sheets).")


# ── Discovery ─────────────────────────────────────────────────────────────────
def discover_all_codes():
    resp = requests.get(PRICES_API, headers={**HEADERS, "Accept": "application/json"}, timeout=20)
    resp.raise_for_status()
    payload = resp.json()
    cols, rows = payload["cols"], payload["rows"]
    ci, gi = cols.index("code"), cols.index("category")
    si = cols.index("sector") if "sector" in cols else None
    oi = cols.index("open") if "open" in cols else None
    seen, unique = set(), []
    for r in rows:
        code = str(r[ci]).strip()
        if code and str(r[gi]).upper() in CATEGORIES and code not in seen:
            seen.add(code)
            unique.append(code)
            PRICE_META[code] = {"category": str(r[gi]).upper(),
                                "sector": str(r[si]) if si is not None else "",
                                "open": r[oi] if oi is not None else ""}
    print(f"Total unique trading codes discovered: {len(unique)}\n")
    return unique


# ── Worker ────────────────────────────────────────────────────────────────────
def scrape_worker(args):
    index, total, code = args
    time.sleep(REQUEST_DELAY)
    try:
        data = scrape_company(code)
        return index, total, code, data, None
    except Exception as e:
        return index, total, code, None, e


# ── Main ──────────────────────────────────────────────────────────────────────
def run_daily_scrape():
    print(f"=== DSE Daily Scrape — {today_bdt()} ===\n")
    print(f"Thread pool size: {MAX_WORKERS}\n")

    all_codes = discover_all_codes()
    total     = len(all_codes)
    if total == 0:
        sys.exit("No trading codes discovered - aborting, existing files untouched.")

    store   = {}
    lock    = threading.Lock()
    success = fail = 0

    work_items = [(i, total, code) for i, code in enumerate(all_codes, 1)]

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(scrape_worker, item): item for item in work_items}

        for future in as_completed(futures):
            index, total_, code, data, err = future.result()
            if err:
                with lock:
                    fail += 1
                print(f"[{index}/{total_}] {code} — FAILED: {err}")
            else:
                with lock:
                    success += 1
                    store[code] = data
                print(f"[{index}/{total_}] {code} — OK")

    if success < total * MIN_SUCCESS_RATIO:
        sys.exit(f"Only {success}/{total} companies scraped (< {MIN_SUCCESS_RATIO:.0%}). "
                 "Aborting - existing JSON/Excel left untouched.")

    scraped = [store[code] for code in all_codes if code in store]

    # Field coverage: shows exactly which columns come out blank, and for whom
    print("\n=== FIELD COVERAGE (filled / scraped) ===")
    for _, key in DAILY_COLUMNS[1:] + [("Sector", "sector"), ("Listing year", "listing_year"),
                                       ("Face value", "face_par_value")]:
        missing = [d["trading_code"] for d in scraped if not str(d.get(key, "")).strip()]
        eg = f"  e.g. {', '.join(missing[:4])}" if missing and len(missing) < len(scraped) else ""
        print(f"  {key:<22} {len(scraped) - len(missing):>4}/{len(scraped)}{eg}")
    print()

    tmp = OUTPUT_JSON + ".tmp"            # write atomically: never leave a half file
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(store, f, indent=2, ensure_ascii=False)
    os.replace(tmp, OUTPUT_JSON)
    print(f"\nJSON backup saved: '{OUTPUT_JSON}'.")

    update_excel(OUTPUT_EXCEL, scraped)
    print(f"\nDone. {success} succeeded, {fail} failed.")


if __name__ == "__main__":
    run_daily_scrape()
