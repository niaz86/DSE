<div align="center">

# 📈 DSE Data Scraper & Tracker

[![Python 3.9+](https://img.shields.io/badge/Python-3.9%2B-blue.svg)](https://www.python.org/)
[![GitHub Actions](https://img.shields.io/github/actions/workflow/status/niaz86/DSE/dse_daily_scrape.yml?label=Daily%20Scrape&logo=github)](https://github.com/niaz86/DSE/actions)
[![License](https://img.shields.io/badge/License-Educational-informational.svg)](https://github.com/niaz86/DSE)

**Automated daily scraper for the Dhaka Stock Exchange.**  
Collects per-company financials, shareholding data, and market metrics —  
then appends a new row to a growing Excel workbook every trading day.

[🌐 **Live Project Page**](https://niaz86.github.io/DSE) &nbsp;·&nbsp; [📥 Download ZIP](https://github.com/niaz86/DSE/archive/refs/heads/master.zip) &nbsp;·&nbsp; [⚡ Quick Start](#-quick-start)

</div>

---

## ⚡ Quick Start

```bash
git clone https://github.com/niaz86/DSE.git
cd DSE
pip install -r requirements.txt
python script.py
```

Excel files will be created (or updated) in your local directory. That's it.

---

## ✨ Features

- **Multithreaded** — 10 concurrent threads scrape the full market 8–10× faster than serial
- **Time-series Excel** — one sheet per company, new dated row appended every trading day
- **Auto-updated** — GitHub Actions runs Mon–Fri after DSE closes; files committed back automatically
- **Market summary** — aggregated sector-level view in `dse_market_summary.xlsx`
- **JSON backup** — every scrape also exports to `dse_companies.json`
- **No API keys** — scrapes publicly available data from dsebd.org, zero setup required

---

## How It Works

Each weekday after market close the pipeline does five things:

1. **Discover** — fetches all trading codes from DSE categories A, B, G, N, Z (in parallel)
2. **Scrape** — 10 threads simultaneously pull each company page from dsebd.org
3. **Parse** — extracts price, P/E, EPS, NAV, dividends, and shareholding breakdown
4. **Append** — adds a new dated row to that company's sheet in `dse_companies.xlsx`
5. **Commit** — GitHub Actions pushes the updated files back to this repo

---

## Repository Structure

```
DSE/
├── dse_scraper_individual_company.py   # Core scraper — per-company data & Excel writer
├── script.py                           # Main entry point / orchestrator
├── DSE_Summary.py                      # Market-wide summary generator
│
├── dse_companies.xlsx                  # Per-company sheets, daily rows appended
├── dse_stocks.xlsx                     # Stock price data
├── dse_market_summary.xlsx             # Aggregated market summary
├── dse_companies.json                  # JSON backup of the latest scrape
├── last_run.txt                        # Timestamp of the most recent successful run
│
├── requirements.txt                    # Python dependencies
├── docs/
│   └── index.html                      # Animated project landing page (GitHub Pages)
└── .github/
    └── workflows/
        └── dse_daily_scrape.yml        # GitHub Actions workflow (runs Mon–Fri)
```

---

## Data Collected Per Company

Each daily row in `dse_companies.xlsx` captures 14 fields:

| Field | Description |
|---|---|
| Date | Scrape date (YYYY-MM-DD) |
| Opening Price | Price at market open in BDT |
| Market Cap (mn) | Market capitalisation in BDT millions |
| 52W High / Low | 52-week trading range |
| Trailing P/E | Price-to-earnings ratio from latest unaudited data |
| EPS Basic | Latest interim earnings per share (continuing operations) |
| NAV Per Share | Net asset value from latest audited accounts |
| Cash Dividend | Most recent cash dividend declared |
| Bonus Dividend | Most recent bonus stock dividend |
| Paid-up Capital (mn) | Total paid-up capital in BDT millions |
| Authorized Capital (mn) | Total authorised capital in BDT millions |
| Sponsor/Dir % | Sponsor & director shareholding percentage |
| Institute % | Institutional shareholding percentage |
| Public % | Public shareholding percentage |

Static company info (name, sector, address, listing year, debut date) is written once to rows 1–10 of each sheet and preserved across all future runs.

---

## Getting Started

### Prerequisites

- Python 3.9+
- pip

### Installation

```bash
git clone https://github.com/niaz86/DSE.git
cd DSE
pip install -r requirements.txt
```

### Run Manually

```bash
# Scrape all DSE-listed companies
python script.py

# Scrape a single company
python dse_scraper_individual_company.py
```

### Configuration

Edit the top of `dse_scraper_individual_company.py`:

| Variable | Default | Description |
|---|---|---|
| `MAX_WORKERS` | `10` | Number of parallel scraping threads |
| `REQUEST_DELAY` | `0.5s` | Per-thread delay between requests |
| `CATEGORIES` | `A,B,G,N,Z` | DSE market categories to include |
| `DSE_OUTPUT_EXCEL` | env var | Override output path for the Excel file |
| `DSE_OUTPUT_JSON` | env var | Override output path for the JSON backup |

---

## Automated Daily Updates (GitHub Actions)

The workflow in `.github/workflows/dse_daily_scrape.yml` runs automatically:

- **Schedule:** Monday–Friday at **12:30 UTC** (18:30 BDT — after DSE closes at 14:30 BDT)
- **What it does:** checks out the repo → runs the scraper → commits updated Excel/JSON files back
- **Manual trigger:** Actions → DSE Daily Scrape → Run workflow

No secrets or tokens need to be configured — the workflow uses the built-in `GITHUB_TOKEN`.

---

## Output Files

**`dse_companies.xlsx`** — one sheet per trading code. Rows 1–10 hold static company info; row 12 has column headers; row 13 onward is one row per trading day, newest at the bottom.

**`dse_market_summary.xlsx`** — aggregated view across all companies, useful for sector-level or market-wide analysis.

**`dse_companies.json`** — flat JSON snapshot of the most recent scrape, useful for programmatic access without parsing Excel.

---

## 🔧 Troubleshooting

**Excel files not updating?**  
Check `last_run.txt` to see when the scraper last ran. Run `python script.py` manually to check for errors.

**GitHub Actions not running?**  
Go to Actions → DSE Daily Scrape and check the workflow logs. Manually trigger with Run workflow.

**`ModuleNotFoundError`?**  
Run `pip install -r requirements.txt`. Consider using a virtual environment: `python -m venv venv && source venv/bin/activate`.

**Empty cells in Excel?**  
Some companies don't report all fields (e.g. EPS for newly listed companies). DSE also publishes holidays with no trading data — those runs produce no diff and nothing is committed.

---

## 📊 Use Cases

- **Portfolio tracking** — monitor holdings' price movements and key metrics over time
- **Sector analysis** — compare performance across DSE sectors
- **Fundamental research** — track P/E, NAV, dividend yields, and shareholding patterns over time
- **Data science** — feed historical DSE data into models or visualisations

---

## Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/your-feature`
3. Commit your changes: `git commit -m 'Add your feature'`
4. Push: `git push origin feature/your-feature`
5. Open a Pull Request

Ideas welcome: technical indicators, data visualisation, session pooling for faster scraping, or support for additional exchanges.

---

## 📝 Notes

- Scrapes publicly available data from [dsebd.org](https://www.dsebd.org). Use responsibly; avoid running at very high concurrency.
- DSE is closed on weekends and Bangladeshi public holidays. The workflow runs Mon–Fri but produces an empty diff on holidays — nothing is committed.
- Data quality depends on what DSE publishes. Missing values appear as empty cells.

---

## 📄 License

For personal research and educational use. Data sourced from the Dhaka Stock Exchange website.
