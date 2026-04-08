# 📈 DSE Data Scraper & Tracker

[![Python 3.9+](https://img.shields.io/badge/Python-3.9%2B-blue.svg)](https://www.python.org/)
[![GitHub Actions](https://img.shields.io/badge/GitHub%20Actions-Automated-green.svg)](https://github.com/features/actions)
[![License](https://img.shields.io/badge/License-Educational-informational.svg)](https://github.com/niaz86/DSE)

Automated daily scraper for the **Dhaka Stock Exchange (DSE)** that collects per-company financials, shareholding data, and market metrics — then appends them into a growing Excel workbook, one row per trading day.

Data files are updated automatically every weekday via **GitHub Actions** and committed back to this repository, so the Excel files always contain the latest available figures.

> **Perfect for:** Financial analysts, researchers, portfolio managers, and anyone tracking DSE-listed companies' performance over time.

---

## ⚡ Quick Start

### 30-Second Setup

```bash
git clone https://github.com/niaz86/DSE.git
cd DSE
pip install -r requirements.txt
python script.py
```

That's it! Excel files will be created/updated in your local directory.

---

## ✨ Features

- ✅ **Automatic Daily Updates** — GitHub Actions runs Mon–Fri after market close
- ✅ **Per-Company Tracking** — Individual Excel sheets for each DSE-listed company
- ✅ **Rich Data Collection** — Price, P/E, EPS, NAV, dividends, shareholding breakdown
- ✅ **Market-Wide Summary** — Aggregated metrics across all companies
- ✅ **Parallel Scraping** — 10 concurrent threads for fast data collection
- ✅ **Zero Configuration** — Works out of the box; no API keys or authentication needed
- ✅ **JSON Backup** — Latest scrape also exported as JSON for programmatic access
- ✅ **Timestamped** — Track when each scrape was last run

---

## What It Does

Each weekday after market close, the pipeline:

1. Discovers all listed companies across DSE market categories (A, B, G, N, Z)
2. Scrapes each company's page on [dsebd.org](https://www.dsebd.org) in parallel (10 threads)
3. Appends a new dated row to that company's sheet in `dse_companies.xlsx`
4. Updates the market-wide summary in `dse_market_summary.xlsx`
5. Commits all changed files back to this repo automatically

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
├── dse_companies.json                  # JSON backup of latest scrape
├── last_run.txt                        # Timestamp of the most recent successful run
│
├── requirements.txt                    # Python dependencies
└── .github/
    └── workflows/
        └── dse_daily_scrape.yml        # GitHub Actions workflow (runs Mon–Fri)
```

---

## Data Collected Per Company

Each daily row in `dse_companies.xlsx` captures:

| Field | Description |
|---|---|
| Date | Scrape date |
| Opening Price | Price at market open |
| Market Cap (mn) | Market capitalisation in BDT millions |
| 52W High / Low | 52-week price range |
| Trailing P/E | Price-to-earnings ratio |
| EPS Basic | Latest interim earnings per share |
| NAV Per Share | Net asset value from latest audited accounts |
| Cash Dividend | Most recent cash dividend declared |
| Bonus Dividend | Most recent bonus stock dividend |
| Paid-up Capital (mn) | Total paid-up capital |
| Authorized Capital (mn) | Total authorised capital |
| Sponsor/Dir % | Sponsor & director shareholding % |
| Institute % | Institutional shareholding % |
| Public % | Public shareholding % |

Static company info (name, sector, address, listing year, etc.) is written once to the top of each sheet and preserved across runs.

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
python script.py
```

This will scrape all DSE-listed companies and update the local Excel files.

To scrape a single company only:

```bash
python dse_scraper_individual_company.py
```

### Configuration

At the top of `dse_scraper_individual_company.py`:

| Variable | Default | Description |
|---|---|---|
| `MAX_WORKERS` | `10` | Number of parallel scraping threads |
| `REQUEST_DELAY` | `0.5s` | Per-thread delay between requests |
| `CATEGORIES` | `A,B,G,N,Z` | DSE market categories to scrape |
| `DSE_OUTPUT_EXCEL` | env var | Override output path for Excel file |
| `DSE_OUTPUT_JSON` | env var | Override output path for JSON file |

---

## Automated Daily Updates (GitHub Actions)

The workflow in `.github/workflows/dse_daily_scrape.yml` runs automatically:

- **Schedule:** Monday–Friday at **12:30 UTC** (18:30 BDT — after DSE closes at 14:30 BDT)
- **What it does:** Checks out the repo → runs the scraper → commits updated Excel/JSON files back
- **Manual trigger:** Go to **Actions → DSE Daily Scrape → Run workflow** to trigger on demand

No secrets or tokens need to be configured — the workflow uses the built-in `GITHUB_TOKEN`.

---

## Output Files

### `dse_companies.xlsx`
One sheet per listed company (named by trading code). Each sheet has:
- **Rows 1–10:** Static company info (name, sector, address, etc.)
- **Row 12:** Column headers
- **Row 13+:** One row per trading day, newest at the bottom

### `dse_market_summary.xlsx`
Aggregated view across all companies — useful for sector-level or market-wide analysis.

### `dse_companies.json`
Flat JSON snapshot of the most recent scrape. Useful for programmatic access without parsing Excel.

---

## Dependencies

```
requests
beautifulsoup4
openpyxl
lxml
```

Install with:
```bash
pip install -r requirements.txt
```

---

## 🔧 Troubleshooting

### Excel files not updating?
- Check `last_run.txt` to see when the scraper last ran successfully
- Verify your internet connection and that [dsebd.org](https://www.dsebd.org) is accessible
- Run `python script.py` manually to check for error messages

### GitHub Actions workflow not running?
- Go to **Actions → DSE Daily Scrape** and check the workflow logs
- Ensure the workflow file is enabled in repository settings
- Manually trigger it: **Actions → DSE Daily Scrape → Run workflow**

### ModuleNotFoundError?
- Ensure you've run `pip install -r requirements.txt`
- Try upgrading pip: `pip install --upgrade pip`
- Consider using a virtual environment: `python -m venv venv && source venv/bin/activate`

### Missing or empty data in Excel?
- Some companies may not report all fields (e.g., EPS for newly listed companies)
- DSE market may be closed (weekends/holidays)
- The column may not exist on [dsebd.org](https://www.dsebd.org) for that company

---

## 📊 Use Cases

- **Portfolio Tracking** — Monitor holdings' price movements and key metrics over time
- **Sector Analysis** — Compare performance across different DSE sectors
- **Technical Analysis** — Use historical price data for charting and trend analysis
- **Fundamental Research** — Track P/E, NAV, dividend yields, and shareholding patterns
- **Data Science Projects** — Feed DSE data into machine learning models

---

## Contributing

Contributions are welcome! Here's how to help:

1. **Fork** the repository
2. **Create a feature branch** (`git checkout -b feature/your-feature`)
3. **Commit your changes** (`git commit -m 'Add your feature'`)
4. **Push to the branch** (`git push origin feature/your-feature`)
5. **Open a Pull Request** with a clear description

### Ideas for Contributions
- Add support for technical indicators (MA, RSI, MACD, etc.)
- Improve scraping performance (caching, session pooling)
- Add data visualization/dashboard features
- Extend to other stock exchanges
- Improve error handling and resilience

---

## 📝 Notes

- This tool scrapes publicly available data from [dsebd.org](https://www.dsebd.org). Please use responsibly and avoid running at very high concurrency.
- DSE is closed on weekends and Bangladeshi public holidays. The workflow runs Mon–Fri but will produce an empty diff on holidays — nothing will be committed.
- Data quality depends on what DSE publishes. Missing fields (e.g. EPS for newly listed companies) will appear as empty cells.

---

## 📄 License

This project is for personal research and educational use. Data sourced from the Dhaka Stock Exchange website.