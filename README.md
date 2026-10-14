# SPC Tenders Extraction Pipeline

[![Daily SPC Tenders Synchronization](https://github.com/budd9442/spc-tenders/actions/workflows/daily_sync.yml/badge.svg)](https://github.com/budd9442/spc-tenders/actions/workflows/daily_sync.yml)
[![Latest Release](https://img.shields.io/github/v/release/budd9442/spc-tenders?color=blue)](https://github.com/budd9442/spc-tenders/releases/latest)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

## About

Daily extraction pipeline for Sri Lanka State Pharmaceuticals Corporation (SPC) and Directorate of Healthcare Services (DHS) tenders. Automatically extracts Schedule of Requirements tables from PDF tenders into Excel and CSV datasets.

## Releases

- [Latest Release](https://github.com/budd9442/spc-tenders/releases/latest): Contains newly detected tenders from daily runs.
- [Initial Archive (v2026.10.14)](https://github.com/budd9442/spc-tenders/releases/tag/v2026.10.14): Complete historical dataset covering 569 tenders and 4,866 line items.
  - `tables_xlsx.zip`: Individual Excel files for all 569 tenders.
  - `all_tenders_consolidated.xlsx`: Multi-sheet consolidated workbook.
  - `extracted_items.csv`: Flat CSV containing all extracted line items.

## How It Works

1. A GitHub Actions workflow runs daily at 06:00 UTC.
2. It queries active tenders on `spc.lk` and compares them against `data/manifest.json`.
3. If no new tenders exist, it exits without making changes.
4. If new tenders are found:
   - Downloads the tender PDFs.
   - Extracts item tables using PyMuPDF.
   - Generates individual `.xlsx` files and packs them into `new_tenders_<DATE>.zip`.
   - Updates `all_tenders_consolidated.xlsx` and `extracted_items.csv`.
   - Commits updated tracking data to `main`.
   - Publishes a new GitHub Release with the zip and consolidated files.

## Local Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# Check for new tenders without downloading
python run_sync.py --dry-run

# Run full sync
python run_sync.py
```

## License

MIT
