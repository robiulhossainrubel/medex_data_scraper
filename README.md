# MedEx Bangladesh Medicine Scraper

A resilient web scraper that extracts detailed medicine information from **[medex.com.bd](https://medex.com.bd/brands)** (all pages) and saves it to an Excel file.  
It is built with **auto-resume**, **auto-sync** for temporary files, and **anti-block** measures (random delays + rotating User‑Agents).

## Features

- 📋 **Comprehensive data** – Brand name, type, dose, generic, company, price (unit/strip/pack size), and all available clinical sections (indications, pharmacology, dosage, side effects, etc.).
- 🔄 **Resume from interruption** – A `scraper_progress.txt` file remembers the exact page and item; restarting continues right where it left off.
- 🛡️ **Anti‑block** – Randomised delays between requests and a list of rotating User‑Agent headers to mimic human behaviour.
- 💾 **Safe saving** – If the main Excel file is locked (e.g. open in Excel), data is temporarily written to `_temp.xlsx` (or `_temp_1.xlsx`, etc.) so **no data is ever lost**.
- ♻️ **Auto‑sync** – On restart, all temporary files are automatically merged into the main Excel file and cleaned up.
- 🧮 **Batch writing** – Saves every 5 records (configurable) to minimise memory usage and allow progress tracking.

## Requirements

- Python 3.7+
- Install required libraries:

```bash
pip install requests beautifulsoup4 openpyxl
