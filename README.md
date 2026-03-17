# Barcelona Wine Week 2026 - Exhibitor Scraper

A Python scraper that extracts data on all 1355+ exhibitors from the 
Barcelona Wine Week 2026 e-catalogue and writes it to a structured Excel file.

## What it does

- Queries the catalogue's internal API to collect all exhibitor IDs
- Visits each exhibitor's page using Playwright to extract:
  - Winery name and description
  - Website, phone number and location
  - Trade show floor location
  - All listed products (name and description)
- Writes everything to a clean, auto-formatted Excel file with dynamic 
  product columns

## Why Playwright

The Barcelona Wine Week catalogue is built with Angular, meaning the page 
content is rendered by JavaScript at runtime. Standard HTTP requests return 
an empty shell. Playwright launches a real browser to load each page fully 
before extracting data.

## Output

A single Excel file — `Final_barcelona_wine_week_2026.xlsx` — with one row 
per exhibitor and dynamically generated product columns based on the 
exhibitor with the most products listed.

## Requirements
```
pip install playwright openpyxl requests
playwright install webkit
```

## Usage

Run the script directly:
```
python barcelona_wine_week_FINAL.py
```

The script will print progress as it scrapes each exhibitor and save the 
Excel file to the current directory on completion.
