# Barcelona Wine Week 2026 - Exhibitor Scraper
# Scrapes all exhibitor data from the BWW 2026 e-catalogue and writes to Excel

from playwright.sync_api import sync_playwright
from openpyxl import Workbook, load_workbook
import re
import requests


def clean_text(value):
    """Remove non-printable control characters from a string value."""
    if value is None:
        return ''
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', str(value))


# ---------------------------------------------------------------------------
# Step 1: Collect all exhibitor IDs via the catalogue's internal search API
# ---------------------------------------------------------------------------

headers = {
    'Content-Type': 'application/json',
    'Accept': 'application/json, text/plain, */*',
    'Origin': 'https://ecatalogue.firabarcelona.com',
    'Referer': 'https://ecatalogue.firabarcelona.com/'
}

body = {
    "eventName": "",
    "sapCode": "J134026",
    "eventCode": "",
    "searchText": "",
    "brand": "",
    "selectedHierarchicalProperties": [],
    "selectedProperties": [],
    "selectedSectors": [],
    "selectedHomeOds": [],
    "selectedCountries": [],
    "filter": "ONLY_EXHIBITORS",
    "searchOrder": "BY_RELEVANCE",
    "language": "en_GB",
    "maxTextLength": 500,
    "selectedMultiEvents": [],
    "createCache": 0
}

all_ids = []
for page_num in range(2):
    response = requests.post(
        f'https://ecatalogueusearch-api.firabarcelona.com/v1/us/unifiedSearch?page={page_num}&size=1000&language=en_GB',
        json=body,
        headers=headers
    )
    data = response.json()
    ids = [item['entityId'] for item in data['list']]
    all_ids.extend(ids)

print(f'Total IDs collected: {len(all_ids)}')


# ---------------------------------------------------------------------------
# Step 2: Build exhibitor detail page URLs from collected IDs
# ---------------------------------------------------------------------------

BASE_URL = 'https://ecatalogue.firabarcelona.com/163/barcelonawineweek2026/exhibitor/{}/detail?lang=en_GB'
urls = [BASE_URL.format(id) for id in all_ids]


# ---------------------------------------------------------------------------
# Step 3: Scrape each exhibitor page using Playwright
# The site is Angular-rendered, so requests alone cannot access the content
# ---------------------------------------------------------------------------

all_wineries = []

with sync_playwright() as p:
    browser = p.webkit.launch(headless=True)
    page = browser.new_page()

    for i, url in enumerate(urls):
        print(f'Scraping winery {i+1} of {len(urls)}...')
        try:
            page.goto(url, wait_until='domcontentloaded')
            page.wait_for_selector('.detail-content__title', timeout=4000)

            try:
                winery_title = page.locator('.detail-content__title').inner_text(timeout=4000)
            except:
                winery_title = ''

            try:
                winery_description = page.locator('.detail-content__description .description').inner_text(timeout=4000)
            except:
                winery_description = ''

            try:
                winery_website = page.locator('.detail-contact__item--underline .is-link').get_attribute('href', timeout=4000)
            except:
                winery_website = ''

            try:
                winery_location = page.locator('.detail-contact__item .text').inner_text(timeout=4000)
            except:
                winery_location = ''

            # Phone number is not in a dedicated element — extract via regex from contact items
            try:
                contact_items = page.locator('.detail-contact__item').all_inner_texts()
                phone = next((x for x in contact_items if re.search(r'\+?\d[\d\s]{7,}', x)), None)
            except:
                phone = ''

            try:
                tradeshow_location = page.locator('.detail-map__location').inner_text(timeout=4000)
            except:
                tradeshow_location = ''

            # Each exhibitor may list multiple products — iterate through all product cards
            products = []
            product_cards = page.locator('.card.card-custom').all()
            for product in product_cards:
                try:
                    product_name = product.locator('.ex__data-title').inner_text(timeout=4000)
                except:
                    product_name = ''
                try:
                    product_description = product.locator('.product__exhibitor-name.long-text-four-line').inner_text(timeout=4000)
                except:
                    product_description = ''
                products.append((product_name, product_description))

            all_wineries.append({
                'name': winery_title,
                'description': winery_description,
                'website': winery_website,
                'location': winery_location,
                'phone': phone,
                'tradeshow_location': tradeshow_location,
                'products': products
            })

        except Exception as e:
            print(f'Failed on {url}: {e}')
            all_wineries.append(None)  # Placeholder to maintain correct row count


# ---------------------------------------------------------------------------
# Step 4: Write all scraped data to a structured Excel file
# Product columns are generated dynamically based on the exhibitor with the
# highest number of products
# ---------------------------------------------------------------------------

max_products = max(len(w['products']) for w in all_wineries if w is not None)
print(f'Max products for any exhibitor: {max_products}')

wb = Workbook()
ws = wb.active
ws.title = 'Wineries'

# Build dynamic headers — fixed columns followed by one pair per product slot
fixed_headers = ['Winery Name', 'Winery Description', 'Website', 'Winery Location', 'Phone', 'Trade Show Location']
product_headers = []
for i in range(1, max_products + 1):
    product_headers.append(f'Product {i} Name')
    product_headers.append(f'Product {i} Description')

ws.append(fixed_headers + product_headers)

# Write one row per exhibitor
for winery in all_wineries:
    if winery is None:
        ws.append(['SCRAPE FAILED'])
        continue

    row = [
        clean_text(winery['name']),
        clean_text(winery['description']),
        clean_text(winery['website']),
        clean_text(winery['location']),
        clean_text(winery['phone']),
        clean_text(winery['tradeshow_location'])
    ]
    for name, description in winery['products']:
        row.append(clean_text(name))
        row.append(clean_text(description))
    ws.append(row)

# Auto-size columns for readability, capped at 50 characters wide
for column in ws.columns:
    max_length = 0
    column_letter = column[0].column_letter
    for cell in column:
        if cell.value:
            max_length = max(max_length, len(str(cell.value)))
    ws.column_dimensions[column_letter].width = min(max_length + 2, 50)

wb.save('Final_barcelona_wine_week_2026.xlsx')
print('Scraping complete. Data saved to Final_barcelona_wine_week_2026.xlsx')