import os
import re
import csv
import smtplib
import urllib.request
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

# ── Configuration ─────────────────────────────────────────────────────────────

GMAIL_SENDER       = os.environ["GMAIL_SENDER"]
GMAIL_APP_PASSWORD = os.environ["GMAIL_APP_PASSWORD"]
RECIPIENT_EMAIL    = os.environ["RECIPIENT_EMAIL"]
ARTISTS_SHEET_URL  = os.environ["ARTISTS_SHEET_URL"]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}

CONSIGNER_MAP = {
    "lougher":   "lougher",
    "clifton":   "clifton",
    "bents":     "bents",
    "fils":      "fils",
    "fontaine":  "fontaine",
    "poligrafa": "poligrafa",
}

PRICE_COLUMN = {
    "lougher":   "net",
    "clifton":   "net",
    "bents":     "net",
    "fils":      "retail",
    "poligrafa": "retail",
    "fontaine":  None,
}

ACTIVE_GALLERIES = {"lougher", "clifton", "bents", "fils", "fontaine", "poligrafa"}

# ── Helpers ───────────────────────────────────────────────────────────────────

def extract_url(cell_value):
    if not cell_value:
        return None
    text = str(cell_value)
    match = re.search(r'https?://\S+', text)
    return match.group(0).strip() if match else None


def parse_price(raw):
    """Parse price string to float, handling both English and European formats."""
    if raw is None:
        return None
    text = str(raw).strip()
    # European format: 7.200,00 → 7200.00
    if re.search(r'\d+\.\d{3},\d{2}', text):
        text = text.replace('.', '').replace(',', '.')
    elif re.search(r'\d+,\d{2}', text):
        text = text.replace(',', '.')
    text = re.sub(r'[^\d.]', '', text)
    try:
        return float(text) if text else None
    except ValueError:
        return None


def fetch_page(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=20, allow_redirects=True)
        soup = BeautifulSoup(r.text, "lxml")
        return r.status_code, soup, r.text
    except Exception:
        return None, None, None


def fetch_soup(url):
    status, soup, _ = fetch_page(url)
    if status == 200:
        return soup
    return None


def url_handle(url):
    """Extract just the product handle from a Lougher URL for comparison."""
    if not url:
        return None
    # Get last non-empty path segment, strip query params
    path = url.lower().split("?")[0].rstrip("/")
    return path.split("/")[-1]


def load_tracked_artists():
    artists = set()
    try:
        with urllib.request.urlopen(ARTISTS_SHEET_URL, timeout=15) as resp:
            lines = resp.read().decode("utf-8").splitlines()
        reader = csv.DictReader(lines)
        for row in reader:
            name = row.get("Artist", "").strip()
            if name:
                artists.add(name.lower())
    except Exception as e:
        print(f"DEBUG: failed to load artists sheet: {e}")
    print(f"DEBUG: loaded {len(artists)} tracked artists")
    return artists


def load_consignments():
    path = os.path.join("data", next(
        f for f in os.listdir("data") if f.endswith(".xlsx")
    ))
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    headers = [str(c.value).strip() if c.value else "" for c in ws[2]]

    def col(row, name):
        try:
            return row[headers.index(name)].value
        except (ValueError, IndexError):
            return None

    items = []
    for row in ws.iter_rows(min_row=3):
        consigner_raw = col(row, "Consigner") or ""
        gallery_key = None
        for fragment, key in CONSIGNER_MAP.items():
            if fragment in consigner_raw.lower():
                gallery_key = key
                break
        if not gallery_key:
            continue

        source_url = extract_url(col(row, "Internal Comment"))
        admin_link = col(row, "Admin Link")
        retail_raw = col(row, "Retail")
        net_raw    = col(row, "Net Price")
        price_col  = PRICE_COLUMN[gallery_key]
        our_price  = parse_price(net_raw if price_col == "net" else retail_raw)

        items.append({
            "id":         col(row, "Inventory Id"),
            "artist":     col(row, "Artist") or "",
            "title":      col(row, "Title") or "",
            "consigner":  consigner_raw,
            "gallery":    gallery_key,
            "source_url": source_url,
            "admin_link": admin_link,
            "our_price":  our_price,
        })
    return items


# ── Scrapers ──────────────────────────────────────────────────────────────────

def scrape_lougher(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=20)
        if r.status_code == 404:
            return False, None
        if r.status_code != 200:
            return None, None
        soup = BeautifulSoup(r.text, "lxml")
        for script in soup.find_all("script"):
            if not script.string:
                continue
            if '"price"' in script.string:
                m = re.search(r'"price"\s*:\s*(\d+)', script.string)
                if m:
                    price = int(m.group(1)) / 100
                    if price > 0:
                        return True, price
        return True, None
    except Exception:
        return None, None


def scrape_clifton(url):
    status, soup, _ = fetch_page(url)
    if status is None:
        return None, None
    if status == 404:
        return False, None
    if soup is None:
        return None, None
    # Check for "Sold" text
    page_text = soup.get_text()
    if re.search(r'\bSold\b', page_text):
        return False, None
    # Extract price — look for £X,XXX.XX pattern
    price_match = re.search(r'£([\d,]+\.?\d*)', page_text)
    if price_match:
        return True, parse_price(price_match.group(1))
    return True, None


def scrape_bents(url):
    status, soup, _ = fetch_page(url)
    if status is None:
        return None, None
    if status == 404:
        return False, None
    if soup is None:
        return None, None
    page_text = soup.get_text()
    if re.search(r'\bSold\b', page_text):
        return False, None
    price_match = re.search(r'£([\d,]+\.?\d*)', page_text)
    if price_match:
        return True, parse_price(price_match.group(1))
    return True, None


def scrape_fils(url):
    status, soup, raw_text = fetch_page(url)
    if status is None:
        return None, None
    if status == 404:
        return False, None
    if soup is None:
        return None, None
    # Check we didn't get redirected to home page
    if not soup.find(string=re.compile(r'In den Warenkorb|Warenkorb|inkl.*MwSt', re.I)):
        return None, None
    # Extract European price format: 7.200,00 €
    price_match = re.search(r'([\d]+(?:\.\d{3})*,\d{2})\s*€', soup.get_text())
    if price_match:
        return True, parse_price(price_match.group(1))
    return True, None


def scrape_fontaine(url):
    status, soup, _ = fetch_page(url)
    if status is None:
        return None, None
    if status == 404:
        return False, None
    if soup is None:
        return None, None
    return True, None


def scrape_poligrafa(url, title):
    status, soup, raw_text = fetch_page(url)
    if status is None:
        return None, None
    if status == 404:
        return False, None
    if not raw_text:
        return None, None
    # Title is in page source even though displayed as popup
    if title and title.lower() in raw_text.lower():
        # Try to find price near the title
        price_match = re.search(r'([\d]+(?:[.,]\d+)?)\s*€', raw_text)
        if price_match:
            return True, parse_price(price_match.group(1))
        return True, None
    return False, None


SCRAPERS = {
    "lougher":   scrape_lougher,
    "clifton":   scrape_clifton,
    "bents":     scrape_bents,
    "fils":      scrape_fils,
    "fontaine":  scrape_fontaine,
    "poligrafa": scrape_poligrafa,
}


def scrape(item):
    url = item["source_url"]
    if not url:
        return None, None
    gallery = item["gallery"]
    scraper = SCRAPERS[gallery]
    if gallery == "poligrafa":
        return scraper(url, item["title"])
    return scraper(url)


# ── New arrivals ──────────────────────────────────────────────────────────────

def check_new_arrivals(tracked_artists, consignments):
    already_have_handles = set()
    for item in consignments:
        if item["gallery"] == "lougher" and item["source_url"]:
            h = url_handle(item["source_url"])
            if h:
                already_have_handles.add(h)
    print(f"DEBUG: {len(already_have_handles)} Lougher handles in inventory")
    arrivals = []
    arrivals += _lougher_arrivals(tracked_artists, already_have_handles)
    return arrivals


def _lougher_arrivals(tracked_artists, already_have_handles):
    found = []
    seen_handles = set()
    base_url = "https://www.loughercontemporary.com/collections/all?filter.v.availability=1&filter.p.vendor=Lougher+Contemporary&sort_by=created-descending"
    page = 1
    while True:
        url = base_url if page == 1 else f"{base_url}&page={page}"
        soup = fetch_soup(url)
        if not soup:
            print(f"DEBUG: could not fetch Lougher page {page}")
            break
        links = soup.select("a[href*='/products/']")
        print(f"DEBUG Lougher page {page}: {len(links)} product links")
        if not links:
            break
        new_on_page = 0
        for a in links:
            href = a.get("href", "")
            if not href:
                continue
            full_url = ("https://www.loughercontemporary.com" + href
                        if href.startswith("/") else href)
            handle = url_handle(full_url)
            if not handle:
                continue
            if handle in already_have_handles:
                continue
            if handle in seen_handles:
                continue
            text = a.get_text(strip=True).lower()
            if not text:
                continue
            for artist in tracked_artists:
                if artist in text:
                    seen_handles.add(handle)
                    new_on_page += 1
                    found.append({
                        "artist":  artist.title(),
                        "title":   a.get_text(strip=True),
                        "gallery": "Lougher",
                        "url":     full_url,
                    })
                    break
        print(f"DEBUG Lougher page {page}: {new_on_page} new arrivals")
        next_page = soup.select_one("a[href*='page=']")
        if not next_page:
            break
        page += 1
        if page > 10:
            break
    print(f"DEBUG: total Lougher new arrivals: {len(found)}")
    return found


# ── Email ─────────────────────────────────────────────────────────────────────

def send_email(subject, html_body):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = GMAIL_SENDER
    msg["To"]      = RECIPIENT_EMAIL
    msg.attach(MIMEText(html_body, "html"))
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_SENDER, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_SENDER, RECIPIENT_EMAIL, msg.as_string())


def build_email(today, sold, mismatches, arrivals, total_checked):
    parts = []
    if sold:       parts.append(f"{len(sold)} sold/removed")
    if mismatches: parts.append(f"{len(mismatches)} price")
    if arrivals:   parts.append(f"{len(arrivals)} new")
    summary_str = ", ".join(parts) if parts else "no changes"
    subject = f"Consignment Watch — {today.strftime('%-d %b %Y')} — {summary_str}"

    def link(url, text):
        return f'<a href="{url}">{text}</a>' if url else text

    def group_and_sort(items):
        groups = {}
        for i in items:
            key = i["consigner"]
            groups.setdefault(key, []).append(i)
        # Sort items within each group by artist name
        for key in groups:
            groups[key].sort(key=lambda x: x.get("artist", "").lower())
        return groups

    sold_groups = group_and_sort(sold)
    rows_sold = ""
    for consigner in sorted(sold_groups.keys()):
        rows_sold += f"<li><strong>{consigner}</strong><ul>"
        for i in sold_groups[consigner]:
            rows_sold += (
                f"<li>{i['artist']} &ldquo;{i['title']}&rdquo; "
                f"[{link(i['source_url'], 'gallery')} | {link(i['admin_link'], 'admin')}]</li>"
            )
        rows_sold += "</ul></li>"

    mismatch_groups = group_and_sort(mismatches)
    rows_mismatch = ""
    for consigner in sorted(mismatch_groups.keys()):
        rows_mismatch += f"<li><strong>{consigner}</strong><ul>"
        for i in mismatch_groups[consigner]:
            rows_mismatch += (
                f"<li>{i['artist']} &ldquo;{i['title']}&rdquo;: "
                f"our {i['our_price']} vs theirs {i['their_price']} "
                f"[{link(i['source_url'], 'gallery')} | {link(i['admin_link'], 'admin')}]</li>"
            )
        rows_mismatch += "</ul></li>"

    arrival_groups = {}
    for a in arrivals:
        arrival_groups.setdefault(a["gallery"], []).append(a)
    for key in arrival_groups:
        arrival_groups[key].sort(key=lambda x: x.get("artist", "").lower())
    rows_arrivals = ""
    for gallery in sorted(arrival_groups.keys()):
        rows_arrivals += f"<li><strong>{gallery}</strong><ul>"
        for a in arrival_groups[gallery]:
            rows_arrivals += f"<li>{a['artist']} &mdash; {link(a['url'], a['title'])}</li>"
        rows_arrivals += "</ul></li>"

    html = f"""
    <p>Good morning Kris and Nana.</p>
    <p>Daily consignment report &mdash; {today.strftime('%-d %B %Y')}.</p>
    <p><strong>Summary:</strong> {total_checked} consigned artworks checked.
    {len(sold)} sold/removed &bull; {len(mismatches)} price mismatches &bull;
    {len(arrivals)} new arrivals.</p>

    <h3>🔴 Sold or removed ({len(sold)})</h3>
    <ul>{rows_sold if sold else "<li>None</li>"}</ul>

    <h3>🟡 Price mismatches ({len(mismatches)})</h3>
    <ul>{rows_mismatch if mismatches else "<li>None</li>"}</ul>

    <h3>🟢 New arrivals ({len(arrivals)})</h3>
    <ul>{rows_arrivals if arrivals else "<li>None</li>"}</ul>
    """
    return subject, html


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    today = date.today()
    consignments = load_consignments()
    tracked_artists = load_tracked_artists()

    sold       = []
    mismatches = []

    for item in consignments:
        if item["gallery"] not in ACTIVE_GALLERIES:
            continue
        if not item["source_url"]:
            continue
        found, their_price = scrape(item)
        if found is None:
            continue
        if not found:
            sold.append(item)
        elif item["our_price"] and their_price:
            diff = abs(item["our_price"] - their_price)
            pct  = diff / item["our_price"]
            if pct > 0.02:
                item["their_price"] = their_price
                mismatches.append(item)

    arrivals = check_new_arrivals(tracked_artists, consignments)
    subject, html = build_email(today, sold, mismatches, arrivals, len(consignments))
    send_email(subject, html)
    print(f"Email sent: {subject}")


if __name__ == "__main__":
    main()
