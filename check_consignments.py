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

# ── Configuration ────────────────────────────────────────────────────────────

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

# Consigner name fragments (lowercase) → gallery key
CONSIGNER_MAP = {
    "lougher":   "lougher",
    "clifton":   "clifton",
    "bents":     "bents",
    "fils":      "fils",
    "fontaine":  "fontaine",
    "poligrafa": "poligrafa",
}

# Which price column to use per gallery
PRICE_COLUMN = {
    "lougher":   "net",
    "clifton":   "net",
    "bents":     "net",
    "fils":      "retail",
    "poligrafa": "retail",
    "fontaine":  None,   # prices not compared
}

# ── Helpers ───────────────────────────────────────────────────────────────────

def extract_url(cell_value):
    """Pull the first http URL out of a cell that may contain other text."""
    if not cell_value:
        return None
    text = str(cell_value)
    match = re.search(r'https?://\S+', text)
    return match.group(0).strip() if match else None


def parse_price(raw):
    """Turn '£1,234.00' or '1234 GBP' or '1,234' into a float, or None."""
    if raw is None:
        return None
    text = re.sub(r'[^\d.]', '', str(raw))
    try:
        return float(text) if text else None
    except ValueError:
        return None


def fetch_soup(url):
    """Fetch a page and return a BeautifulSoup object, or None on error."""
    try:
        r = requests.get(url, headers=HEADERS, timeout=20)
        r.raise_for_status()
        return BeautifulSoup(r.text, "lxml")
    except Exception:
        return None


def load_tracked_artists():
    """Load artist names from the Google Sheet CSV export URL."""
    artists = set()
    try:
        with urllib.request.urlopen(ARTISTS_SHEET_URL, timeout=15) as resp:
            lines = resp.read().decode("utf-8").splitlines()
        reader = csv.DictReader(lines)
        for row in reader:
            name = row.get("Artist", "").strip()
            if name:
                artists.add(name.lower())
    except Exception:
        pass
    return artists


def load_consignments():
    """Read the xlsx and return a list of dicts."""
    path = os.path.join("data", next(
        f for f in os.listdir("data") if f.endswith(".xlsx")
    ))
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    headers = [str(c.value).strip() if c.value else "" for c in ws[1]]

    def col(row, name):
        try:
            return row[headers.index(name)].value
        except (ValueError, IndexError):
            return None

    items = []
    for row in ws.iter_rows(min_row=2):
        consigner_raw = col(row, "Consigner") or ""
        gallery_key = None
        for fragment, key in CONSIGNER_MAP.items():
            if fragment in consigner_raw.lower():
                gallery_key = key
                break
        if not gallery_key:
            continue

        source_url = extract_url(col(row, "Internal Comment / source webste"))
        admin_link = col(row, "Admin Link")

        retail_raw = col(row, "Retail Price")
        net_raw    = col(row, "Net Price")

        price_col  = PRICE_COLUMN[gallery_key]
        our_price  = parse_price(net_raw if price_col == "net" else retail_raw)

        items.append({
            "id":          col(row, "Inventory Id"),
            "artist":      col(row, "Artist") or "",
            "title":       col(row, "Title") or "",
            "consigner":   consigner_raw,
            "gallery":     gallery_key,
            "source_url":  source_url,
            "admin_link":  admin_link,
            "our_price":   our_price,
        })
    return items


# ── Per-gallery scrapers ──────────────────────────────────────────────────────

def scrape_lougher(url):
    """Returns (found: bool, price: float|None)"""
    soup = fetch_soup(url)
    if soup is None:
        return False, None
    # Lougher uses Shopify — product JSON is embedded
    script = soup.find("script", string=re.compile(r'"price"'))
    if script:
        m = re.search(r'"price"\s*:\s*(\d+)', script.string)
        if m:
            # Shopify stores price in pence
            return True, int(m.group(1)) / 100
    # Fallback: look for sold-out indicator
    sold = soup.find(string=re.compile(r'sold out', re.I))
    if sold:
        return False, None
    # If page loads and no price found, mark as found with unknown price
    return True, None


def scrape_clifton(url):
    soup = fetch_soup(url)
    if soup is None:
        return False, None
    sold = soup.find(string=re.compile(r'sold out|not available', re.I))
    if sold:
        return False, None
    price_tag = soup.find(class_=re.compile(r'price', re.I))
    price = parse_price(price_tag.get_text()) if price_tag else None
    return True, price


def scrape_bents(url):
    soup = fetch_soup(url)
    if soup is None:
        return False, None
    sold = soup.find(string=re.compile(r'sold|not available', re.I))
    if sold:
        return False, None
    price_tag = soup.find(class_=re.compile(r'price', re.I))
    price = parse_price(price_tag.get_text()) if price_tag else None
    return True, price


def scrape_fils(url):
    soup = fetch_soup(url)
    if soup is None:
        return False, None
    sold = soup.find(string=re.compile(r'sold|vergriffen|not available', re.I))
    if sold:
        return False, None
    price_tag = soup.find(class_=re.compile(r'price', re.I))
    price = parse_price(price_tag.get_text()) if price_tag else None
    return True, price


def scrape_fontaine(url):
    """Fontaine pages are per-artist. We just check the artwork title exists."""
    soup = fetch_soup(url)
    if soup is None:
        return False, None
    return True, None


def scrape_poligrafa(url, title):
    """Poligrafa pages are per-artist. Check if the artwork title appears."""
    soup = fetch_soup(url)
    if soup is None:
        return False, None
    if title and title.lower() in soup.get_text().lower():
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

def check_new_arrivals(tracked_artists):
    """
    For each gallery, fetch their main listings and look for tracked artists
    not already in our consignment list. Returns list of dicts.
    """
    # For now only Lougher is implemented
    arrivals = []
    arrivals += _lougher_arrivals(tracked_artists)
    return arrivals


def _lougher_arrivals(tracked_artists):
    found = []
    url = "https://www.loughercontemporary.com/collections/all"
    soup = fetch_soup(url)
    if not soup:
        return found
    for a in soup.select("a.card__heading, a[href*='/products/']"):
        text = a.get_text(strip=True).lower()
        for artist in tracked_artists:
            if artist in text:
                found.append({
                    "artist": artist.title(),
                    "title":  a.get_text(strip=True),
                    "gallery": "Lougher",
                    "url": "https://www.loughercontemporary.com" + a["href"]
                    if a.get("href", "").startswith("/") else a.get("href", ""),
                })
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
    parts_summary = []
    if sold:        parts_summary.append(f"{len(sold)} sold/removed")
    if mismatches:  parts_summary.append(f"{len(mismatches)} price")
    if arrivals:    parts_summary.append(f"{len(arrivals)} new")
    summary_str = ", ".join(parts_summary) if parts_summary else "no changes"

    subject = f"Consignment Watch — {today.strftime('%-d %b %Y')} — {summary_str}"

    def link(url, text):
        if url:
            return f'<a href="{url}">{text}</a>'
        return text

    rows_sold = "".join(
        f"<li>{i['artist']} &ldquo;{i['title']}&rdquo; at {i['consigner']} "
        f"[{link(i['source_url'], 'gallery')} | {link(i['admin_link'], 'admin')}]</li>"
        for i in sold
    )
    rows_mismatch = "".join(
        f"<li>{i['artist']} &ldquo;{i['title']}&rdquo; at {i['consigner']}: "
        f"our {i['our_price']} vs theirs {i['their_price']} "
        f"[{link(i['source_url'], 'gallery')} | {link(i['admin_link'], 'admin')}]</li>"
        for i in mismatches
    )
    rows_arrivals = "".join(
        f"<li>{a['artist']} &mdash; {link(a['url'], a['title'])} at {a['gallery']}</li>"
        for a in arrivals
    )

    html = f"""
    <p>Good morning Kris and Nana.</p>
    <p>Daily consignment report &mdash; {today.strftime('%-d %B %Y')}.</p>
    <p><strong>Summary:</strong> {total_checked} consigned artworks checked.
    {len(sold)} sold/removed &bull; {len(mismatches)} price mismatches &bull; {len(arrivals)} new arrivals.</p>

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
            if pct > 0.02:   # more than 2% difference
                item["their_price"] = their_price
                mismatches.append(item)

    arrivals = check_new_arrivals(tracked_artists)

    subject, html = build_email(today, sold, mismatches, arrivals, len(consignments))
    send_email(subject, html)
    print(f"Email sent: {subject}")


if __name__ == "__main__":
    main()
