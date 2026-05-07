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

# ── Helpers ───────────────────────────────────────────────────────────────────

def extract_url(cell_value):
    if not cell_value:
        return None
    text = str(cell_value)
    match = re.search(r'https?://\S+', text)
    return match.group(0).strip() if match else None


def parse_price(raw):
    if raw is None:
        return None
    text = re.sub(r'[^\d.]', '', str(raw))
    try:
        return float(text) if text else None
    except ValueError:
        return None


def fetch_soup(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=20)
        r.raise_for_status()
        return BeautifulSoup(r.text, "lxml")
    except Exception:
        return None


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
    except Exception:
        pass
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

        source_url = extract_url(col(row, "Internal Comment / source website"))
        admin_link = col(row, "Admin Link")
        retail_raw = col(row, "Retail Price")
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
    soup = fetch_soup(url)
    if soup is None:
        return False, None
    sold = soup.find(string=re.compile(r'sold out', re.I))
    if sold:
        return False, None
    script = soup.find("script", string=re.compile(r'"price"'))
    if script:
        m = re.search(r'"price"\s*:\s*(\d+)', script.string)
        if m:
            return True, int(m.group(1)) / 100
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
    soup = fetch_soup(url)
    if soup is None:
        return False, None
    return True, None


def scrape_poligrafa(url, title):
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

def check_new_arrivals(tracked_artists, consignments):
    # Build a set of (artist_lower, url) pairs we already consign from Lougher
    already_have = set()
    for item in consignments:
        if item["gallery"] == "lougher" and item["source_url"]:
            already_have.add(item["source_url"].lower().split("?")[0].rstrip("/"))

    arrivals = []
    arrivals += _lougher_arrivals(tracked_artists, already_have)
    return arrivals


def _lougher_arrivals(tracked_artists, already_have):
    found = []
    url = "https://www.loughercontemporary.com/collections/all"
    soup = fetch_soup(url)
    if not soup:
        return found
    seen_urls = set()
    for a in soup.select("a.card__heading, a[href*='/products/']"):
        text = a.get_text(strip=True).lower()
        href = a.get("href", "")
        full_url = ("https://www.loughercontemporary.com" + href
                    if href.startswith("/") else href)
        clean_url = full_url.lower().split("?")[0].rstrip("/")
        if clean_url in already_have:
            continue
        if clean_url in seen_urls:
            continue
        # Check if this is a consignment sale page
        product_soup = fetch_soup(full_url)
        if product_soup:
            page_text = product_soup.get_text()
            if any(phrase in page_text for phrase in [
                "Consignment sales",
                "Collect works sold by our trusted network",
                "What is a Consignment work"
            ]):
                continue
        for artist in tracked_artists:
            if artist in text:
                seen_urls.add(clean_url)
                found.append({
                    "artist":  artist.title(),
                    "title":   a.get_text(strip=True),
                    "gallery": "Lougher",
                    "url":     full_url,
                })
                break
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
