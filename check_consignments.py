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

ACTIVE_GALLERIES = {"lougher"}

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
    already_have_urls = set()
    already_have_titles = set()
    for item in consignments:
        if item["gallery"] == "lougher":
            if item["source_url"]:
                already_have_urls.add(item["source_url"].lower().split("?")[0].rstrip("/"))
            if item["artist"] and item["title"]:
                key = (item["artist"].lower().strip(), item["title"].lower().strip())
                already_have_titles.add(key)
    print(f"DEBUG: {len(already_have_urls)} Lougher URLs, {len(already_have_titles)} Lougher titles in inventory")
    arrivals = []
    arrivals += _lougher_arrivals(tracked_artists, already_have_urls, already_have_titles)
    return arrivals


def _lougher_arrivals(tracked_artists, already_have_urls, already_have_titles):
    found = []
    seen_urls = set()
    base_url = "https://www.loughercontemporary.com/collections/all?filter.v.availability=1&filter.p.vendor=Lougher+Contemporary&sort_by=created-descending"
    page = 1
    while True:
        url = base_url if page == 1 else f"{base_url}&page={page}"
        soup = fetch_soup(url)
        if not soup:
            print(f"DEBUG: could not fetch page {page}")
            break
        links = soup.select("a[href*='/products/']")
        print(f"DEBUG page {page}: found {len(links)} product links")
        if not links:
            break
        new_on_page = 0
        for a in links:
            href = a.get("href", "")
            if not href:
                continue
            full_url = ("https://www.loughercontemporary.com" + href
                        if href.startswith("/") else href)
            clean_url = full_url.lower().split("?")[0].rstrip("/")
            if clean_url in already_have_urls:
                continue
            if clean_url in seen_urls:
                continue
            text = a.get_text(strip=True).lower()
            if not text:
                continue
            for artist in tracked_artists:
                if artist in text:
                    seen_urls.add(clean_url)
                    new_on_page += 1
                    found.append({
                        "artist":  artist.title(),
                        "title":   a.get_text(strip=True),
                        "gallery": "Lougher",
                        "url":     full_url,
                    })
                    break
        print(f"DEBUG page {page}: {new_on_page} new arrivals found")
        next_page = soup.select_one("a[href*='page=']")
        if not next_page:
            break
        page += 1
        if page > 10:
            break
    print(f"DEBUG: total new arrivals: {len(found)}")
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

    def group_by_consigner(items):
        groups = {}
        for i in items:
            key = i["consigner"]
            groups.setdefault(key, []).append(i)
        return groups

    sold_groups = group_by_consigner(sold)
    rows_sold = ""
    for consigner, items in sorted(sold_groups.items()):
        rows_sold += f"<li><strong>{consigner}</strong><ul>"
        for i in items:
            rows_sold += (
                f"<li>{i['artist']} &ldquo;{i['title']}&rdquo; "
                f"[{link(i['source_url'], 'gallery')} | {link(i['admin_link'], 'admin')}]</li>"
            )
        rows_sold += "</ul></li>"

    mismatch_groups = group_by_consigner(mismatches)
    rows_mismatch = ""
    for consigner, items in sorted(mismatch_groups.items()):
        rows_mismatch += f"<li><strong>{consigner}</strong><ul>"
        for i in items:
            rows_mismatch += (
                f"<li>{i['artist']} &ldquo;{i['title']}&rdquo;: "
                f"our {i['our_price']} vs theirs {i['their_price']} "
                f"[{link(i['source_url'], 'gallery')} | {link(i['admin_link'], 'admin')}]</li>"
            )
        rows_mismatch += "</ul></li>"

    arrival_groups = {}
    for a in arrivals:
        arrival_groups.setdefault(a["gallery"], []).append(a)
    rows_arrivals = ""
    for gallery, items in sorted(arrival_groups.items()):
        rows_arrivals += f"<li><strong>{gallery}</strong><ul>"
        for a in items:
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
