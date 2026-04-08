#!/usr/bin/env python3
"""
Enrich company data with phone numbers, emails, and real product info
by scraping company websites and Polish business directories.

Sources (in order of priority):
  1. Company website (from GUS/KRS data) — scrape contact page + products
  2. Panorama Firm (panoramafirm.pl) — free business directory
  3. PKT.pl — Polish Yellow Pages
  4. Google Maps / DuckDuckGo search fallback

Input: Excel file from lookup scripts (needs columns: api_name, website, nip)
Output: Enriched Excel with phone, email, products, social media

Usage:
    python enrich_contacts.py results.xlsx --output enriched.xlsx
    python enrich_contacts.py results.xlsx --limit 5
"""

import argparse
import logging
import re
import sys
import time

from dotenv import load_dotenv
import pandas as pd
import requests
import urllib3

load_dotenv()

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept-Language": "pl-PL,pl;q=0.9,en;q=0.8",
}


# ─── Phone / email extraction from HTML ─────────────────────────────────────

# Polish phone patterns: +48 XXX XXX XXX, 48XXXXXXXXX, XXX-XXX-XXX, etc.
PHONE_PATTERNS = [
    r'(?:\+48[\s\-]?)?\d{2,3}[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}',
    r'(?:\+48[\s\-]?)?\d{3}[\s\-]?\d{3}[\s\-]?\d{3}',
    r'(?:\+48[\s\-]?)?\(\d{2,3}\)[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}',
    r'tel[:\.\s]+[\+\d][\d\s\-\(\)]{8,15}',
]

EMAIL_PATTERN = r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}'

# Words that signal a contact page
CONTACT_PAGE_WORDS = [
    "kontakt", "contact", "dane-kontaktowe", "o-nas", "about",
    "skontaktuj", "napisz-do-nas", "telefon",
]

# Words that signal a product/service page
PRODUCT_PAGE_WORDS = [
    "oferta", "produkty", "uslugi", "services", "products",
    "asortyment", "katalog", "co-robimy", "nasza-oferta",
]


def clean_phone(phone):
    """Normalize a phone number."""
    digits = re.sub(r'[^\d+]', '', phone)
    # Remove "tel" prefix
    digits = re.sub(r'^tel', '', digits, flags=re.IGNORECASE)
    if len(digits) < 7:
        return None
    # Add +48 if missing and looks Polish
    if len(digits) == 9 and digits[0] in "2345678":
        digits = "+48" + digits
    elif digits.startswith("48") and len(digits) == 11:
        digits = "+" + digits
    return digits


def extract_phones(html):
    """Extract phone numbers from HTML text."""
    phones = set()
    for pattern in PHONE_PATTERNS:
        for match in re.finditer(pattern, html, re.IGNORECASE):
            cleaned = clean_phone(match.group())
            if cleaned and len(cleaned) >= 10:
                phones.add(cleaned)
    return list(phones)[:5]  # max 5 numbers


def extract_emails(html):
    """Extract email addresses from HTML text."""
    emails = set()
    # Also check for obfuscated emails: name [at] domain [dot] com
    text = html.replace("[at]", "@").replace("[dot]", ".").replace(" at ", "@").replace(" dot ", ".")
    for match in re.finditer(EMAIL_PATTERN, text):
        email = match.group().lower()
        # Skip image/css/js files
        if not any(email.endswith(ext) for ext in ['.png', '.jpg', '.gif', '.css', '.js', '.svg']):
            emails.add(email)
    return list(emails)[:5]


def extract_social_media(html):
    """Extract social media links from HTML."""
    social = {}
    patterns = {
        "facebook": r'(?:https?://)?(?:www\.)?facebook\.com/[^\s"\'<>]+',
        "instagram": r'(?:https?://)?(?:www\.)?instagram\.com/[^\s"\'<>]+',
        "linkedin": r'(?:https?://)?(?:www\.)?linkedin\.com/(?:company|in)/[^\s"\'<>]+',
        "twitter": r'(?:https?://)?(?:www\.)?(?:twitter|x)\.com/[^\s"\'<>]+',
        "youtube": r'(?:https?://)?(?:www\.)?youtube\.com/[^\s"\'<>]+',
    }
    for platform, pattern in patterns.items():
        match = re.search(pattern, html, re.IGNORECASE)
        if match:
            social[platform] = match.group()
    return social


def extract_products_from_html(html):
    """Try to extract product/service descriptions from HTML.

    Looks for meta description, h1/h2 tags, and structured data.
    """
    products = []

    # Meta description often has a business summary
    meta_match = re.search(
        r'<meta\s+(?:name|property)=["\'](?:description|og:description)["\']\s+content=["\']([^"\']{10,300})',
        html, re.IGNORECASE
    )
    if meta_match:
        products.append(("meta_description", meta_match.group(1).strip()))

    # Title tag
    title_match = re.search(r'<title>([^<]{5,200})</title>', html, re.IGNORECASE)
    if title_match:
        products.append(("title", title_match.group(1).strip()))

    # H1/H2 headings (often describe what the company does)
    for tag in ["h1", "h2"]:
        for match in re.finditer(rf'<{tag}[^>]*>([^<]{{5,150}})</{tag}>', html, re.IGNORECASE):
            text = re.sub(r'\s+', ' ', match.group(1)).strip()
            if text and len(text) > 5:
                products.append((tag, text))

    return products[:10]


def fetch_page(url, timeout=15):
    """Fetch a web page, return HTML text or None."""
    if not url:
        return None
    # Ensure URL has scheme
    if not url.startswith("http"):
        url = "https://" + url
    try:
        resp = requests.get(url, headers=HEADERS, timeout=timeout, verify=False,
                            allow_redirects=True)
        if resp.status_code == 200:
            return resp.text
    except Exception as e:
        logger.debug(f"  Fetch failed: {url} — {e}")
    # Try http if https failed
    if url.startswith("https://"):
        try:
            http_url = url.replace("https://", "http://")
            resp = requests.get(http_url, headers=HEADERS, timeout=timeout, verify=False,
                                allow_redirects=True)
            if resp.status_code == 200:
                return resp.text
        except Exception:
            pass
    return None


def find_contact_page(html, base_url):
    """Find a link to the contact page within the HTML."""
    if not base_url.startswith("http"):
        base_url = "https://" + base_url
    base_url = base_url.rstrip("/")

    for word in CONTACT_PAGE_WORDS:
        pattern = rf'href=["\']([^"\']*{word}[^"\']*)["\']'
        match = re.search(pattern, html, re.IGNORECASE)
        if match:
            href = match.group(1)
            if href.startswith("http"):
                return href
            elif href.startswith("/"):
                return base_url + href
            else:
                return base_url + "/" + href
    return None


def find_product_page(html, base_url):
    """Find a link to the products/services page within the HTML."""
    if not base_url.startswith("http"):
        base_url = "https://" + base_url
    base_url = base_url.rstrip("/")

    for word in PRODUCT_PAGE_WORDS:
        pattern = rf'href=["\']([^"\']*{word}[^"\']*)["\']'
        match = re.search(pattern, html, re.IGNORECASE)
        if match:
            href = match.group(1)
            if href.startswith("http"):
                return href
            elif href.startswith("/"):
                return base_url + href
            else:
                return base_url + "/" + href
    return None


# ─── Business directory searches ─────────────────────────────────────────────

def search_panorama_firm(company_name, nip=""):
    """Search Panorama Firm (panoramafirm.pl) for contact info."""
    try:
        url = "https://panoramafirm.pl/szukaj"
        params = {"k": company_name}
        resp = requests.get(url, params=params, headers=HEADERS, timeout=15, verify=False)
        if resp.status_code != 200:
            return {}

        html = resp.text
        phones = extract_phones(html)
        emails = extract_emails(html)

        # Try to find the company's dedicated page
        name_lower = company_name.lower()
        # Look for links containing the company name
        link_match = re.search(
            rf'href=["\']([^"\']*panoramafirm\.pl[^"\']+)["\'][^>]*>[^<]*{re.escape(name_lower[:10])}',
            html, re.IGNORECASE
        )
        if link_match:
            detail_html = fetch_page(link_match.group(1))
            if detail_html:
                phones = extract_phones(detail_html) or phones
                emails = extract_emails(detail_html) or emails

        return {"phones": phones, "emails": emails, "source": "panoramafirm.pl"}
    except Exception as e:
        logger.debug(f"  Panorama Firm search failed: {e}")
        return {}


def search_pkt(company_name):
    """Search PKT.pl (Polish Yellow Pages) for contact info."""
    try:
        url = "https://www.pkt.pl/szukaj"
        params = {"q": company_name}
        resp = requests.get(url, params=params, headers=HEADERS, timeout=15, verify=False)
        if resp.status_code != 200:
            return {}

        html = resp.text
        phones = extract_phones(html)
        emails = extract_emails(html)
        return {"phones": phones, "emails": emails, "source": "pkt.pl"}
    except Exception as e:
        logger.debug(f"  PKT search failed: {e}")
        return {}


def search_duckduckgo_contact(company_name, city=""):
    """Search DuckDuckGo for company phone number."""
    query = f'"{company_name}" telefon kontakt'
    if city:
        query += f" {city}"
    try:
        resp = requests.get(
            "https://html.duckduckgo.com/html/",
            params={"q": query},
            headers={"User-Agent": "Mozilla/5.0"},
            timeout=15, verify=False,
        )
        if resp.status_code == 200:
            phones = extract_phones(resp.text)
            emails = extract_emails(resp.text)
            return {"phones": phones, "emails": emails, "source": "duckduckgo"}
    except Exception:
        pass
    return {}


# ─── Main enrichment logic ──────────────────────────────────────────────────

def enrich_company(name, website="", nip="", existing_phone="", existing_email="", city=""):
    """
    Try to find phone, email, products for a company.

    Returns dict with: phones, emails, social_media, products, website_status, sources
    """
    result = {
        "phones_found": [],
        "emails_found": [],
        "social_media": {},
        "products_summary": "",
        "website_title": "",
        "website_status": "",
        "sources_used": [],
    }

    all_phones = set()
    all_emails = set()
    if existing_phone:
        all_phones.add(existing_phone)
    if existing_email:
        all_emails.add(existing_email)

    # Strategy 1: Scrape company website
    if website:
        logger.info(f"  Scraping website: {website}")
        homepage = fetch_page(website)
        if homepage:
            result["website_status"] = "LIVE"
            phones = extract_phones(homepage)
            emails = extract_emails(homepage)
            social = extract_social_media(homepage)
            products = extract_products_from_html(homepage)

            all_phones.update(phones)
            all_emails.update(emails)
            result["social_media"] = social
            if products:
                result["products_summary"] = " | ".join(
                    f"{src}: {txt}" for src, txt in products[:5]
                )
                result["website_title"] = products[0][1] if products else ""

            # Also check contact page
            contact_url = find_contact_page(homepage, website)
            if contact_url:
                logger.info(f"  Found contact page: {contact_url}")
                contact_html = fetch_page(contact_url)
                if contact_html:
                    all_phones.update(extract_phones(contact_html))
                    all_emails.update(extract_emails(contact_html))
                    result["sources_used"].append("website_contact")

            # Check product page
            product_url = find_product_page(homepage, website)
            if product_url:
                logger.info(f"  Found product page: {product_url}")
                product_html = fetch_page(product_url)
                if product_html:
                    more_products = extract_products_from_html(product_html)
                    if more_products:
                        result["products_summary"] += " | " + " | ".join(
                            f"{src}: {txt}" for src, txt in more_products[:3]
                        )

            result["sources_used"].append("website")
        else:
            result["website_status"] = "UNREACHABLE"

    # Strategy 2: Panorama Firm directory
    if len(all_phones) == 0 or len(all_emails) == 0:
        logger.info(f"  Searching Panorama Firm...")
        pf = search_panorama_firm(name, nip)
        if pf.get("phones"):
            all_phones.update(pf["phones"])
            result["sources_used"].append("panoramafirm.pl")
        if pf.get("emails"):
            all_emails.update(pf["emails"])
        time.sleep(1)

    # Strategy 3: PKT.pl
    if len(all_phones) == 0:
        logger.info(f"  Searching PKT.pl...")
        pkt = search_pkt(name)
        if pkt.get("phones"):
            all_phones.update(pkt["phones"])
            result["sources_used"].append("pkt.pl")
        if pkt.get("emails"):
            all_emails.update(pkt["emails"])
        time.sleep(1)

    # Strategy 4: DuckDuckGo
    if len(all_phones) == 0:
        logger.info(f"  Searching DuckDuckGo...")
        ddg = search_duckduckgo_contact(name, city)
        if ddg.get("phones"):
            all_phones.update(ddg["phones"])
            result["sources_used"].append("duckduckgo")
        if ddg.get("emails"):
            all_emails.update(ddg["emails"])
        time.sleep(2)

    result["phones_found"] = sorted(all_phones)
    result["emails_found"] = sorted(all_emails)

    return result


def main():
    parser = argparse.ArgumentParser(
        description="Enrich company data with phone numbers, emails, and products from websites."
    )
    parser.add_argument("input_file", help="Excel file with company data")
    parser.add_argument("--output", "-o", default="enriched_contacts.xlsx",
                        help="Output Excel file")
    parser.add_argument("--limit", "-n", type=int, default=0,
                        help="Only process first N rows (0=all)")
    parser.add_argument("--skip-directories", action="store_true",
                        help="Skip Panorama Firm / PKT.pl lookups (faster)")
    parser.add_argument("--skip-web-search", action="store_true",
                        help="Skip DuckDuckGo fallback (faster)")
    args = parser.parse_args()

    logger.info(f"Reading {args.input_file}")
    try:
        df = pd.read_excel(args.input_file, header=0)
    except FileNotFoundError:
        logger.error(f"File not found: {args.input_file}")
        sys.exit(1)

    logger.info(f"Loaded {len(df)} rows")

    # Find columns
    def find_col(candidates):
        for c in candidates:
            for col in df.columns:
                if col.lower().strip() == c.lower():
                    return col
        return None

    name_col = find_col(["api_name", "name", "company_name", "chain_name", "current_name", "input_name"])
    website_col = find_col(["website", "www", "url", "web"])
    phone_col = find_col(["phone", "telefon", "tel"])
    email_col = find_col(["email", "e-mail", "mail"])
    nip_col = find_col(["nip", "NIP"])
    city_col = find_col(["api_city", "city", "miasto"])

    if not name_col:
        logger.error(f"Can't find name column. Available: {list(df.columns)}")
        sys.exit(1)

    logger.info(f"Columns: name={name_col}, website={website_col}, phone={phone_col}, email={email_col}")

    # Process
    rows_to_process = df if args.limit == 0 else df.head(args.limit)
    enriched_data = []
    stats = {"total": 0, "phones_found": 0, "emails_found": 0, "websites_live": 0, "products_found": 0}

    for idx, row in rows_to_process.iterrows():
        name = str(row.get(name_col, "")).strip()
        if name.lower() in ("nan", "none", ""):
            continue

        website = str(row.get(website_col, "")).strip() if website_col else ""
        phone = str(row.get(phone_col, "")).strip() if phone_col else ""
        email = str(row.get(email_col, "")).strip() if email_col else ""
        nip = str(row.get(nip_col, "")).strip() if nip_col else ""
        city = str(row.get(city_col, "")).strip() if city_col else ""

        for val in [website, phone, email, nip, city]:
            if val.lower() in ("nan", "none"):
                val = ""

        stats["total"] += 1
        logger.info(f"[{stats['total']}] {name} | website: {website} | phone: {phone}")

        result = enrich_company(
            name, website, nip,
            existing_phone=phone if phone.lower() not in ("nan", "none", "") else "",
            existing_email=email if email.lower() not in ("nan", "none", "") else "",
            city=city if city.lower() not in ("nan", "none", "") else "",
        )

        if result["phones_found"]:
            stats["phones_found"] += 1
        if result["emails_found"]:
            stats["emails_found"] += 1
        if result["website_status"] == "LIVE":
            stats["websites_live"] += 1
        if result["products_summary"]:
            stats["products_found"] += 1

        # Build enriched row — keep all original columns + add new ones
        enriched_row = row.to_dict()
        enriched_row["enriched_phone"] = "; ".join(result["phones_found"]) if result["phones_found"] else ""
        enriched_row["enriched_email"] = "; ".join(result["emails_found"]) if result["emails_found"] else ""
        enriched_row["website_status"] = result["website_status"]
        enriched_row["website_title"] = result.get("website_title", "")
        enriched_row["products_from_website"] = result.get("products_summary", "")
        enriched_row["facebook"] = result["social_media"].get("facebook", "")
        enriched_row["instagram"] = result["social_media"].get("instagram", "")
        enriched_row["linkedin"] = result["social_media"].get("linkedin", "")
        enriched_row["contact_sources"] = "; ".join(result["sources_used"])
        enriched_data.append(enriched_row)

        logger.info(f"  -> Phones: {result['phones_found'][:3]} | "
                     f"Emails: {result['emails_found'][:2]} | "
                     f"Website: {result['website_status']}")

    # Save
    df_out = pd.DataFrame(enriched_data).fillna("")

    logger.info(f"Saving to {args.output}")
    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False, sheet_name="Enriched Data")

        # Sheet 2: Missing contacts (still no phone after enrichment)
        missing = df_out[df_out["enriched_phone"] == ""]
        if not missing.empty:
            missing.to_excel(writer, index=False, sheet_name="Still Missing Phone")

    logger.info("=" * 55)
    logger.info(f"DONE  |  Total processed: {stats['total']}")
    logger.info(f"  Phones found:    {stats['phones_found']}/{stats['total']}")
    logger.info(f"  Emails found:    {stats['emails_found']}/{stats['total']}")
    logger.info(f"  Websites live:   {stats['websites_live']}/{stats['total']}")
    logger.info(f"  Products found:  {stats['products_found']}/{stats['total']}")
    logger.info(f"Output: {args.output}")


if __name__ == "__main__":
    main()
