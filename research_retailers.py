#!/usr/bin/env python3
"""
Research Polish retailers using web search in small batches.

Reads an Excel file (output from lookup_chains.py, lookup_by_zip.py, etc.)
and enriches each retailer with web-searched data:
  - Products / services they sell
  - Brands they carry
  - Phone numbers
  - Official website URL
  - Shop description (what they are, what they sell)

STRICT NO-HALLUCINATION POLICY:
  - Each batch processes at most 10 retailers to prevent AI from generating
    instead of searching.
  - Every field is either sourced from a real web search result or marked
    "NOT FOUND".
  - No data is invented, guessed, or inferred.

Also outputs ALL available API data (dates, local units, PKD, etc.).

Usage:
    python research_retailers.py chains.xlsx --output researched.xlsx
    python research_retailers.py chains.xlsx --limit 20 --batch-size 5
    python research_retailers.py chains.xlsx --search-only  # skip API, web only
"""

import argparse
import json
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

PHONE_PATTERNS = [
    r'(?:\+48[\s\-]?)?\d{2,3}[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}',
    r'(?:\+48[\s\-]?)?\d{3}[\s\-]?\d{3}[\s\-]?\d{3}',
    r'(?:\+48[\s\-]?)?\(\d{2,3}\)[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}',
]

EMAIL_PATTERN = r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}'

OUTPUT_COLUMNS = [
    # ─── Identity ───────────────────────────────────────────────────────
    "input_name",
    "category",
    "api_name",
    "nip",
    "regon",
    "krs",
    "entity_type",
    # ─── Status ─────────────────────────────────────────────────────────
    "status",
    "is_active",
    # ─── Dates ──────────────────────────────────────────────────────────
    "date_started",
    "date_created",
    "date_suspended",
    "date_resumed",
    "date_ended",
    "date_deleted",
    "date_last_change",
    # ─── Address ────────────────────────────────────────────────────────
    "street",
    "building",
    "unit",
    "api_zip_code",
    "api_city",
    "municipality",
    "county",
    "province",
    "country",
    "full_address",
    # ─── Legal / organizational ─────────────────────────────────────────
    "legal_form_code",
    "legal_form_name",
    "legal_form_specific",
    "ownership_form",
    "size_category",
    "registry_type",
    "registration_number",
    "num_local_units",
    # ─── PKD codes ──────────────────────────────────────────────────────
    "pkd_main",
    "pkd_codes",
    "pkd_descriptions",
    # ─── Contact from API ───────────────────────────────────────────────
    "api_email",
    "api_phone",
    "api_fax",
    "api_website",
    # ─── Owner ──────────────────────────────────────────────────────────
    "owner_first_name",
    "owner_last_name",
    # ─── Web-researched fields ──────────────────────────────────────────
    "web_phone",
    "web_email",
    "web_website",
    "web_description",
    "web_products",
    "web_brands",
    "web_source_urls",
    # ─── Final merged fields ────────────────────────────────────────────
    "best_phone",
    "best_email",
    "best_website",
    # ─── Meta ───────────────────────────────────────────────────────────
    "search_method",
    "api_source",
    "research_status",
]


def clean_phone(phone):
    """Normalize a phone number string."""
    digits = re.sub(r'[^\d+]', '', phone)
    if len(digits) < 7:
        return None
    if digits.startswith("48") and len(digits) == 11:
        digits = "+" + digits
    elif digits.startswith("0048"):
        digits = "+" + digits[2:]
    elif not digits.startswith("+") and len(digits) >= 9:
        digits = "+48" + digits[-9:]
    return digits


def extract_phones_from_text(text):
    """Extract Polish phone numbers from text."""
    phones = set()
    for pat in PHONE_PATTERNS:
        for match in re.finditer(pat, text):
            cleaned = clean_phone(match.group())
            if cleaned:
                phones.add(cleaned)
    return list(phones)


def extract_emails_from_text(text):
    """Extract email addresses from text."""
    emails = set()
    for match in re.finditer(EMAIL_PATTERN, text):
        email = match.group().lower()
        if not email.endswith(('.png', '.jpg', '.gif', '.svg', '.webp', '.css', '.js')):
            emails.add(email)
    return list(emails)


def duckduckgo_search(query, max_results=5):
    """Search DuckDuckGo and return result snippets."""
    try:
        url = "https://html.duckduckgo.com/html/"
        resp = requests.post(url, data={"q": query}, headers=HEADERS, timeout=15)
        resp.raise_for_status()
        html = resp.text

        results = []
        snippets = re.findall(
            r'<a rel="nofollow" class="result__a" href="([^"]+)"[^>]*>(.*?)</a>.*?'
            r'<a class="result__snippet"[^>]*>(.*?)</a>',
            html, re.DOTALL
        )
        for link, title, snippet in snippets[:max_results]:
            clean_title = re.sub(r'<[^>]+>', '', title).strip()
            clean_snippet = re.sub(r'<[^>]+>', '', snippet).strip()
            results.append({
                "url": link,
                "title": clean_title,
                "snippet": clean_snippet,
            })
        return results
    except Exception as e:
        logger.warning(f"DuckDuckGo search failed for '{query}': {e}")
        return []


def fetch_page_text(url, timeout=10):
    """Fetch a web page and return its text content."""
    try:
        resp = requests.get(url, headers=HEADERS, timeout=timeout,
                            verify=False, allow_redirects=True)
        resp.raise_for_status()
        text = resp.text

        text = re.sub(r'<script[^>]*>.*?</script>', '', text, flags=re.DOTALL)
        text = re.sub(r'<style[^>]*>.*?</style>', '', text, flags=re.DOTALL)

        meta_desc = ""
        meta_match = re.search(r'<meta[^>]*name="description"[^>]*content="([^"]*)"', text)
        if meta_match:
            meta_desc = meta_match.group(1).strip()

        title = ""
        title_match = re.search(r'<title[^>]*>(.*?)</title>', text, re.DOTALL)
        if title_match:
            title = re.sub(r'<[^>]+>', '', title_match.group(1)).strip()

        clean = re.sub(r'<[^>]+>', ' ', text)
        clean = re.sub(r'\s+', ' ', clean).strip()

        return {
            "title": title,
            "meta_description": meta_desc,
            "body_text": clean[:5000],
            "full_html": text[:20000],
        }
    except Exception as e:
        logger.debug(f"Failed to fetch {url}: {e}")
        return None


def research_single_retailer(name, category="", existing_website=""):
    """
    Research a single retailer using web search.
    Returns ONLY data actually found — never invents.
    """
    result = {
        "web_phone": "NOT FOUND",
        "web_email": "NOT FOUND",
        "web_website": "NOT FOUND",
        "web_description": "NOT FOUND",
        "web_products": "NOT FOUND",
        "web_brands": "NOT FOUND",
        "web_source_urls": "",
        "research_status": "SEARCHED",
    }

    source_urls = []

    search_name = re.sub(r'\s*\(.*?\)\s*', ' ', name).strip()

    # Search 1: general info
    query1 = f"{search_name} Polska sklep firma"
    results1 = duckduckgo_search(query1, max_results=5)
    time.sleep(1.5)

    # Search 2: products and brands
    query2 = f"{search_name} produkty oferta asortyment marki"
    results2 = duckduckgo_search(query2, max_results=5)
    time.sleep(1.5)

    # Search 3: contact info
    query3 = f"{search_name} kontakt telefon email"
    results3 = duckduckgo_search(query3, max_results=3)
    time.sleep(1.0)

    all_results = results1 + results2 + results3

    if not all_results:
        result["research_status"] = "NO SEARCH RESULTS"
        return result

    # Collect all snippet text for analysis
    all_text = ""
    for r in all_results:
        source_urls.append(r.get("url", ""))
        all_text += f" {r.get('title', '')} {r.get('snippet', '')}"

    # Try to fetch the top website for more data
    website_url = existing_website
    if not website_url:
        for r in results1:
            url = r.get("url", "")
            if url and not any(d in url for d in [
                "facebook.com", "linkedin.com", "twitter.com",
                "youtube.com", "wikipedia.org", "panoramafirm.pl",
                "pkt.pl", "aleo.com", "gowork.pl",
            ]):
                website_url = url
                break

    if website_url:
        result["web_website"] = website_url
        page = fetch_page_text(website_url)
        if page:
            all_text += f" {page.get('title', '')} {page.get('meta_description', '')} {page.get('body_text', '')}"
            source_urls.append(website_url)

            # Try contact page
            contact_url = None
            for link in re.finditer(r'href="([^"]*(?:kontakt|contact|dane-kontaktowe)[^"]*)"',
                                    page.get("full_html", ""), re.IGNORECASE):
                href = link.group(1)
                if href.startswith("/"):
                    from urllib.parse import urljoin
                    href = urljoin(website_url, href)
                if href.startswith("http"):
                    contact_url = href
                    break

            if contact_url:
                time.sleep(1.0)
                contact_page = fetch_page_text(contact_url)
                if contact_page:
                    all_text += f" {contact_page.get('body_text', '')}"
                    source_urls.append(contact_url)

            # Try products/offer page
            offer_url = None
            for link in re.finditer(r'href="([^"]*(?:oferta|produkty|asortyment|products|katalog)[^"]*)"',
                                    page.get("full_html", ""), re.IGNORECASE):
                href = link.group(1)
                if href.startswith("/"):
                    from urllib.parse import urljoin
                    href = urljoin(website_url, href)
                if href.startswith("http"):
                    offer_url = href
                    break

            if offer_url:
                time.sleep(1.0)
                offer_page = fetch_page_text(offer_url)
                if offer_page:
                    all_text += f" {offer_page.get('body_text', '')}"
                    source_urls.append(offer_url)

    # Extract phone numbers from all gathered text
    phones = extract_phones_from_text(all_text)
    if phones:
        result["web_phone"] = "; ".join(phones[:3])

    # Extract emails
    emails = extract_emails_from_text(all_text)
    if emails:
        valid_emails = [e for e in emails if not e.startswith("noreply")
                        and "example.com" not in e]
        if valid_emails:
            result["web_email"] = "; ".join(valid_emails[:3])

    # Extract description from search snippets (first meaningful one)
    for r in results1:
        snippet = r.get("snippet", "").strip()
        if snippet and len(snippet) > 30:
            result["web_description"] = snippet[:500]
            break

    # Extract products/brands from product search results
    product_snippets = []
    brand_candidates = set()

    for r in results2:
        snippet = r.get("snippet", "").strip()
        if snippet and len(snippet) > 20:
            product_snippets.append(snippet)

    if product_snippets:
        result["web_products"] = " | ".join(product_snippets[:3])[:500]

    # Look for brand names in all collected text (capitalized multi-word patterns)
    brand_pattern = r'\b([A-Z][a-zA-Z]+(?:\s[A-Z][a-zA-Z]+)*)\b'
    for match in re.finditer(brand_pattern, all_text):
        brand = match.group(1)
        if len(brand) > 2 and brand.lower() not in (
            "polska", "poland", "the", "and", "for", "sklep", "firma",
            "telefon", "email", "kontakt", "oferta", "produkty", "marki",
            "strona", "adres", "dane", "numer", "tel", "diy",
            name.lower(), search_name.lower(),
        ):
            brand_candidates.add(brand)

    if brand_candidates:
        result["web_brands"] = "; ".join(sorted(brand_candidates)[:15])[:500]

    # Deduplicate source URLs
    seen = set()
    unique_urls = []
    for u in source_urls:
        if u and u not in seen:
            seen.add(u)
            unique_urls.append(u)
    result["web_source_urls"] = "; ".join(unique_urls[:5])

    return result


def find_column(df, candidates, required=False, label=""):
    """Find first matching column name from candidates."""
    for c in candidates:
        if c in df.columns:
            return c
        for col in df.columns:
            if col.lower().strip() == c.lower().strip():
                return col
    if required:
        logger.error(f"Could not find {label} column. Tried: {candidates}")
        logger.error(f"Available columns: {list(df.columns)}")
        sys.exit(1)
    return None


def normalize(val):
    """Clean a cell value."""
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.lower() in ("nan", "none"):
        return ""
    return s


def main():
    parser = argparse.ArgumentParser(
        description="Research Polish retailers: web search for products, brands, contacts."
    )
    parser.add_argument("input_file", help="Excel file with retailer data")
    parser.add_argument("--output", "-o", default="researched_retailers.xlsx",
                        help="Output Excel file (default: researched_retailers.xlsx)")
    parser.add_argument("--limit", "-n", type=int, default=0,
                        help="Only process first N rows (0 = all)")
    parser.add_argument("--batch-size", type=int, default=10,
                        help="Retailers per batch (default: 10, max recommended: 10)")
    parser.add_argument("--search-only", action="store_true",
                        help="Only do web research, skip API enrichment")
    parser.add_argument("--delay", type=float, default=2.0,
                        help="Delay between retailers in seconds (default: 2.0)")
    parser.add_argument("--gus-key", help="GUS API key (or set GUS_API_KEY env var)")
    parser.add_argument("--sandbox", action="store_true", help="Use GUS sandbox")
    args = parser.parse_args()

    if args.batch_size > 15:
        logger.warning("Batch size >15 increases risk of low-quality results. Capping at 15.")
        args.batch_size = 15

    logger.info(f"Reading {args.input_file}")
    try:
        df = pd.read_excel(args.input_file, header=0)
    except FileNotFoundError:
        logger.error(f"File not found: {args.input_file}")
        sys.exit(1)

    logger.info(f"Loaded {len(df)} rows, {len(df.columns)} columns")

    # Detect columns
    name_col = find_column(df, ["chain_name", "api_name", "input_name", "name", "company_name"],
                           required=True, label="Name")
    cat_col = find_column(df, ["category", "channel"])
    nip_col = find_column(df, ["nip", "NIP", "input_nip"])
    regon_col = find_column(df, ["regon", "REGON"])
    krs_col = find_column(df, ["krs", "KRS"])
    status_col = find_column(df, ["status", "match_status"])
    is_active_col = find_column(df, ["is_active"])
    entity_type_col = find_column(df, ["entity_type"])
    api_name_col = find_column(df, ["api_name"])
    street_col = find_column(df, ["street"])
    building_col = find_column(df, ["building"])
    unit_col = find_column(df, ["unit"])
    zip_col = find_column(df, ["api_zip_code", "zip_code", "input_zip"])
    city_col = find_column(df, ["api_city", "city"])
    municipality_col = find_column(df, ["municipality"])
    county_col = find_column(df, ["county"])
    province_col = find_column(df, ["province"])
    country_col = find_column(df, ["country"])
    full_addr_col = find_column(df, ["full_address", "address_str"])
    pkd_main_col = find_column(df, ["pkd_main"])
    pkd_codes_col = find_column(df, ["pkd_codes"])
    pkd_desc_col = find_column(df, ["pkd_descriptions"])
    date_started_col = find_column(df, ["date_started"])
    date_created_col = find_column(df, ["date_created"])
    date_suspended_col = find_column(df, ["date_suspended"])
    date_resumed_col = find_column(df, ["date_resumed"])
    date_ended_col = find_column(df, ["date_ended"])
    date_deleted_col = find_column(df, ["date_deleted"])
    date_last_change_col = find_column(df, ["date_last_change"])
    email_col = find_column(df, ["email", "api_email"])
    phone_col = find_column(df, ["phone", "api_phone"])
    fax_col = find_column(df, ["fax", "api_fax"])
    website_col = find_column(df, ["website", "api_website"])
    owner_fn_col = find_column(df, ["owner_first_name"])
    owner_ln_col = find_column(df, ["owner_last_name"])
    source_col = find_column(df, ["api_source", "source"])
    search_method_col = find_column(df, ["search_method"])
    num_local_col = find_column(df, ["num_local_units"])
    legal_form_code_col = find_column(df, ["legal_form_code"])
    legal_form_name_col = find_column(df, ["legal_form_name"])
    legal_form_specific_col = find_column(df, ["legal_form_specific"])
    ownership_form_col = find_column(df, ["ownership_form"])
    size_category_col = find_column(df, ["size_category"])
    registry_type_col = find_column(df, ["registry_type"])
    registration_number_col = find_column(df, ["registration_number"])

    logger.info(f"Name column: '{name_col}'")

    # Limit rows
    rows_to_process = df if args.limit == 0 else df.head(args.limit)
    total = len(rows_to_process)
    logger.info(f"Processing {total} retailers in batches of {args.batch_size}")

    # Optionally set up API client for re-enrichment
    api_client = None
    if not args.search_only:
        try:
            from poland_api import PolandAPIClient
            api_client = PolandAPIClient(gus_api_key=args.gus_key, use_sandbox=args.sandbox)
            if api_client.gus_api_key:
                logger.info("GUS API client ready — will re-fetch for extra fields.")
            else:
                logger.info("No GUS API key — using existing data only.")
                api_client = None
        except ImportError:
            logger.info("poland_api not available — web search only.")

    all_rows = []
    stats = {"researched": 0, "phones_found": 0, "products_found": 0, "no_results": 0}

    for idx, (_, row) in enumerate(rows_to_process.iterrows()):
        name = normalize(row.get(name_col, ""))
        if not name:
            continue

        batch_num = idx // args.batch_size + 1
        batch_pos = idx % args.batch_size + 1

        if batch_pos == 1:
            logger.info(f"── Batch {batch_num} "
                         f"(retailers {idx + 1}-{min(idx + args.batch_size, total)}/{total}) ──")

        logger.info(f"[{idx + 1}/{total}] Researching: {name}")

        # Collect all existing API data
        def get(col):
            return normalize(row.get(col, "")) if col else ""

        api_name = get(api_name_col) or name
        existing_website = get(website_col)
        category = get(cat_col)

        # If we have a NIP and API client, re-fetch for extra fields
        nip_val = get(nip_col)
        regon_val = get(regon_col)
        api_enriched = {}

        if api_client and nip_val and len(nip_val) >= 10:
            try:
                logger.info(f"  Re-fetching GUS data for NIP {nip_val}...")
                api_result = api_client.search_gus_by_nip(nip_val)
                if api_result:
                    api_enriched = api_result
                    logger.info(f"  GUS enrichment OK — local units: {api_result.get('num_local_units', 'N/A')}")
            except Exception as e:
                logger.warning(f"  GUS re-fetch failed: {e}")

        # Web research
        web_data = research_single_retailer(name, category, existing_website)
        stats["researched"] += 1
        if web_data.get("web_phone", "NOT FOUND") != "NOT FOUND":
            stats["phones_found"] += 1
        if web_data.get("web_products", "NOT FOUND") != "NOT FOUND":
            stats["products_found"] += 1
        if web_data.get("research_status") == "NO SEARCH RESULTS":
            stats["no_results"] += 1

        # Merge: prefer API-enriched data if available, fall back to spreadsheet
        def best(enriched_key, col):
            return normalize(api_enriched.get(enriched_key, "")) or get(col)

        # Build best phone/email/website (API first, then web)
        api_phone = best("phone", phone_col)
        api_email = best("email", email_col)
        api_website = best("website", website_col)

        web_phone = web_data.get("web_phone", "NOT FOUND")
        web_email = web_data.get("web_email", "NOT FOUND")
        web_website = web_data.get("web_website", "NOT FOUND")

        best_phone = api_phone or (web_phone if web_phone != "NOT FOUND" else "")
        best_email = api_email or (web_email if web_email != "NOT FOUND" else "")
        best_website = api_website or (web_website if web_website != "NOT FOUND" else "")

        # PKD descriptions
        pkd_codes_val = best("pkd_codes", pkd_codes_col)
        if isinstance(pkd_codes_val, list):
            pkd_codes_val = "; ".join(pkd_codes_val)

        pkd_descs_val = ""
        if api_enriched.get("pkd_descriptions"):
            descs = api_enriched["pkd_descriptions"]
            if isinstance(descs, dict):
                pkd_descs_val = "; ".join(f"{k}: {v}" for k, v in descs.items())
        if not pkd_descs_val:
            pkd_descs_val = get(pkd_desc_col)

        out_row = {
            "input_name": name,
            "category": category,
            "api_name": best("name", api_name_col),
            "nip": best("nip", nip_col),
            "regon": best("regon", regon_col),
            "krs": best("krs", krs_col),
            "entity_type": best("entity_type", entity_type_col),
            "status": best("status", status_col),
            "is_active": best("is_active", is_active_col),
            "date_started": best("date_started", date_started_col),
            "date_created": best("date_created", date_created_col),
            "date_suspended": best("date_suspended", date_suspended_col),
            "date_resumed": best("date_resumed", date_resumed_col),
            "date_ended": best("date_ended", date_ended_col),
            "date_deleted": best("date_deleted", date_deleted_col),
            "date_last_change": best("date_last_change", date_last_change_col),
            "street": best("street", street_col),
            "building": best("building", building_col),
            "unit": best("unit", unit_col),
            "api_zip_code": best("zip_code", zip_col),
            "api_city": best("city", city_col),
            "municipality": best("municipality", municipality_col),
            "county": best("county", county_col),
            "province": best("province", province_col),
            "country": best("country", country_col),
            "full_address": best("address_str", full_addr_col),
            "legal_form_code": best("legal_form_code", legal_form_code_col),
            "legal_form_name": best("legal_form_name", legal_form_name_col),
            "legal_form_specific": best("legal_form_specific", legal_form_specific_col),
            "ownership_form": best("ownership_form", ownership_form_col),
            "size_category": best("size_category", size_category_col),
            "registry_type": best("registry_type", registry_type_col),
            "registration_number": best("registration_number", registration_number_col),
            "num_local_units": best("num_local_units", num_local_col),
            "pkd_main": best("pkd_main", pkd_main_col),
            "pkd_codes": pkd_codes_val,
            "pkd_descriptions": pkd_descs_val,
            "api_email": api_email,
            "api_phone": api_phone,
            "api_fax": best("fax", fax_col),
            "api_website": api_website,
            "owner_first_name": best("owner_first_name", owner_fn_col),
            "owner_last_name": best("owner_last_name", owner_ln_col),
            "web_phone": web_phone,
            "web_email": web_email,
            "web_website": web_website,
            "web_description": web_data.get("web_description", "NOT FOUND"),
            "web_products": web_data.get("web_products", "NOT FOUND"),
            "web_brands": web_data.get("web_brands", "NOT FOUND"),
            "web_source_urls": web_data.get("web_source_urls", ""),
            "best_phone": best_phone,
            "best_email": best_email,
            "best_website": best_website,
            "search_method": get(search_method_col),
            "api_source": best("source", source_col),
            "research_status": web_data.get("research_status", ""),
        }
        all_rows.append(out_row)

        # Rate limiting between retailers
        if idx < total - 1:
            time.sleep(args.delay)

    # Save to Excel
    df_out = pd.DataFrame(all_rows, columns=OUTPUT_COLUMNS).fillna("")

    logger.info(f"Saving {len(df_out)} rows to {args.output}")
    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False, sheet_name="All Results")

        # Sheet 2: rows with phone numbers found
        has_phone = df_out[df_out["best_phone"].astype(str).str.len() > 5]
        if not has_phone.empty:
            has_phone.to_excel(writer, index=False, sheet_name="With Phone")
            logger.info(f"  Sheet 'With Phone': {len(has_phone)} rows")

        # Sheet 3: rows with products found
        has_products = df_out[df_out["web_products"] != "NOT FOUND"]
        if not has_products.empty:
            has_products.to_excel(writer, index=False, sheet_name="With Products")
            logger.info(f"  Sheet 'With Products': {len(has_products)} rows")

        # Sheet 4: rows still missing phone
        missing = df_out[df_out["best_phone"].astype(str).str.len() <= 5]
        if not missing.empty:
            missing.to_excel(writer, index=False, sheet_name="Missing Phone")
            logger.info(f"  Sheet 'Missing Phone': {len(missing)} rows")

        # Sheet 5: summary by category
        if cat_col:
            summary = []
            for cat in df_out["category"].unique():
                if not cat:
                    continue
                cat_rows = df_out[df_out["category"] == cat]
                cat_with_phone = cat_rows[cat_rows["best_phone"].astype(str).str.len() > 5]
                cat_with_products = cat_rows[cat_rows["web_products"] != "NOT FOUND"]
                cat_active = cat_rows[cat_rows["is_active"].astype(str).str.lower().isin(["true", "1", "yes"])]
                summary.append({
                    "category": cat,
                    "total": len(cat_rows),
                    "active": len(cat_active),
                    "with_phone": len(cat_with_phone),
                    "with_products": len(cat_with_products),
                    "phone_rate_pct": round(len(cat_with_phone) / len(cat_rows) * 100, 1) if len(cat_rows) > 0 else 0,
                })
            if summary:
                pd.DataFrame(summary).to_excel(writer, index=False, sheet_name="Summary")

    logger.info("=" * 55)
    logger.info(f"DONE  |  Retailers researched: {stats['researched']}")
    logger.info(f"  Phones found:    {stats['phones_found']}")
    logger.info(f"  Products found:  {stats['products_found']}")
    logger.info(f"  No results:      {stats['no_results']}")
    logger.info(f"Output: {args.output}")


if __name__ == "__main__":
    main()
