#!/usr/bin/env python3
"""
Find current names of Polish businesses that may have changed names.

Polish businesses often change names (rebrand, restructure, tax optimization).
This script takes old/known names and tries to find the current registered name.

Search strategies:
  1. GUS fuzzy name search -> get REGON/NIP -> current registered name
  2. KRS full history lookup -> name change records
  3. Google/web search -> find rebranded businesses online
  4. Keyword extraction -> search with core words from old name

Input Excel:
  Column A = Old/known business name
  Column B = City or zip (optional, helps narrow search)
  Column C = NIP (optional, if known — direct lookup)

Usage:
    python lookup_name_changes.py input.xlsx --output name_changes.xlsx
    python lookup_name_changes.py input.xlsx --limit 3
    python lookup_name_changes.py input.xlsx --web-search   # enable Google search
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

from poland_api import PolandAPIClient

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

KRS_API_BASE = "https://api-krs.ms.gov.pl/api/krs"

OUTPUT_COLUMNS = [
    "old_name",
    "input_city_zip",
    "input_nip",
    "search_method",
    "match_confidence",
    "current_name",
    "name_changed",
    "nip",
    "regon",
    "krs",
    "status",
    "is_active",
    "street",
    "building",
    "api_zip_code",
    "api_city",
    "full_address",
    "pkd_main",
    "pkd_codes",
    "pkd_descriptions",
    "date_started",
    "email",
    "phone",
    "website",
    "name_history",
    "api_source",
    "web_search_snippet",
]


# ─── Name utilities ──────────────────────────────────────────────────────────

LEGAL_SUFFIXES = [
    " sp. z o.o. sp. k.", " sp. z o.o. sp.k.",
    " sp. z o.o.", " sp.z o.o.", " spółka z o.o.", " spolka z o.o.",
    " sp. z o. o.", " s.a.", " sp.j.", " sp.k.", " sp. komandytowa",
    " spółka jawna", " spolka jawna", " spółka akcyjna", " spolka akcyjna",
    " s.c.", " sp. cywilna", " spółka cywilna", " spolka cywilna",
    " z o.o.", " z.o.o.",
]


def strip_legal_form(name):
    """Remove legal form suffixes from business name."""
    lower = name.lower().strip()
    for suffix in LEGAL_SUFFIXES:
        if lower.endswith(suffix):
            return name[:len(name) - len(suffix)].strip()
    return name.strip()


def extract_keywords(name):
    """Extract meaningful keywords from a business name (skip common Polish words)."""
    STOP_WORDS = {
        "sp", "z", "o", "oo", "sa", "sc", "sk", "sj", "spolka", "spółka",
        "przedsiębiorstwo", "przedsiebiorstwo", "firma", "zaklad", "zakład",
        "handel", "handlowe", "usługi", "uslugi", "produkcja", "bis",
        "i", "w", "do", "na", "pod", "dla", "oraz", "the", "and", "of",
        "pphu", "phu", "fhu", "phup", "fphu", "zph", "zpchr",
    }
    # Split on spaces and punctuation
    words = re.split(r'[\s\-\.\,\/\(\)]+', name)
    keywords = []
    for w in words:
        w_clean = w.strip().lower()
        if len(w_clean) >= 2 and w_clean not in STOP_WORDS:
            keywords.append(w)
    return keywords


def name_similarity(name1, name2):
    """Simple word-overlap similarity score between two names (0.0 to 1.0)."""
    if not name1 or not name2:
        return 0.0
    words1 = set(extract_keywords(name1.lower()))
    words2 = set(extract_keywords(name2.lower()))
    if not words1 or not words2:
        return 0.0
    overlap = words1 & words2
    return len(overlap) / max(len(words1), len(words2))


def names_are_same(name1, name2):
    """Check if two names are essentially the same (ignoring legal form)."""
    n1 = strip_legal_form(name1).lower().strip()
    n2 = strip_legal_form(name2).lower().strip()
    return n1 == n2


# ─── KRS full history lookup ────────────────────────────────────────────────

def krs_search_by_name(name, max_results=5):
    """Search KRS registry by company name. Returns list of basic results."""
    try:
        # KRS has a free name search endpoint
        response = requests.get(
            f"{KRS_API_BASE}/OdpisPelny",
            params={"nazwa": name, "limit": max_results},
            headers={"Accept": "application/json"},
            timeout=30,
            verify=False,
        )
        if response.status_code == 404:
            return []
        if response.status_code != 200:
            return []
        data = response.json()
        if isinstance(data, list):
            return data
        return data.get("items", data.get("odpisy", []))
    except Exception as e:
        logger.debug(f"KRS name search failed: {e}")
        return []


def krs_get_full_history(krs_number):
    """Get full KRS record including name change history."""
    if not krs_number:
        return None
    krs_number = str(krs_number).strip().zfill(10)

    try:
        # OdpisPelny has historical data; OdpisAktualny only has current
        response = requests.get(
            f"{KRS_API_BASE}/OdpisPelny/{krs_number}",
            headers={"Accept": "application/json"},
            timeout=30,
            verify=False,
        )
        if response.status_code == 404:
            # Fall back to OdpisAktualny
            response = requests.get(
                f"{KRS_API_BASE}/OdpisAktualny/{krs_number}",
                headers={"Accept": "application/json"},
                timeout=30,
                verify=False,
            )
        if response.status_code != 200:
            return None
        return response.json()
    except Exception as e:
        logger.debug(f"KRS full history error: {e}")
        return None


def extract_name_history(krs_data):
    """Extract historical names from a KRS full record."""
    if not krs_data:
        return []

    names = []
    odpis = (krs_data.get("odppisPelnyGrupa")
             or krs_data.get("odppisTresciGrupa")
             or krs_data)

    # Current name
    dzial1 = odpis.get("dzial1", {})
    dane = dzial1.get("danePodmiotu", {})
    current = dane.get("nazwa", "")
    if current:
        names.append(("CURRENT", current))

    # Historical names are in dzial6 or in name change records
    dzial6 = odpis.get("dzial6", {})

    # Check for name change entries in various possible locations
    for key in ["informacjeOZmiananieNazwy", "zmianyNazwy", "zmianaNazwy",
                "danePodmiotuHistoria", "historiaZmian"]:
        history = dzial6.get(key, odpis.get(key, []))
        if isinstance(history, list):
            for entry in history:
                old_name = (entry.get("nazwa", "")
                            or entry.get("nazwaPoprzednia", "")
                            or entry.get("wartoscPoprzednia", ""))
                if old_name and old_name != current:
                    date = entry.get("dataZmiany", entry.get("data", ""))
                    names.append((date or "HISTORICAL", old_name))

    # Also check dzial1 for previous names
    nazwy_hist = dane.get("nazwyPoprzednie", dane.get("historyczne", []))
    if isinstance(nazwy_hist, list):
        for entry in nazwy_hist:
            n = entry.get("nazwa", "") if isinstance(entry, dict) else str(entry)
            if n and n != current:
                names.append(("HISTORICAL", n))

    return names


# ─── Web search ──────────────────────────────────────────────────────────────

def web_search_business(old_name, city=""):
    """Search Google for a Polish business name to find current name.

    Uses a simple request to Google — may be rate-limited.
    Returns a snippet or None.
    """
    query = f'"{old_name}" firma Polska'
    if city:
        query += f" {city}"
    query += " nowa nazwa OR zmiana nazwy OR KRS OR NIP"

    try:
        # Use DuckDuckGo HTML (more lenient than Google)
        resp = requests.get(
            "https://html.duckduckgo.com/html/",
            params={"q": query},
            headers={"User-Agent": "Mozilla/5.0"},
            timeout=15,
            verify=False,
        )
        if resp.status_code == 200:
            # Extract snippets from results
            text = resp.text
            snippets = re.findall(r'class="result__snippet">(.*?)</a>', text, re.DOTALL)
            if snippets:
                # Clean HTML
                clean = re.sub(r'<[^>]+>', '', snippets[0]).strip()
                return clean[:300]

            # Try another pattern
            snippets = re.findall(r'class="result__snippet"[^>]*>(.*?)</td>', text, re.DOTALL)
            if snippets:
                clean = re.sub(r'<[^>]+>', '', snippets[0]).strip()
                return clean[:300]

        return None
    except Exception as e:
        logger.debug(f"Web search failed for '{old_name}': {e}")
        return None


# ─── Main search logic ──────────────────────────────────────────────────────

def search_current_name(api_client, old_name, city_zip="", nip="", do_web_search=False):
    """
    Search for the current name of a business given its old name.

    Returns list of candidate dicts with:
      current_name, match_confidence, search_method, name_changed, etc.
    """
    candidates = []

    # Strategy 0: Direct NIP lookup (most reliable)
    if nip:
        nip_clean = nip.strip().replace("-", "").replace(" ", "")
        if nip_clean.isdigit() and len(nip_clean) == 10:
            logger.info(f"  Trying direct NIP lookup: {nip_clean}")
            result = api_client.search_gus_by_nip(nip_clean)
            if result:
                current = result.get("name", "")
                changed = not names_are_same(old_name, current)
                sim = name_similarity(old_name, current)
                candidates.append({
                    "result": result,
                    "current_name": current,
                    "name_changed": "YES" if changed else "NO",
                    "search_method": "NIP_DIRECT",
                    "match_confidence": "HIGH",
                    "similarity": sim,
                    "name_history": "",
                    "web_snippet": "",
                })
                # If found by NIP, also check KRS for name history
                krs = result.get("krs", "")
                if krs:
                    krs_data = krs_get_full_history(krs)
                    history = extract_name_history(krs_data)
                    if history:
                        hist_str = " -> ".join(f"{n} ({d})" for d, n in history)
                        candidates[-1]["name_history"] = hist_str
                return candidates

    # Strategy 1: GUS fuzzy name search
    logger.info(f"  Trying GUS name search: '{old_name}'")
    name_variants = [old_name]

    # Add stripped version
    stripped = strip_legal_form(old_name)
    if stripped.lower() != old_name.lower():
        name_variants.append(stripped)

    # Add individual significant keywords
    keywords = extract_keywords(old_name)
    if len(keywords) >= 2:
        # Try first keyword (usually the distinctive brand name)
        name_variants.append(keywords[0])
        # Try first two keywords
        name_variants.append(" ".join(keywords[:2]))

    for variant in name_variants:
        results = api_client.search_gus_by_name(variant)
        if results:
            for r in results:
                current = r.get("name", "")
                sim = name_similarity(old_name, current)
                changed = not names_are_same(old_name, current)

                # Determine confidence
                if sim >= 0.7:
                    confidence = "HIGH"
                elif sim >= 0.4:
                    confidence = "MEDIUM"
                elif sim >= 0.2:
                    confidence = "LOW"
                else:
                    confidence = "VERY LOW"

                # Boost confidence if city/zip matches
                if city_zip:
                    city_zip_lower = city_zip.lower().strip()
                    r_city = r.get("city", "").lower()
                    r_zip = r.get("zip_code", "").replace("-", "")
                    input_zip = city_zip_lower.replace("-", "")
                    if city_zip_lower in r_city or r_city in city_zip_lower:
                        if confidence == "LOW":
                            confidence = "MEDIUM"
                        elif confidence == "MEDIUM":
                            confidence = "HIGH"
                    elif input_zip == r_zip:
                        if confidence == "LOW":
                            confidence = "MEDIUM"

                candidates.append({
                    "result": r,
                    "current_name": current,
                    "name_changed": "YES" if changed else "NO",
                    "search_method": f"GUS_NAME ({variant})",
                    "match_confidence": confidence,
                    "similarity": sim,
                    "name_history": "",
                    "web_snippet": "",
                })

            # If we found good matches, don't try looser variants
            good = [c for c in candidates if c["match_confidence"] in ("HIGH", "MEDIUM")]
            if good:
                break

    # Strategy 2: Check KRS for name history on best candidates
    for cand in candidates[:3]:
        krs = cand["result"].get("krs", "")
        if krs:
            logger.info(f"  Checking KRS history for {krs}")
            krs_data = krs_get_full_history(krs)
            history = extract_name_history(krs_data)
            if history:
                hist_str = " -> ".join(f"{n} ({d})" for d, n in history)
                cand["name_history"] = hist_str
                # Check if old name appears in history
                for _, hist_name in history:
                    if name_similarity(old_name, hist_name) > 0.5:
                        cand["match_confidence"] = "HIGH"
                        cand["name_changed"] = "YES"
                        cand["search_method"] += " + KRS_HISTORY"
                        break

    # Strategy 3: Web search (optional)
    if do_web_search and not any(c["match_confidence"] == "HIGH" for c in candidates):
        logger.info(f"  Trying web search for '{old_name}'")
        snippet = web_search_business(old_name, city_zip)
        if snippet:
            if candidates:
                candidates[0]["web_snippet"] = snippet
            else:
                candidates.append({
                    "result": {},
                    "current_name": "",
                    "name_changed": "UNKNOWN",
                    "search_method": "WEB_SEARCH",
                    "match_confidence": "WEB ONLY",
                    "similarity": 0,
                    "name_history": "",
                    "web_snippet": snippet,
                })
        time.sleep(2)  # be polite with web searches

    # Sort by confidence then similarity
    conf_order = {"HIGH": 0, "MEDIUM": 1, "LOW": 2, "VERY LOW": 3, "WEB ONLY": 4}
    candidates.sort(key=lambda c: (conf_order.get(c["match_confidence"], 5), -c["similarity"]))

    return candidates


def get_pkd_description(code):
    try:
        from pkd_codes import get_pkd_description as _get_desc
        return _get_desc(code)
    except ImportError:
        return ""


def build_row(old_name, city_zip, input_nip, candidate):
    r = candidate.get("result", {})
    pkd_codes = r.get("pkd_codes", [])
    api_descs = r.get("pkd_descriptions", {})
    pkd_descs = [api_descs.get(c, "") or get_pkd_description(c) for c in pkd_codes]

    return {
        "old_name": old_name,
        "input_city_zip": city_zip,
        "input_nip": input_nip,
        "search_method": candidate.get("search_method", ""),
        "match_confidence": candidate.get("match_confidence", ""),
        "current_name": candidate.get("current_name", ""),
        "name_changed": candidate.get("name_changed", ""),
        "nip": r.get("nip", ""),
        "regon": r.get("regon", ""),
        "krs": r.get("krs", ""),
        "status": r.get("status", ""),
        "is_active": r.get("is_active", ""),
        "street": r.get("street", ""),
        "building": r.get("building", ""),
        "api_zip_code": r.get("zip_code", ""),
        "api_city": r.get("city", ""),
        "full_address": r.get("address_str", ""),
        "pkd_main": r.get("pkd_main", ""),
        "pkd_codes": "; ".join(pkd_codes),
        "pkd_descriptions": "; ".join(pkd_descs),
        "date_started": r.get("date_started", ""),
        "email": r.get("email", ""),
        "phone": r.get("phone", ""),
        "website": r.get("website", ""),
        "name_history": candidate.get("name_history", ""),
        "api_source": r.get("source", ""),
        "web_search_snippet": candidate.get("web_snippet", ""),
    }


def main():
    parser = argparse.ArgumentParser(
        description="Find current names of Polish businesses that may have been renamed."
    )
    parser.add_argument("input_file", help="Excel: col A=old name, B=city/zip (optional), C=NIP (optional)")
    parser.add_argument("--output", "-o", default="name_changes.xlsx", help="Output Excel file")
    parser.add_argument("--limit", "-n", type=int, default=0, help="Only process first N rows (0=all)")
    parser.add_argument("--gus-key", help="GUS API key (or set GUS_API_KEY env var)")
    parser.add_argument("--max-candidates", type=int, default=3,
                        help="Max candidate matches per business (default: 3)")
    parser.add_argument("--web-search", action="store_true",
                        help="Enable web search fallback (slower, uses DuckDuckGo)")
    parser.add_argument("--sandbox", action="store_true", help="Use GUS sandbox")
    args = parser.parse_args()

    logger.info(f"Reading {args.input_file}")
    try:
        df_in = pd.read_excel(args.input_file, header=0)
    except FileNotFoundError:
        logger.error(f"File not found: {args.input_file}")
        sys.exit(1)

    if len(df_in.columns) < 1:
        logger.error("Need at least 1 column: A = old business name")
        sys.exit(1)

    name_col = df_in.columns[0]
    city_col = df_in.columns[1] if len(df_in.columns) >= 2 else None
    nip_col = df_in.columns[2] if len(df_in.columns) >= 3 else None
    logger.info(f"Columns: Name='{name_col}', City/Zip='{city_col}', NIP='{nip_col}'")

    # Build work list
    work = []
    for i, row in df_in.iterrows():
        name = str(row[name_col]).strip()
        city_zip = str(row[city_col]).strip() if city_col else ""
        nip = str(row[nip_col]).strip() if nip_col else ""
        if name.lower() in ("nan", "none", ""):
            continue
        if city_zip.lower() in ("nan", "none"):
            city_zip = ""
        if nip.lower() in ("nan", "none"):
            nip = ""
        work.append((i, name, city_zip, nip))

    if args.limit > 0:
        work = work[:args.limit]
        logger.info(f"Limited to first {args.limit} rows.")

    logger.info(f"Businesses to search: {len(work)}")

    if not work:
        logger.error("No business names found in input.")
        sys.exit(1)

    # Set up API client
    api_client = PolandAPIClient(gus_api_key=args.gus_key, use_sandbox=args.sandbox)
    if not api_client.gus_api_key:
        logger.error("No GUS API key. Set GUS_API_KEY env var or use --gus-key.")
        sys.exit(1)

    # Process
    all_rows = []
    best_rows = []
    stats = {"total": 0, "name_changed": 0, "same_name": 0, "not_found": 0}

    for idx, (orig_idx, old_name, city_zip, nip) in enumerate(work):
        logger.info(f"[{idx + 1}/{len(work)}] Old name: '{old_name}'")
        stats["total"] += 1

        candidates = search_current_name(
            api_client, old_name, city_zip, nip,
            do_web_search=args.web_search
        )

        if candidates:
            # Best match
            best = candidates[0]
            if best["name_changed"] == "YES":
                stats["name_changed"] += 1
                logger.info(f"  NAME CHANGED: '{old_name}' -> '{best['current_name']}' "
                            f"[{best['match_confidence']}]")
            else:
                stats["same_name"] += 1
                logger.info(f"  Same name: '{best['current_name']}' [{best['match_confidence']}]")

            best_rows.append(build_row(old_name, city_zip, nip, best))

            # All candidates
            for cand in candidates[:args.max_candidates]:
                all_rows.append(build_row(old_name, city_zip, nip, cand))
        else:
            stats["not_found"] += 1
            logger.info(f"  NOT FOUND")
            empty = {
                "old_name": old_name,
                "input_city_zip": city_zip,
                "input_nip": nip,
                "search_method": "ALL_FAILED",
                "match_confidence": "NOT FOUND",
                "current_name": "",
                "name_changed": "UNKNOWN",
            }
            best_rows.append(empty)
            all_rows.append(empty)

    # Output
    logger.info(f"Saving to {args.output}")
    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        # Sheet 1: Best match per business
        df_best = pd.DataFrame(best_rows, columns=OUTPUT_COLUMNS).fillna("")
        df_best.to_excel(writer, index=False, sheet_name="Best Match")

        # Sheet 2: All candidates
        df_all = pd.DataFrame(all_rows, columns=OUTPUT_COLUMNS).fillna("")
        df_all.to_excel(writer, index=False, sheet_name="All Candidates")

        # Sheet 3: Name changes only
        changed = [r for r in best_rows if r.get("name_changed") == "YES"]
        if changed:
            df_changed = pd.DataFrame(changed, columns=OUTPUT_COLUMNS).fillna("")
            df_changed.to_excel(writer, index=False, sheet_name="Name Changes")

    logger.info("=" * 55)
    logger.info(f"DONE  |  Total searched: {stats['total']}")
    logger.info(f"  Name changed:  {stats['name_changed']}")
    logger.info(f"  Same name:     {stats['same_name']}")
    logger.info(f"  Not found:     {stats['not_found']}")
    logger.info(f"Results: {args.output}")


if __name__ == "__main__":
    main()
