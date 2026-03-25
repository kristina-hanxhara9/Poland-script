#!/usr/bin/env python3
"""
Look up Polish businesses: first by NIP (tax ID), then fallback to name + zip.

Column A = NIP/tax ID (can be empty)
Column B = Business name
Column C = Zip code

Logic:
1. If NIP exists → search GUS by NIP → return all details
2. If no NIP → search GUS by name → pick best match (prefer zip match)
3. If name search fails → retry with simplified name / partial name

Usage:
    python lookup_combined.py input.xlsx --output results.xlsx --limit 2
"""

import argparse
import logging
import sys

from dotenv import load_dotenv
import pandas as pd

load_dotenv()

from poland_api import PolandAPIClient

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

OUTPUT_COLUMNS = [
    "input_nip",
    "input_name",
    "input_zip",
    "search_method",
    "match_status",
    "zip_match",
    "api_name",
    "nip",
    "regon",
    "krs",
    "status",
    "is_active",
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
    "pkd_main",
    "pkd_codes",
    "pkd_descriptions",
    "date_started",
    "date_ended",
    "owner_first_name",
    "owner_last_name",
    "email",
    "phone",
    "website",
    "api_source",
]


def get_pkd_description(code):
    try:
        from pkd_codes import get_pkd_description as _get_desc
        return _get_desc(code)
    except ImportError:
        return ""


def clean_nip(raw):
    s = str(raw).strip().replace("-", "").replace(" ", "")
    if s.lower() in ("nan", "none", "") or not s.isdigit():
        return ""
    return s


def normalize_zip(z):
    s = str(z).strip().replace("-", "").replace(" ", "")
    if s.lower() in ("nan", "none", ""):
        return ""
    return s


def zips_match(input_zip, api_zip):
    a = normalize_zip(input_zip)
    b = normalize_zip(api_zip)
    if not a or not b:
        return None
    return a == b


def simplify_name(name):
    name = name.lower().strip()
    for suffix in [
        " sp. z o.o. sp. k.", " sp. z o.o. sp.k.",
        " sp. z o.o.", " sp.z o.o.", " spółka z o.o.", " spolka z o.o.",
        " sp. z o. o.", " s.a.", " sp.j.", " sp.k.", " sp. komandytowa",
        " spółka jawna", " spolka jawna", " spółka akcyjna", " spolka akcyjna",
        " s.c.", " sp. cywilna", " spółka cywilna", " spolka cywilna",
        " z o.o.", " z.o.o.",
    ]:
        if name.endswith(suffix):
            name = name[:-len(suffix)].strip()
            break
    name = name.strip('"\'').strip()
    return name


def build_row(input_nip, input_name, input_zip, result, search_method, match_status, zip_match_str):
    pkd_codes = result.get("pkd_codes", [])
    api_descs = result.get("pkd_descriptions", {})
    pkd_descs = [api_descs.get(c, "") or get_pkd_description(c) for c in pkd_codes]

    return {
        "input_nip": input_nip,
        "input_name": input_name,
        "input_zip": input_zip,
        "search_method": search_method,
        "match_status": match_status,
        "zip_match": zip_match_str,
        "api_name": result.get("name", ""),
        "nip": result.get("nip", ""),
        "regon": result.get("regon", ""),
        "krs": result.get("krs", ""),
        "status": result.get("status", ""),
        "is_active": result.get("is_active", ""),
        "street": result.get("street", ""),
        "building": result.get("building", ""),
        "unit": result.get("unit", ""),
        "api_zip_code": result.get("zip_code", ""),
        "api_city": result.get("city", ""),
        "municipality": result.get("municipality", ""),
        "county": result.get("county", ""),
        "province": result.get("province", ""),
        "country": result.get("country", ""),
        "full_address": result.get("address_str", ""),
        "pkd_main": result.get("pkd_main", ""),
        "pkd_codes": "; ".join(pkd_codes),
        "pkd_descriptions": "; ".join(pkd_descs),
        "date_started": result.get("date_started", ""),
        "date_ended": result.get("date_ended", ""),
        "owner_first_name": result.get("owner_first_name", ""),
        "owner_last_name": result.get("owner_last_name", ""),
        "email": result.get("email", ""),
        "phone": result.get("phone", ""),
        "website": result.get("website", ""),
        "api_source": result.get("source", ""),
    }


def search_by_name_with_retries(api_client, name):
    """Search GUS by name with progressive fallbacks. Returns list of results."""
    # 1. Full exact name
    results = api_client.search_gus_by_name(name)
    if results:
        logger.info(f"  Found {len(results)} results with full name")
        return results

    # 2. Without legal form suffix
    simplified = simplify_name(name)
    if simplified != name.lower().strip():
        logger.info(f"  Retrying without legal form: '{simplified}'")
        results = api_client.search_gus_by_name(simplified)
        if results:
            logger.info(f"  Found {len(results)} results with simplified name")
            return results

    # 3. Progressive word removal (longest to shortest)
    words = name.strip().split()
    if len(words) > 1:
        for n in range(len(words) - 1, 0, -1):
            partial = " ".join(words[:n])
            if len(partial) < 3:
                continue
            logger.info(f"  Retrying with partial: '{partial}'")
            results = api_client.search_gus_by_name(partial)
            if results:
                logger.info(f"  Found {len(results)} results with partial name")
                return results

    return []


def pick_best_result(results, input_zip):
    """From a list of GUS results, pick the best one (prefer zip match)."""
    if not results:
        return None, None

    # First pass: exact zip match
    if input_zip:
        for r in results:
            if zips_match(input_zip, r.get("zip_code", "")):
                return r, True

    # No zip match — take first result
    best = results[0]
    if input_zip:
        zm = zips_match(input_zip, best.get("zip_code", ""))
    else:
        zm = None
    return best, zm


def enrich_with_krs(api_client, result):
    """If result has a KRS number, fetch extra data from KRS."""
    if result and result.get("krs"):
        krs_data = api_client.search_krs(result["krs"])
        if krs_data:
            return api_client._merge_results(result, krs_data)
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Look up Polish businesses: NIP first, then name + zip fallback."
    )
    parser.add_argument("input_file", help="Excel: col A = NIP, col B = name, col C = zip")
    parser.add_argument("--output", "-o", default="combined_results.xlsx")
    parser.add_argument("--gus-key", help="GUS API key (or set GUS_API_KEY env var)")
    parser.add_argument("--limit", "-n", type=int, default=0, help="Process first N rows only")
    parser.add_argument("--sandbox", action="store_true", help="Use GUS sandbox (test data)")
    args = parser.parse_args()

    logger.info(f"Reading {args.input_file}")
    try:
        df_in = pd.read_excel(args.input_file, header=0)
    except FileNotFoundError:
        logger.error(f"File not found: {args.input_file}")
        sys.exit(1)

    if len(df_in.columns) < 3:
        logger.error("Need 3 columns: A = NIP/tax ID, B = name, C = zip code.")
        sys.exit(1)

    nip_col = df_in.columns[0]
    name_col = df_in.columns[1]
    zip_col = df_in.columns[2]
    logger.info(f"Columns: NIP='{nip_col}', Name='{name_col}', Zip='{zip_col}'")

    # Build work list
    work = []
    for i, row in df_in.iterrows():
        nip = clean_nip(row[nip_col])
        name = str(row[name_col]).strip()
        zip_code = str(row[zip_col]).strip()
        if name.lower() in ("nan", "none", ""):
            name = ""
        if zip_code.lower() in ("nan", "none"):
            zip_code = ""
        work.append((i, nip, name, zip_code))

    if args.limit > 0:
        work = work[:args.limit]
        logger.info(f"Limited to first {args.limit} rows.")

    logger.info(f"Total rows: {len(work)}")

    # Set up API
    api_client = PolandAPIClient(gus_api_key=args.gus_key, use_sandbox=args.sandbox)
    if not api_client.gus_api_key:
        logger.error("No GUS API key. Set GUS_API_KEY env var or use --gus-key.")
        sys.exit(1)

    logger.info("GUS API key configured. KRS API always available.")

    # Process
    rows = []
    stats = {"nip_found": 0, "name_match": 0, "name_zip_diff": 0, "not_found": 0}

    for idx, (orig_idx, nip, name, input_zip) in enumerate(work):
        logger.info(f"[{idx + 1}/{len(work)}] NIP: {nip or '(empty)'} | Name: {name} | Zip: {input_zip}")

        result = None
        search_method = ""

        # --- Strategy 1: Search by NIP ---
        if nip:
            search_method = "NIP"
            result = api_client.search_gus_by_nip(nip)
            if result:
                result = enrich_with_krs(api_client, result)
                zm = zips_match(input_zip, result.get("zip_code", ""))

                if zm is True:
                    match_status = "FOUND BY NIP - ZIP MATCH"
                    zip_str = "YES"
                elif zm is False:
                    match_status = "FOUND BY NIP - ZIP DIFFERENT"
                    zip_str = f"NO (input: {input_zip}, api: {result.get('zip_code', '')})"
                else:
                    match_status = "FOUND BY NIP"
                    zip_str = "N/A"

                stats["nip_found"] += 1
                rows.append(build_row(nip, name, input_zip, result, search_method, match_status, zip_str))
                logger.info(f"  -> {result.get('name', '?')} | {match_status}")
                continue
            else:
                logger.info(f"  NIP search returned nothing, falling back to name...")

        # --- Strategy 2: Search by name ---
        if name:
            search_method = "NAME" if not nip else "NIP_FAILED->NAME"
            results = search_by_name_with_retries(api_client, name)

            if results:
                best, zm = pick_best_result(results, input_zip)
                best = enrich_with_krs(api_client, best)

                if zm is True:
                    match_status = "FOUND BY NAME - ZIP MATCH"
                    zip_str = "YES"
                    stats["name_match"] += 1
                elif zm is False:
                    match_status = "FOUND BY NAME - ZIP DIFFERENT"
                    zip_str = f"NO (input: {input_zip}, api: {best.get('zip_code', '')})"
                    stats["name_zip_diff"] += 1
                else:
                    match_status = "FOUND BY NAME"
                    zip_str = "N/A"
                    stats["name_match"] += 1

                rows.append(build_row(nip, name, input_zip, best, search_method, match_status, zip_str))
                logger.info(f"  -> {best.get('name', '?')} | {match_status}")
                continue

        # --- Nothing found ---
        stats["not_found"] += 1
        rows.append({
            "input_nip": nip,
            "input_name": name,
            "input_zip": input_zip,
            "search_method": search_method or "NONE",
            "match_status": "NOT FOUND",
            "zip_match": "",
        })
        logger.info(f"  -> NOT FOUND")

    # Output
    df_out = pd.DataFrame(rows, columns=OUTPUT_COLUMNS).fillna("")

    logger.info(f"Saving to {args.output}")
    df_out.to_excel(args.output, index=False)

    logger.info("=" * 55)
    logger.info(f"DONE  |  Total: {len(work)}")
    logger.info(f"  Found by NIP:              {stats['nip_found']}")
    logger.info(f"  Found by name (zip match): {stats['name_match']}")
    logger.info(f"  Found by name (zip diff):  {stats['name_zip_diff']}")
    logger.info(f"  Not found:                 {stats['not_found']}")
    logger.info(f"Results: {args.output}")


if __name__ == "__main__":
    main()
