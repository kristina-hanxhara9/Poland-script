#!/usr/bin/env python3
"""
Look up Polish businesses by zip code (then street as fallback) using GUS API.

Reads an Excel file:
  Column A = Business name (for reference)
  Column B = Zip code (primary search)
  Column C = Street (fallback if zip returns nothing)

Usage:
    python lookup_by_zip.py input.xlsx --output zip_results.xlsx --limit 2
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
    "input_name",
    "input_zip",
    "input_street",
    "search_method",
    "match_status",
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


def normalize_zip(z):
    s = str(z).strip().replace("-", "").replace(" ", "")
    if s.lower() in ("nan", "none", ""):
        return ""
    return s


def simplify_name(name):
    """Remove Polish legal form suffixes from a business name."""
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


def search_by_name_fuzzy(api_client, name):
    """Search GUS by name with progressive fallbacks. Returns list of results."""
    if not name:
        return []

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


def build_row(input_name, input_zip, input_street, result, search_method, match_status):
    pkd_codes = result.get("pkd_codes", [])
    api_descs = result.get("pkd_descriptions", {})
    pkd_descs = [api_descs.get(c, "") or get_pkd_description(c) for c in pkd_codes]

    return {
        "input_name": input_name,
        "input_zip": input_zip,
        "input_street": input_street,
        "search_method": search_method,
        "match_status": match_status,
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


def main():
    parser = argparse.ArgumentParser(
        description="Look up Polish businesses by zip code via GUS API."
    )
    parser.add_argument("input_file", help="Excel file: col A = name, col B = zip, col C = street")
    parser.add_argument("--output", "-o", default="zip_results.xlsx", help="Output Excel file")
    parser.add_argument("--gus-key", help="GUS API key (or set GUS_API_KEY env var)")
    parser.add_argument("--limit", "-n", type=int, default=0, help="Only process first N rows. 0 = all.")
    parser.add_argument("--max-per-zip", type=int, default=10, help="Max businesses to return per zip code (default: 10)")
    parser.add_argument("--sandbox", action="store_true", help="Use GUS sandbox (test data)")
    args = parser.parse_args()

    logger.info(f"Reading {args.input_file}")
    try:
        df_in = pd.read_excel(args.input_file, header=0)
    except FileNotFoundError:
        logger.error(f"File not found: {args.input_file}")
        sys.exit(1)

    if len(df_in.columns) < 2:
        logger.error("Need at least 2 columns: A = name, B = zip code.")
        sys.exit(1)

    name_col = df_in.columns[0]
    zip_col = df_in.columns[1]
    street_col = df_in.columns[2] if len(df_in.columns) >= 3 else None
    logger.info(f"Columns: Name='{name_col}', Zip='{zip_col}', Street='{street_col}'")

    # Build work list
    work = []
    for i, row in df_in.iterrows():
        name = str(row[name_col]).strip()
        zip_code = str(row[zip_col]).strip()
        street = str(row[street_col]).strip() if street_col else ""
        if name.lower() in ("nan", "none", ""):
            name = ""
        if zip_code.lower() in ("nan", "none"):
            zip_code = ""
        if street.lower() in ("nan", "none", ""):
            street = ""
        norm = normalize_zip(zip_code)
        if norm or street or name:
            work.append((i, name, zip_code, street))

    if args.limit > 0:
        work = work[:args.limit]
        logger.info(f"Limited to first {args.limit} rows.")

    logger.info(f"Rows to process: {len(work)}")

    if not work:
        logger.error("No rows with valid zip or street found.")
        sys.exit(1)

    # Set up API client
    api_client = PolandAPIClient(gus_api_key=args.gus_key, use_sandbox=args.sandbox)
    if not api_client.gus_api_key:
        logger.error("No GUS API key. Set GUS_API_KEY env var or use --gus-key.")
        sys.exit(1)

    logger.info("GUS API key configured.")

    # Process — search by zip first, then street as fallback
    rows = []
    searched_zips = {}
    searched_streets = {}
    stats = {"zips_searched": 0, "streets_searched": 0, "businesses_found": 0, "not_found": 0}

    for idx, (orig_idx, name, input_zip, input_street) in enumerate(work):
        logger.info(f"[{idx + 1}/{len(work)}] Name: '{name}' | Zip: {input_zip} | Street: {input_street}")

        found_results = []
        search_method = ""

        # Strategy 1: Search by zip code
        norm_zip = normalize_zip(input_zip)
        if norm_zip:
            if norm_zip not in searched_zips:
                stats["zips_searched"] += 1
                logger.info(f"  Searching GUS by zip: {input_zip}")
                zip_results = api_client.search_gus_by_zip(input_zip, max_results=args.max_per_zip)
                searched_zips[norm_zip] = zip_results
            else:
                zip_results = searched_zips[norm_zip]
                logger.info(f"  Using cached results for zip {input_zip} ({len(zip_results)} businesses)")

            if zip_results:
                found_results = zip_results
                search_method = "ZIP_SEARCH"

        # Strategy 2: Fallback to street search if zip returned nothing
        if not found_results and input_street:
            street_key = input_street.lower()
            if street_key not in searched_streets:
                stats["streets_searched"] += 1
                logger.info(f"  ZIP empty/no results -> Searching GUS by street: {input_street}")
                street_results = api_client.search_gus_by_street(input_street, max_results=args.max_per_zip)
                searched_streets[street_key] = street_results
            else:
                street_results = searched_streets[street_key]
                logger.info(f"  Using cached results for street '{input_street}' ({len(street_results)} businesses)")

            if street_results:
                found_results = street_results
                search_method = "STREET_SEARCH"

        # Strategy 3: Fallback to fuzzy name search
        if not found_results and name:
            logger.info(f"  ZIP+STREET failed -> Searching GUS by name: {name}")
            stats["names_searched"] = stats.get("names_searched", 0) + 1
            name_results = search_by_name_fuzzy(api_client, name)
            if name_results:
                found_results = name_results
                search_method = "NAME_SEARCH"

        if found_results:
            stats["businesses_found"] += len(found_results)
            for r in found_results:
                if r.get("krs"):
                    krs_data = api_client.search_krs(r["krs"])
                    if krs_data:
                        r = api_client._merge_results(r, krs_data)

                if search_method == "ZIP_SEARCH":
                    match_label = f"FOUND BY ZIP ({input_zip})"
                elif search_method == "STREET_SEARCH":
                    match_label = f"FOUND BY STREET ({input_street})"
                else:
                    match_label = f"FOUND BY NAME ({name})"
                rows.append(build_row(
                    name, input_zip, input_street, r,
                    search_method, match_label
                ))
            logger.info(f"  -> Found {len(found_results)} businesses via {search_method}")
        else:
            stats["not_found"] += 1
            rows.append({
                "input_name": name,
                "input_zip": input_zip,
                "input_street": input_street,
                "search_method": "ZIP+STREET",
                "match_status": "NOT FOUND",
            })
            logger.info(f"  -> No businesses found by zip or street")

    # Output
    df_out = pd.DataFrame(rows, columns=OUTPUT_COLUMNS).fillna("")

    logger.info(f"Saving to {args.output}")
    df_out.to_excel(args.output, index=False)

    logger.info("=" * 55)
    logger.info(f"DONE  |  Total input rows: {len(work)}")
    logger.info(f"  Unique zips searched:    {stats['zips_searched']}")
    logger.info(f"  Unique streets searched: {stats['streets_searched']}")
    logger.info(f"  Names searched (fuzzy):  {stats.get('names_searched', 0)}")
    logger.info(f"  Businesses found:        {stats['businesses_found']}")
    logger.info(f"  Not found:               {stats['not_found']}")
    logger.info(f"Results: {args.output}")


if __name__ == "__main__":
    main()
