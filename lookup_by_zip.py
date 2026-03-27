#!/usr/bin/env python3
"""
Look up Polish businesses by zip code using GUS API.

Reads an Excel file:
  Column A = Business name (for reference)
  Column B = Zip code (used for search)

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


def build_row(input_name, input_zip, result, search_method, match_status):
    pkd_codes = result.get("pkd_codes", [])
    api_descs = result.get("pkd_descriptions", {})
    pkd_descs = [api_descs.get(c, "") or get_pkd_description(c) for c in pkd_codes]

    return {
        "input_name": input_name,
        "input_zip": input_zip,
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
    parser.add_argument("input_file", help="Excel file: col A = name, col B = zip")
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
    logger.info(f"Columns: Name='{name_col}', Zip='{zip_col}'")

    # Build work list
    work = []
    for i, row in df_in.iterrows():
        name = str(row[name_col]).strip()
        zip_code = str(row[zip_col]).strip()
        if name.lower() in ("nan", "none", ""):
            name = ""
        if zip_code.lower() in ("nan", "none"):
            zip_code = ""
        norm = normalize_zip(zip_code)
        if norm:
            work.append((i, name, zip_code))

    if args.limit > 0:
        work = work[:args.limit]
        logger.info(f"Limited to first {args.limit} rows.")

    logger.info(f"Rows with valid zip codes: {len(work)}")

    if not work:
        logger.error("No rows with valid zip codes found.")
        sys.exit(1)

    # Set up API client
    api_client = PolandAPIClient(gus_api_key=args.gus_key, use_sandbox=args.sandbox)
    if not api_client.gus_api_key:
        logger.error("No GUS API key. Set GUS_API_KEY env var or use --gus-key.")
        sys.exit(1)

    logger.info("GUS API key configured.")

    # Process — search by zip, cache results per unique zip
    rows = []
    searched_zips = {}
    stats = {"zips_searched": 0, "businesses_found": 0, "empty_zips": 0}

    for idx, (orig_idx, name, input_zip) in enumerate(work):
        logger.info(f"[{idx + 1}/{len(work)}] Name: '{name}' | Zip: {input_zip}")

        norm_zip = normalize_zip(input_zip)

        # Use cached results if we already searched this zip
        if norm_zip not in searched_zips:
            stats["zips_searched"] += 1
            logger.info(f"  Searching GUS by zip: {input_zip}")
            zip_results = api_client.search_gus_by_zip(input_zip, max_results=args.max_per_zip)
            searched_zips[norm_zip] = zip_results
        else:
            zip_results = searched_zips[norm_zip]
            logger.info(f"  Using cached results for zip {input_zip} ({len(zip_results)} businesses)")

        if zip_results:
            stats["businesses_found"] += len(zip_results)
            for r in zip_results:
                # Enrich with KRS if available
                if r.get("krs"):
                    krs_data = api_client.search_krs(r["krs"])
                    if krs_data:
                        r = api_client._merge_results(r, krs_data)

                rows.append(build_row(
                    name, input_zip, r,
                    "ZIP_SEARCH",
                    f"FOUND BY ZIP ({input_zip})"
                ))
            logger.info(f"  -> Found {len(zip_results)} businesses at zip {input_zip}")
        else:
            stats["empty_zips"] += 1
            rows.append({
                "input_name": name,
                "input_zip": input_zip,
                "search_method": "ZIP_SEARCH",
                "match_status": "NO BUSINESSES AT ZIP",
            })
            logger.info(f"  -> No businesses found at zip {input_zip}")

    # Output
    df_out = pd.DataFrame(rows, columns=OUTPUT_COLUMNS).fillna("")

    logger.info(f"Saving to {args.output}")
    df_out.to_excel(args.output, index=False)

    logger.info("=" * 55)
    logger.info(f"DONE  |  Total input rows: {len(work)}")
    logger.info(f"  Unique zips searched:   {stats['zips_searched']}")
    logger.info(f"  Businesses found:       {stats['businesses_found']}")
    logger.info(f"  Empty zips (no results):{stats['empty_zips']}")
    logger.info(f"Results: {args.output}")


if __name__ == "__main__":
    main()
