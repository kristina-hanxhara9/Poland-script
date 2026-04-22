#!/usr/bin/env python3
"""
Look up Polish businesses by street name using GUS API.

Reads an Excel file:
  Column A = Business name (for reference, not used in search)
  Column B = Zip code (for reference)
  Column C = Street name (used for search)
  Column D = Street number (optional, used to filter results)

Usage:
    python lookup_by_street.py input.xlsx --output street_results.xlsx
    python lookup_by_street.py input.xlsx --limit 3
    python lookup_by_street.py input.xlsx --max-per-street 20
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
    "input_street_number",
    "search_method",
    "match_status",
    "number_match",
    "api_name",
    "nip",
    "regon",
    "krs",
    "entity_type",
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
    "legal_form_code",
    "legal_form_name",
    "legal_form_specific",
    "ownership_form",
    "size_category",
    "registry_type",
    "registration_number",
    "num_local_units",
    "pkd_main",
    "pkd_codes",
    "pkd_descriptions",
    "date_started",
    "date_created",
    "date_suspended",
    "date_resumed",
    "date_ended",
    "date_deleted",
    "date_last_change",
    "owner_first_name",
    "owner_last_name",
    "email",
    "phone",
    "fax",
    "website",
    "api_source",
]


def get_pkd_description(code):
    try:
        from pkd_codes import get_pkd_description as _get_desc
        return _get_desc(code)
    except ImportError:
        return ""


def normalize(s):
    """Clean a cell value."""
    if pd.isna(s) or s is None:
        return ""
    return str(s).strip()


def normalize_number(n):
    """Normalize a street/building number for comparison."""
    s = normalize(n).upper().replace(" ", "")
    # Remove trailing letters (e.g. "12A" -> "12" for loose matching)
    return s


def numbers_match(input_num, api_building):
    """Check if input street number matches API building number."""
    if not input_num:
        return True  # no filter requested
    inp = normalize_number(input_num)
    api = normalize_number(api_building)
    if not inp or not api:
        return not inp  # match if input is empty
    # Exact match
    if inp == api:
        return True
    # Loose: numeric part matches
    inp_digits = "".join(c for c in inp if c.isdigit())
    api_digits = "".join(c for c in api if c.isdigit())
    if inp_digits and api_digits and inp_digits == api_digits:
        return True
    return False


def build_row(input_name, input_zip, input_street, input_number, result, search_method, match_status, num_match):
    pkd_codes = result.get("pkd_codes", [])
    api_descs = result.get("pkd_descriptions", {})
    pkd_descs = [api_descs.get(c, "") or get_pkd_description(c) for c in pkd_codes]

    return {
        "input_name": input_name,
        "input_zip": input_zip,
        "input_street": input_street,
        "input_street_number": input_number,
        "search_method": search_method,
        "match_status": match_status,
        "number_match": "YES" if num_match else "NO",
        "api_name": result.get("name", ""),
        "nip": result.get("nip", ""),
        "regon": result.get("regon", ""),
        "krs": result.get("krs", ""),
        "entity_type": result.get("entity_type", ""),
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
        "legal_form_code": result.get("legal_form_code", ""),
        "legal_form_name": result.get("legal_form_name", ""),
        "legal_form_specific": result.get("legal_form_specific", ""),
        "ownership_form": result.get("ownership_form", ""),
        "size_category": result.get("size_category", ""),
        "registry_type": result.get("registry_type", ""),
        "registration_number": result.get("registration_number", ""),
        "num_local_units": result.get("num_local_units", ""),
        "pkd_main": result.get("pkd_main", ""),
        "pkd_codes": "; ".join(pkd_codes),
        "pkd_descriptions": "; ".join(pkd_descs),
        "date_started": result.get("date_started", ""),
        "date_created": result.get("date_created", ""),
        "date_suspended": result.get("date_suspended", ""),
        "date_resumed": result.get("date_resumed", ""),
        "date_ended": result.get("date_ended", ""),
        "date_deleted": result.get("date_deleted", ""),
        "date_last_change": result.get("date_last_change", ""),
        "owner_first_name": result.get("owner_first_name", ""),
        "owner_last_name": result.get("owner_last_name", ""),
        "email": result.get("email", ""),
        "phone": result.get("phone", ""),
        "fax": result.get("fax", ""),
        "website": result.get("website", ""),
        "api_source": result.get("source", ""),
    }


def main():
    parser = argparse.ArgumentParser(
        description="Look up Polish businesses by street name via GUS API."
    )
    parser.add_argument("input_file", help="Excel file: col A=name, B=zip, C=street, D=street_number")
    parser.add_argument("--output", "-o", default="street_results.xlsx", help="Output Excel file")
    parser.add_argument("--gus-key", help="GUS API key (or set GUS_API_KEY env var)")
    parser.add_argument("--limit", "-n", type=int, default=0, help="Only process first N rows (0=all)")
    parser.add_argument("--max-per-street", type=int, default=15, help="Max results per street (default: 15)")
    parser.add_argument("--filter-number", action="store_true",
                        help="Only show results matching the street number (default: show all, mark match)")
    parser.add_argument("--sandbox", action="store_true", help="Use GUS sandbox")
    args = parser.parse_args()

    logger.info(f"Reading {args.input_file}")
    try:
        df_in = pd.read_excel(args.input_file, header=0)
    except FileNotFoundError:
        logger.error(f"File not found: {args.input_file}")
        sys.exit(1)

    if len(df_in.columns) < 3:
        logger.error("Need at least 3 columns: A=name, B=zip, C=street")
        sys.exit(1)

    name_col = df_in.columns[0]
    zip_col = df_in.columns[1]
    street_col = df_in.columns[2]
    number_col = df_in.columns[3] if len(df_in.columns) >= 4 else None
    logger.info(f"Columns: Name='{name_col}', Zip='{zip_col}', "
                f"Street='{street_col}', Number='{number_col}'")

    # Build work list
    work = []
    for i, row in df_in.iterrows():
        name = normalize(row[name_col])
        zip_code = normalize(row[zip_col])
        street = normalize(row[street_col])
        number = normalize(row[number_col]) if number_col else ""

        if name.lower() in ("nan", "none"):
            name = ""
        if zip_code.lower() in ("nan", "none"):
            zip_code = ""
        if street.lower() in ("nan", "none"):
            street = ""
        if number.lower() in ("nan", "none"):
            number = ""

        if street:
            work.append((i, name, zip_code, street, number))

    if args.limit > 0:
        work = work[:args.limit]
        logger.info(f"Limited to first {args.limit} rows.")

    logger.info(f"Rows with street names: {len(work)}")

    if not work:
        logger.error("No rows with street names found.")
        sys.exit(1)

    # Set up API client
    api_client = PolandAPIClient(gus_api_key=args.gus_key, use_sandbox=args.sandbox)
    if not api_client.gus_api_key:
        logger.error("No GUS API key. Set GUS_API_KEY env var or use --gus-key.")
        sys.exit(1)

    logger.info("GUS API key configured.")

    # Process
    rows = []
    searched_streets = {}
    stats = {"streets_searched": 0, "businesses_found": 0, "number_matches": 0, "not_found": 0}

    for idx, (orig_idx, name, zip_code, street, number) in enumerate(work):
        logger.info(f"[{idx + 1}/{len(work)}] Name: '{name}' | Street: {street} {number}")

        street_key = street.lower().strip()

        # Search by street (cached)
        if street_key not in searched_streets:
            stats["streets_searched"] += 1
            logger.info(f"  Searching GUS by street: '{street}'")
            results = api_client.search_gus_by_street(street, max_results=args.max_per_street)
            searched_streets[street_key] = results
        else:
            results = searched_streets[street_key]
            logger.info(f"  Using cached results for street '{street}' ({len(results)} businesses)")

        if results:
            matched_count = 0
            for r in results:
                num_match = numbers_match(number, r.get("building", ""))

                if args.filter_number and not num_match:
                    continue  # skip non-matching numbers

                # Enrich with KRS if available
                if r.get("krs"):
                    krs_data = api_client.search_krs(r["krs"])
                    if krs_data:
                        r = api_client._merge_results(r, krs_data)

                match_label = f"FOUND ON {street}"
                if number:
                    match_label += f" {number}" if num_match else f" (searched {number})"

                rows.append(build_row(
                    name, zip_code, street, number, r,
                    "STREET_SEARCH", match_label, num_match
                ))
                matched_count += 1
                if num_match:
                    stats["number_matches"] += 1

            stats["businesses_found"] += matched_count
            logger.info(f"  -> {matched_count} businesses on '{street}'"
                        + (f" ({stats['number_matches']} at number {number})" if number else ""))
        else:
            stats["not_found"] += 1
            rows.append({
                "input_name": name,
                "input_zip": zip_code,
                "input_street": street,
                "input_street_number": number,
                "search_method": "STREET_SEARCH",
                "match_status": f"NO RESULTS FOR '{street}'",
                "number_match": "",
            })
            logger.info(f"  -> No businesses found on street '{street}'")

    # Output — two sheets: all results + number-matched only
    df_out = pd.DataFrame(rows, columns=OUTPUT_COLUMNS).fillna("")

    logger.info(f"Saving to {args.output}")
    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False, sheet_name="All Results")

        # Sheet 2: only rows where number matches (if number was provided)
        if number_col:
            matched = df_out[df_out["number_match"] == "YES"]
            if not matched.empty:
                matched.to_excel(writer, index=False, sheet_name="Number Matches")
                logger.info(f"  Sheet 'Number Matches': {len(matched)} rows")

    logger.info("=" * 55)
    logger.info(f"DONE  |  Total input rows: {len(work)}")
    logger.info(f"  Unique streets searched:  {stats['streets_searched']}")
    logger.info(f"  Businesses found:         {stats['businesses_found']}")
    logger.info(f"  Number matches:           {stats['number_matches']}")
    logger.info(f"  Streets with no results:  {stats['not_found']}")
    logger.info(f"Results: {args.output}")


if __name__ == "__main__":
    main()
