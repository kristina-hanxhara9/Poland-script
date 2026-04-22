#!/usr/bin/env python3
"""
Look up Polish businesses by name + zip code using GUS + KRS APIs.

Reads an Excel file with business names in column A and zip codes in column B.
Searches GUS by name, checks if the zip matches, and returns all details.

Usage:
    python lookup_by_name.py input.xlsx --output results.xlsx --limit 2

Env:
    GUS_API_KEY  - your GUS BIR1 API key
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
    "match_status",
    "zip_match",
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


def simplify_name(name):
    """Simplify a business name for retry search — remove legal form suffixes, extra words."""
    name = name.lower().strip()
    # Remove common Polish legal form suffixes
    for suffix in [
        " sp. z o.o.", " sp.z o.o.", " spółka z o.o.", " spolka z o.o.",
        " sp. z o. o.", " s.a.", " sp.j.", " sp.k.", " sp. komandytowa",
        " spółka jawna", " spolka jawna", " spółka akcyjna", " spolka akcyjna",
        " s.c.", " sp. cywilna", " spółka cywilna", " spolka cywilna",
        " sp. z o.o. sp. k.", " sp. z o.o. sp.k.",
        " z o.o.", " z.o.o.",
    ]:
        if name.endswith(suffix):
            name = name[:-len(suffix)].strip()
            break
    # Remove quotes and extra punctuation
    name = name.strip('"\'').strip()
    return name


def normalize_zip(z):
    """Normalize zip code: strip whitespace, dashes, etc."""
    s = str(z).strip().replace("-", "").replace(" ", "")
    if s.lower() in ("nan", "none", ""):
        return ""
    return s


def zips_match(input_zip, api_zip):
    """Check if two zip codes match (ignoring dashes/spaces)."""
    a = normalize_zip(input_zip)
    b = normalize_zip(api_zip)
    if not a or not b:
        return None  # can't determine
    return a == b


def build_row(input_name, input_zip, result, match_status, zip_match_str):
    """Build an output row dict from a result."""
    pkd_codes = result.get("pkd_codes", [])
    api_descs = result.get("pkd_descriptions", {})
    pkd_descs = [api_descs.get(c, "") or get_pkd_description(c) for c in pkd_codes]

    return {
        "input_name": input_name,
        "input_zip": input_zip,
        "match_status": match_status,
        "zip_match": zip_match_str,
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
        description="Look up Polish businesses by name + zip via GUS + KRS APIs."
    )
    parser.add_argument("input_file", help="Excel file with names in col A, zip in col B")
    parser.add_argument(
        "--output", "-o", default="name_results.xlsx",
        help="Output Excel file (default: name_results.xlsx)",
    )
    parser.add_argument(
        "--gus-key", help="GUS API key (or set GUS_API_KEY env var)",
    )
    parser.add_argument(
        "--limit", "-n", type=int, default=0,
        help="Only process first N rows (for testing). 0 = all rows.",
    )
    parser.add_argument(
        "--sandbox", action="store_true",
        help="Use GUS sandbox (test data, no real key needed)",
    )
    args = parser.parse_args()

    # Load input
    logger.info(f"Reading {args.input_file}")
    try:
        df_in = pd.read_excel(args.input_file, header=0)
    except FileNotFoundError:
        logger.error(f"File not found: {args.input_file}")
        sys.exit(1)

    if len(df_in.columns) < 2:
        logger.error("Need at least 2 columns: name (A) and zip (B).")
        sys.exit(1)

    name_col = df_in.columns[0]
    zip_col = df_in.columns[1]
    logger.info(f"Name column: '{name_col}', Zip column: '{zip_col}'")
    logger.info(f"Total rows: {len(df_in)}")

    # Build work list
    work = []
    for i, row in df_in.iterrows():
        name = str(row[name_col]).strip()
        zip_code = str(row[zip_col]).strip()
        if name.lower() in ("nan", "none", ""):
            continue
        if zip_code.lower() in ("nan", "none"):
            zip_code = ""
        work.append((i, name, zip_code))

    if args.limit > 0:
        work = work[:args.limit]
        logger.info(f"Limited to first {args.limit} rows for testing.")

    logger.info(f"Valid rows to process: {len(work)}")

    if not work:
        logger.error("No valid rows found.")
        sys.exit(1)

    # Set up API client
    api_client = PolandAPIClient(
        gus_api_key=args.gus_key,
        use_sandbox=args.sandbox,
    )

    if not api_client.gus_api_key:
        logger.error("No GUS API key. Set GUS_API_KEY env var or use --gus-key.")
        sys.exit(1)

    logger.info("GUS API key configured.")

    # Process
    rows = []
    found_match = 0
    found_nomatch = 0
    not_found = 0

    for idx, (orig_idx, name, input_zip) in enumerate(work):
        logger.info(f"[{idx + 1}/{len(work)}] Searching: '{name}' (zip: {input_zip})")

        results = api_client.search_gus_by_name(name)
        logger.info(f"  Full name search returned {len(results)} results")

        # If no results, try with simplified name (remove legal form)
        if not results:
            simplified = simplify_name(name)
            if simplified != name.lower().strip():
                logger.info(f"  Retrying without legal form: '{simplified}'")
                results = api_client.search_gus_by_name(simplified)
                logger.info(f"  Simplified search returned {len(results)} results")

        # If still no results, try searching word by word (longest first)
        if not results:
            words = name.strip().split()
            if len(words) > 1:
                # Try first N words, decreasing
                for n in range(len(words), 0, -1):
                    partial = " ".join(words[:n])
                    if len(partial) < 3:
                        continue
                    logger.info(f"  Retrying with partial name: '{partial}'")
                    results = api_client.search_gus_by_name(partial)
                    if results:
                        logger.info(f"  Partial search returned {len(results)} results")
                        break

        if not results:
            not_found += 1
            rows.append({
                "input_name": name,
                "input_zip": input_zip,
                "match_status": "NOT FOUND",
                "zip_match": "",
            })
            logger.info(f"  -> NOT FOUND")
            continue

        # Find best match considering zip
        best = None
        best_zip_match = None

        # First pass: look for exact zip match
        if input_zip:
            for r in results:
                if zips_match(input_zip, r.get("zip_code", "")):
                    best = r
                    best_zip_match = True
                    break

        # If no zip match, take the best name match (first result)
        if not best:
            best = results[0]
            if input_zip:
                best_zip_match = zips_match(input_zip, best.get("zip_code", ""))
            else:
                best_zip_match = None

        # Enrich with KRS if available
        if best.get("krs"):
            krs_data = api_client.search_krs(best["krs"])
            if krs_data:
                best = api_client._merge_results(best, krs_data)

        # Determine match status
        if best_zip_match is True:
            match_status = "MATCH"
            zip_match_str = "YES"
            found_match += 1
        elif best_zip_match is False:
            match_status = "FOUND - ZIP DIFFERENT"
            zip_match_str = f"NO (input: {input_zip}, api: {best.get('zip_code', '')})"
            found_nomatch += 1
        else:
            match_status = "FOUND - NO ZIP TO COMPARE"
            zip_match_str = "N/A"
            found_nomatch += 1

        rows.append(build_row(name, input_zip, best, match_status, zip_match_str))

        logger.info(
            f"  -> {best.get('name', '?')} | {match_status} | "
            f"zip: {best.get('zip_code', '')} | PKD: {best.get('pkd_main', '')}"
        )

    # Build output
    df_out = pd.DataFrame(rows, columns=OUTPUT_COLUMNS).fillna("")

    logger.info(f"Saving results to {args.output}")
    df_out.to_excel(args.output, index=False)

    # Summary
    logger.info("=" * 50)
    logger.info(f"DONE  |  Total: {len(work)}")
    logger.info(f"  MATCH (zip confirmed): {found_match}")
    logger.info(f"  FOUND (zip different or N/A): {found_nomatch}")
    logger.info(f"  NOT FOUND: {not_found}")
    logger.info(f"Results saved to: {args.output}")


if __name__ == "__main__":
    main()
