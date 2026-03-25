#!/usr/bin/env python3
"""
Look up Polish businesses by NIP code using GUS + KRS APIs.

Reads an Excel file with NIP codes in column A, fetches all available
business data, and writes results to a new Excel file with one column per field.

Usage:
    python lookup_by_nip.py input.xlsx --output results.xlsx

Env:
    GUS_API_KEY  - your GUS BIR1 API key (get free from https://api.stat.gov.pl)
"""

import argparse
import logging
import sys
import time

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

# Output columns in order
OUTPUT_COLUMNS = [
    "nip",
    "name",
    "regon",
    "krs",
    "status",
    "is_active",
    "street",
    "building",
    "unit",
    "zip_code",
    "city",
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
    "date_suspended",
    "date_resumed",
    "date_deleted",
    "owner_first_name",
    "owner_last_name",
    "email",
    "phone",
    "website",
    "api_source",
    "lookup_status",
]


def clean_nip(raw):
    """Normalize a NIP value: strip whitespace, dashes, and leading zeros."""
    s = str(raw).strip().replace("-", "").replace(" ", "")
    if s.lower() in ("nan", "none", ""):
        return ""
    return s


def lookup_single(api_client, nip):
    """Look up one NIP via GUS, then enrich with KRS if possible."""
    result = api_client.search_gus_by_nip(nip)

    if not result:
        return None

    # If GUS found a KRS number, get extra data from KRS
    if result.get("krs"):
        krs_data = api_client.search_krs(result["krs"])
        if krs_data:
            result = api_client._merge_results(result, krs_data)

    return result


def get_pkd_description(code):
    """Try to get PKD description, fallback to empty string."""
    try:
        from pkd_codes import get_pkd_description as _get_desc
        return _get_desc(code)
    except ImportError:
        return ""


def main():
    parser = argparse.ArgumentParser(
        description="Look up Polish businesses by NIP via GUS + KRS APIs."
    )
    parser.add_argument("input_file", help="Excel file with NIP codes in column A")
    parser.add_argument(
        "--output", "-o", default="nip_results.xlsx",
        help="Output Excel file (default: nip_results.xlsx)",
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

    # Get NIP from first column (column A)
    nip_column = df_in.columns[0]
    nips = df_in[nip_column].apply(clean_nip).tolist()
    logger.info(f"Found {len(nips)} rows. NIP column: '{nip_column}'")

    # Filter out empty NIPs
    valid_nips = [(i, n) for i, n in enumerate(nips) if n]

    if args.limit > 0:
        valid_nips = valid_nips[:args.limit]
        logger.info(f"Limited to first {args.limit} rows for testing.")

    logger.info(f"Valid NIP codes: {len(valid_nips)} / {len(nips)}")

    if not valid_nips:
        logger.error("No valid NIP codes found in column A.")
        sys.exit(1)

    # Set up API client
    api_client = PolandAPIClient(
        gus_api_key=args.gus_key,
        use_sandbox=args.sandbox,
    )

    if not api_client.gus_api_key:
        logger.error(
            "No GUS API key. Set GUS_API_KEY env var or use --gus-key, "
            "or use --sandbox for test data."
        )
        sys.exit(1)

    logger.info("GUS API key configured.")
    logger.info("KRS API: always available (free).")

    # Prepare output rows
    rows = []
    found = 0
    not_found = 0

    for idx, (orig_idx, nip) in enumerate(valid_nips):
        logger.info(f"[{idx + 1}/{len(valid_nips)}] Looking up NIP: {nip}")

        result = lookup_single(api_client, nip)

        if result:
            found += 1
            pkd_codes = result.get("pkd_codes", [])
            pkd_descs = [get_pkd_description(c) for c in pkd_codes]

            rows.append({
                "nip": nip,
                "name": result.get("name", ""),
                "regon": result.get("regon", ""),
                "krs": result.get("krs", ""),
                "status": result.get("status", ""),
                "is_active": result.get("is_active", ""),
                "street": result.get("street", ""),
                "building": result.get("building", ""),
                "unit": result.get("unit", ""),
                "zip_code": result.get("zip_code", ""),
                "city": result.get("city", ""),
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
                "date_suspended": result.get("date_suspended", ""),
                "date_resumed": result.get("date_resumed", ""),
                "date_deleted": result.get("date_deleted", ""),
                "owner_first_name": result.get("owner_first_name", ""),
                "owner_last_name": result.get("owner_last_name", ""),
                "email": result.get("email", ""),
                "phone": result.get("phone", ""),
                "website": result.get("website", ""),
                "api_source": result.get("source", ""),
                "lookup_status": "found",
            })
            logger.info(
                f"  -> {result.get('name', '?')} | {result.get('city', '')} | "
                f"PKD: {result.get('pkd_main', '')} | {result.get('status', '')}"
            )
        else:
            not_found += 1
            rows.append({
                "nip": nip,
                "lookup_status": "not_found",
            })
            logger.info(f"  -> NOT FOUND")

    # Build output dataframe
    df_out = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)

    # Fill NaN with empty string for clean output
    df_out = df_out.fillna("")

    # Save
    logger.info(f"Saving results to {args.output}")
    df_out.to_excel(args.output, index=False)

    # Summary
    logger.info("=" * 50)
    logger.info(f"DONE  |  Total: {len(valid_nips)}  |  Found: {found}  |  Not found: {not_found}")
    logger.info(f"Results saved to: {args.output}")


if __name__ == "__main__":
    main()
