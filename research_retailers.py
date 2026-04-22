#!/usr/bin/env python3
"""
Enrich Polish retailer data with full API fields.

Reads an Excel file (output from lookup_by_zip.py, lookup_by_name.py, etc.)
and re-fetches each company by NIP from GUS + KRS APIs to get ALL fields:
  - Start date, suspension date, resumption date, closure date, deletion date
  - Number of local units
  - Legal form, ownership form, registry type
  - Full PKD codes with descriptions
  - Contact: email, phone, fax, website
  - Owner name (for sole traders)

Works with both column naming styles:
  - input_nip / input_name (from lookup_by_zip, lookup_by_name)
  - nip / name (from lookup_by_nip)

Usage:
    python research_retailers.py testcombined.xlsx --output testcombined_enriched.xlsx
    python research_retailers.py resultsmobile.xlsx --output resultsmobile_enriched.xlsx
    python research_retailers.py results.xlsx --limit 10
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
    "input_nip",
    "category",
    "api_name",
    "nip",
    "regon",
    "krs",
    "entity_type",
    "status",
    "is_active",
    "date_started",
    "date_created",
    "date_suspended",
    "date_resumed",
    "date_ended",
    "date_deleted",
    "date_last_change",
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
    "owner_first_name",
    "owner_last_name",
    "email",
    "phone",
    "fax",
    "website",
    "api_source",
    "enrichment_status",
]


def get_pkd_description(code):
    try:
        from pkd_codes import get_pkd_description as _get_desc
        return _get_desc(code)
    except ImportError:
        return ""


def find_column(df, candidates, required=False, label=""):
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
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.lower() in ("nan", "none"):
        return ""
    return s


def build_row(input_name, input_nip, category, result):
    pkd_codes = result.get("pkd_codes", [])
    api_descs = result.get("pkd_descriptions", {})
    if isinstance(api_descs, dict):
        pkd_descs = [api_descs.get(c, "") or get_pkd_description(c) for c in pkd_codes]
    else:
        pkd_descs = []

    return {
        "input_name": input_name,
        "input_nip": input_nip,
        "category": category,
        "api_name": result.get("name", ""),
        "nip": result.get("nip", ""),
        "regon": result.get("regon", ""),
        "krs": result.get("krs", ""),
        "entity_type": result.get("entity_type", ""),
        "status": result.get("status", ""),
        "is_active": result.get("is_active", ""),
        "date_started": result.get("date_started", ""),
        "date_created": result.get("date_created", ""),
        "date_suspended": result.get("date_suspended", ""),
        "date_resumed": result.get("date_resumed", ""),
        "date_ended": result.get("date_ended", ""),
        "date_deleted": result.get("date_deleted", ""),
        "date_last_change": result.get("date_last_change", ""),
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
        "pkd_codes": "; ".join(pkd_codes) if isinstance(pkd_codes, list) else pkd_codes,
        "pkd_descriptions": "; ".join(pkd_descs),
        "owner_first_name": result.get("owner_first_name", ""),
        "owner_last_name": result.get("owner_last_name", ""),
        "email": result.get("email", ""),
        "phone": result.get("phone", ""),
        "fax": result.get("fax", ""),
        "website": result.get("website", ""),
        "api_source": result.get("source", ""),
        "enrichment_status": "OK",
    }


def main():
    parser = argparse.ArgumentParser(
        description="Enrich Polish retailer data with full GUS + KRS API fields."
    )
    parser.add_argument("input_file", help="Excel file with retailer data (needs NIP column)")
    parser.add_argument("--output", "-o", default="enriched_retailers.xlsx",
                        help="Output Excel file (default: enriched_retailers.xlsx)")
    parser.add_argument("--limit", "-n", type=int, default=0,
                        help="Only process first N rows (0 = all)")
    parser.add_argument("--gus-key", help="GUS API key (or set GUS_API_KEY env var)")
    parser.add_argument("--sandbox", action="store_true", help="Use GUS sandbox")
    args = parser.parse_args()

    logger.info(f"Reading {args.input_file}")
    try:
        df = pd.read_excel(args.input_file, header=0)
    except FileNotFoundError:
        logger.error(f"File not found: {args.input_file}")
        sys.exit(1)

    logger.info(f"Loaded {len(df)} rows, {len(df.columns)} columns")
    logger.info(f"Columns: {list(df.columns)}")

    # Detect columns — try input_* variants first
    name_col = find_column(df, ["input_name", "api_name", "name", "chain_name", "company_name"],
                           required=True, label="Name")
    nip_col = find_column(df, ["input_nip", "nip", "NIP"],
                          required=True, label="NIP")
    cat_col = find_column(df, ["category", "channel"])

    logger.info(f"Using: Name='{name_col}', NIP='{nip_col}', Category='{cat_col}'")

    # Set up API client
    api_client = PolandAPIClient(gus_api_key=args.gus_key, use_sandbox=args.sandbox)
    if not api_client.gus_api_key:
        logger.error("No GUS API key. Set GUS_API_KEY env var or use --gus-key.")
        sys.exit(1)

    logger.info("GUS API key configured. KRS API: always available (free).")

    rows_to_process = df if args.limit == 0 else df.head(args.limit)
    total = len(rows_to_process)
    logger.info(f"Processing {total} retailers")

    all_rows = []
    stats = {"found": 0, "not_found": 0, "no_nip": 0, "errors": 0}

    for idx, (_, row) in enumerate(rows_to_process.iterrows()):
        name = normalize(row.get(name_col, ""))
        nip_val = normalize(row.get(nip_col, "")).replace("-", "").replace(" ", "")
        category = normalize(row.get(cat_col, "")) if cat_col else ""

        if not nip_val or len(nip_val) < 10:
            stats["no_nip"] += 1
            all_rows.append({
                "input_name": name,
                "input_nip": nip_val,
                "category": category,
                "enrichment_status": "NO VALID NIP",
            })
            logger.info(f"[{idx+1}/{total}] {name} — no valid NIP ('{nip_val}'), skipping")
            continue

        logger.info(f"[{idx+1}/{total}] {name} — NIP: {nip_val}")

        # GUS lookup by NIP
        try:
            result = api_client.search_gus_by_nip(nip_val)
        except Exception as e:
            stats["errors"] += 1
            all_rows.append({
                "input_name": name,
                "input_nip": nip_val,
                "category": category,
                "enrichment_status": f"GUS ERROR: {e}",
            })
            logger.error(f"  GUS error: {e}")
            continue

        if not result:
            stats["not_found"] += 1
            all_rows.append({
                "input_name": name,
                "input_nip": nip_val,
                "category": category,
                "enrichment_status": "NOT FOUND IN GUS",
            })
            logger.info(f"  -> Not found in GUS")
            continue

        # Enrich with KRS if KRS number available
        if result.get("krs"):
            try:
                krs_data = api_client.search_krs(result["krs"])
                if krs_data:
                    result = api_client._merge_results(result, krs_data)
                    logger.info(f"  KRS enrichment OK (KRS: {result['krs']})")
            except Exception as e:
                logger.warning(f"  KRS lookup failed: {e}")

        stats["found"] += 1
        all_rows.append(build_row(name, nip_val, category, result))

        logger.info(f"  -> {result.get('name', '?')} | {result.get('city', '')} | "
                     f"PKD: {result.get('pkd_main', '')} | "
                     f"Local units: {result.get('num_local_units', 'N/A')} | "
                     f"Suspended: {result.get('date_suspended', '') or 'No'}")

    # Save to Excel
    df_out = pd.DataFrame(all_rows, columns=OUTPUT_COLUMNS).fillna("")

    logger.info(f"Saving {len(df_out)} rows to {args.output}")
    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False, sheet_name="All Results")

        # Sheet 2: active businesses only
        active = df_out[df_out["is_active"].astype(str).str.lower().isin(["true", "1", "yes"])]
        if not active.empty:
            active.to_excel(writer, index=False, sheet_name="Active")
            logger.info(f"  Sheet 'Active': {len(active)} rows")

        # Sheet 3: suspended or closed
        suspended = df_out[df_out["date_suspended"].astype(str).str.len() > 0]
        if not suspended.empty:
            suspended.to_excel(writer, index=False, sheet_name="Suspended")
            logger.info(f"  Sheet 'Suspended': {len(suspended)} rows")

        # Sheet 4: with local units > 0
        has_units = df_out[df_out["num_local_units"].astype(str).str.match(r'^[1-9]')]
        if not has_units.empty:
            has_units.to_excel(writer, index=False, sheet_name="Has Local Units")
            logger.info(f"  Sheet 'Has Local Units': {len(has_units)} rows")

        # Sheet 5: not found / errors
        failed = df_out[df_out["enrichment_status"] != "OK"]
        if not failed.empty:
            failed.to_excel(writer, index=False, sheet_name="Not Found")
            logger.info(f"  Sheet 'Not Found': {len(failed)} rows")

        # Sheet 6: summary by category
        if cat_col:
            summary = []
            for cat in df_out["category"].unique():
                if not cat:
                    continue
                cat_rows = df_out[df_out["category"] == cat]
                cat_ok = cat_rows[cat_rows["enrichment_status"] == "OK"]
                cat_active = cat_ok[cat_ok["is_active"].astype(str).str.lower().isin(["true", "1", "yes"])]
                cat_suspended = cat_ok[cat_ok["date_suspended"].astype(str).str.len() > 0]
                summary.append({
                    "category": cat,
                    "total": len(cat_rows),
                    "found": len(cat_ok),
                    "not_found": len(cat_rows) - len(cat_ok),
                    "active": len(cat_active),
                    "suspended": len(cat_suspended),
                })
            if summary:
                pd.DataFrame(summary).to_excel(writer, index=False, sheet_name="Summary")

    logger.info("=" * 55)
    logger.info(f"DONE  |  Total: {total}")
    logger.info(f"  Found:     {stats['found']}")
    logger.info(f"  Not found: {stats['not_found']}")
    logger.info(f"  No NIP:    {stats['no_nip']}")
    logger.info(f"  Errors:    {stats['errors']}")
    logger.info(f"Output: {args.output}")


if __name__ == "__main__":
    main()
