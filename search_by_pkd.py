#!/usr/bin/env python3
"""
Find Polish businesses by PKD (SIC) code using GUS API.

GUS cannot search by PKD code directly, so this script:
  1. Searches GUS using Polish activity descriptions for each PKD code
  2. Searches GUS using broad business name keywords
  3. Filters ALL results to keep only those with matching PKD codes
  4. Enriches with KRS data (free, no key needed)
  5. Optionally fetches local units (branch locations)

This effectively gives you "all businesses with PKD code X" from the GUS database.

Mobile phone PKD codes:
  47.42.Z — Retail sale of telecommunications equipment (MAIN)
  46.52.Z — Wholesale of electronic & telecom equipment
  95.12.Z — Repair of communication equipment
  47.41.Z — Retail sale of computers, software
  61.10.Z — Wired telecom activities
  61.20.Z — Wireless telecom activities

Usage:
    # Find all mobile phone shops (by PKD codes)
    python search_by_pkd.py --preset mobile --output mobile_by_pkd.xlsx

    # Search specific PKD codes
    python search_by_pkd.py --pkd 47.42.Z 46.52.Z 95.12.Z --output telecom_retail.xlsx

    # With local units
    python search_by_pkd.py --preset mobile --local-units --output mobile_full.xlsx
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

# ─── PKD code presets ──────────────────────────────────────────────────────

PRESETS = {
    "mobile": {
        "name": "Mobile Phone Specialists",
        "pkd_codes": ["47.42.Z", "46.52.Z", "95.12.Z", "47.41.Z", "61.10.Z", "61.20.Z"],
        "keywords": [
            # Polish activity descriptions (from PKD classification)
            "sprzedaż detaliczna sprzętu telekomunikacyjnego",
            "sprzedaż hurtowa sprzętu elektronicznego i telekomunikacyjnego",
            "naprawa sprzętu komunikacyjnego",
            "sprzedaż detaliczna komputerów",
            "działalność telekomunikacyjna",
            "telekomunikacja bezprzewodowa",
            # Business name keywords
            "telefon komórkowy",
            "telefony komórkowe",
            "telefon komorkowy",
            "GSM",
            "smartfon",
            "salon telefoniczny",
            "sklep telefoniczny",
            "sklep GSM",
            "salon GSM",
            "akcesoria GSM",
            "akcesoria telefoniczne",
            "serwis GSM",
            "serwis telefonów",
            "naprawa telefonów",
            "hurtownia GSM",
            "hurtownia telefonów",
            "telekomunikacja",
            "telekomunikacyjny",
            # Chain names (to catch specific companies)
            "Teletorium",
            "Teleakces",
            "iSpot",
            "Cortland",
            "Orange Polska",
            "T-Mobile Polska",
            "P4",
            "Polkomtel",
            "Samsung Electronics Polska",
            "Huawei Polska",
            "Komputronik",
            "x-kom",
            "Media Expert",
            "TERG",
            "Euro AGD",
            "EURO-NET",
            "MPTECH",
            "Maxcom",
            "TelForceOne",
            "Xiaomi",
            "Mi Store",
            "NEONET",
            "Media Markt",
            "Cyfrowy Polsat",
            "ABC DATA",
            "ASBIS",
        ],
    },
    "diy": {
        "name": "DIY / Hardware Stores",
        "pkd_codes": ["47.52.Z", "47.53.Z", "47.59.Z", "47.54.Z", "46.73.Z", "46.74.Z"],
        "keywords": [
            "sklep budowlany",
            "materiały budowlane",
            "materialy budowlane",
            "hurtownia budowlana",
            "narzędzia",
            "artykuły budowlane",
            "farby lakiery",
            "Castorama",
            "Leroy Merlin",
            "OBI",
            "Bricomarche",
            "PSB",
            "Mrowka",
        ],
    },
}

OUTPUT_COLUMNS = [
    "target_pkd",
    "matched_pkd",
    "search_keyword",
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
    "zip_code",
    "city",
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
]


def get_pkd_description(code):
    try:
        from pkd_codes import get_pkd_description as _get_desc
        return _get_desc(code)
    except ImportError:
        return ""


def has_matching_pkd(result, target_codes):
    """Check if result has any of the target PKD codes. Returns matched codes."""
    pkd_main = result.get("pkd_main", "")
    pkd_codes = result.get("pkd_codes", [])
    all_codes = set(pkd_codes)
    if pkd_main:
        all_codes.add(pkd_main)

    matched = []
    for code in all_codes:
        for target in target_codes:
            if code.startswith(target.rstrip(".Z")):
                matched.append(code)
    return matched


def build_row(target_pkd, matched_pkd, keyword, result):
    pkd_codes = result.get("pkd_codes", [])
    api_descs = result.get("pkd_descriptions", {})
    if isinstance(api_descs, dict):
        pkd_descs = [api_descs.get(c, "") or get_pkd_description(c) for c in pkd_codes]
    else:
        pkd_descs = []

    return {
        "target_pkd": target_pkd,
        "matched_pkd": "; ".join(matched_pkd),
        "search_keyword": keyword,
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
        "zip_code": result.get("zip_code", ""),
        "city": result.get("city", ""),
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
    }


def main():
    parser = argparse.ArgumentParser(
        description="Find Polish businesses by PKD (SIC) code via GUS API."
    )
    parser.add_argument("--preset", choices=list(PRESETS.keys()),
                        help="Use a preset: 'mobile' for phones, 'diy' for hardware")
    parser.add_argument("--pkd", nargs="+",
                        help="Specific PKD codes to search for (e.g. 47.42.Z 46.52.Z)")
    parser.add_argument("--output", "-o", default="pkd_results.xlsx",
                        help="Output Excel file")
    parser.add_argument("--gus-key", help="GUS API key (or set GUS_API_KEY env var)")
    parser.add_argument("--sandbox", action="store_true", help="Use GUS sandbox")
    parser.add_argument("--active-only", action="store_true",
                        help="Only keep active businesses")
    parser.add_argument("--local-units", action="store_true",
                        help="Fetch local units (branch locations) for each result")
    args = parser.parse_args()

    if not args.preset and not args.pkd:
        logger.error("Specify either --preset (mobile, diy) or --pkd codes")
        parser.print_help()
        sys.exit(1)

    # Determine target PKD codes and keywords
    if args.preset:
        preset = PRESETS[args.preset]
        target_codes = preset["pkd_codes"]
        keywords = preset["keywords"]
        logger.info(f"Preset: {preset['name']}")
    else:
        target_codes = args.pkd
        keywords = []
        for code in target_codes:
            desc = get_pkd_description(code)
            if desc and desc != "Unknown PKD code":
                keywords.append(desc)
        if not keywords:
            keywords = target_codes

    logger.info(f"Target PKD codes: {target_codes}")
    logger.info(f"Search keywords: {len(keywords)}")

    # Set up API
    api_client = PolandAPIClient(gus_api_key=args.gus_key, use_sandbox=args.sandbox)
    if not api_client.gus_api_key:
        logger.error("This script requires a GUS API key to search by name.")
        logger.error("Set GUS_API_KEY in .env file or use --gus-key")
        logger.error("")
        logger.error("TIP: If you don't have a GUS key, use find_mobile_shops.py --krs-only")
        logger.error("     to find known chains via the free KRS API.")
        sys.exit(1)

    # Search
    all_rows = []
    local_unit_rows = []
    seen_regons = set()
    stats = {"searched": 0, "total_hits": 0, "pkd_matches": 0,
             "unique": 0, "active": 0, "krs_enriched": 0, "local_units": 0}

    for idx, keyword in enumerate(keywords):
        logger.info(f"[{idx+1}/{len(keywords)}] Searching: '{keyword}'")
        stats["searched"] += 1

        try:
            results = api_client.search_gus_by_name(keyword)
        except Exception as e:
            logger.error(f"  GUS error: {e}")
            continue

        if not results:
            logger.info(f"  -> 0 results")
            continue

        stats["total_hits"] += len(results)

        new_count = 0
        for r in results:
            regon = r.get("regon", "")
            nip = r.get("nip", "")
            dedup_key = regon or nip
            if not dedup_key or dedup_key in seen_regons:
                continue

            matched_pkd = has_matching_pkd(r, target_codes)
            if not matched_pkd:
                continue

            stats["pkd_matches"] += 1

            if args.active_only and not r.get("is_active"):
                continue

            seen_regons.add(dedup_key)
            if regon:
                seen_regons.add(regon)
            if nip:
                seen_regons.add(nip)

            # KRS enrichment (free)
            if r.get("krs"):
                try:
                    krs_data = api_client.search_krs(r["krs"])
                    if krs_data:
                        r = api_client._merge_results(r, krs_data)
                        stats["krs_enriched"] += 1
                except Exception:
                    pass

            if r.get("is_active"):
                stats["active"] += 1

            target_pkd_str = "; ".join(target_codes)
            all_rows.append(build_row(target_pkd_str, matched_pkd, keyword, r))
            stats["unique"] += 1
            new_count += 1

            # Fetch local units
            if args.local_units and regon and r.get("entity_type"):
                units = api_client.fetch_local_units(regon, r["entity_type"])
                for u in units:
                    local_unit_rows.append({
                        "parent_name": r.get("name", ""),
                        "parent_nip": nip,
                        "parent_regon": regon,
                        **u,
                    })
                stats["local_units"] += len(units)

        logger.info(f"  -> {len(results)} hits, {new_count} new with matching PKD")

    # Save to Excel
    df_out = pd.DataFrame(all_rows, columns=OUTPUT_COLUMNS).fillna("")

    logger.info(f"")
    logger.info(f"Saving {len(df_out)} rows to {args.output}")
    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False, sheet_name="All Results")

        # Active only
        active = df_out[df_out["is_active"].astype(str).str.lower().isin(["true", "1", "yes"])]
        if not active.empty:
            active.to_excel(writer, index=False, sheet_name="Active Only")

        # By PKD code
        pkd_summary = []
        for code in target_codes:
            matching = df_out[df_out["matched_pkd"].str.contains(code.rstrip(".Z"), na=False)]
            main_match = df_out[df_out["pkd_main"] == code]
            pkd_summary.append({
                "pkd_code": code,
                "description": get_pkd_description(code),
                "total_with_code": len(matching),
                "as_main_activity": len(main_match),
            })
        if pkd_summary:
            pd.DataFrame(pkd_summary).to_excel(writer, index=False, sheet_name="PKD Summary")

        # By province
        prov_summary = []
        for prov in sorted(df_out["province"].unique()):
            if not prov:
                continue
            prov_rows = df_out[df_out["province"] == prov]
            prov_active = prov_rows[prov_rows["is_active"].astype(str).str.lower().isin(["true", "1", "yes"])]
            prov_summary.append({"province": prov, "total": len(prov_rows), "active": len(prov_active)})
        if prov_summary:
            pd.DataFrame(prov_summary).to_excel(writer, index=False, sheet_name="By Province")

        # Top cities
        city_counts = df_out["city"].value_counts().head(50)
        if not city_counts.empty:
            city_counts.reset_index().rename(columns={"index": "city", "city": "count"}).to_excel(
                writer, index=False, sheet_name="Top Cities")

        # Has local units
        has_units = df_out[df_out["num_local_units"].astype(str).str.match(r'^[1-9]')]
        if not has_units.empty:
            has_units.to_excel(writer, index=False, sheet_name="Has Local Units")

        # Local unit details
        if local_unit_rows:
            df_units = pd.DataFrame(local_unit_rows).fillna("")
            df_units.to_excel(writer, index=False, sheet_name="Local Units")
            logger.info(f"  Sheet 'Local Units': {len(df_units)} branches")

        # With contact info
        has_contact = df_out[
            (df_out["phone"].astype(str).str.len() > 5) |
            (df_out["email"].astype(str).str.len() > 5) |
            (df_out["website"].astype(str).str.len() > 5)
        ]
        if not has_contact.empty:
            has_contact.to_excel(writer, index=False, sheet_name="With Contact")

    logger.info("=" * 55)
    logger.info(f"DONE  |  Target PKD codes: {', '.join(target_codes)}")
    logger.info(f"  Keywords searched:  {stats['searched']}")
    logger.info(f"  Total GUS hits:     {stats['total_hits']}")
    logger.info(f"  PKD matches:        {stats['pkd_matches']}")
    logger.info(f"  Unique businesses:  {stats['unique']}")
    logger.info(f"  Active:             {stats['active']}")
    logger.info(f"  KRS enriched:       {stats['krs_enriched']}")
    if args.local_units:
        logger.info(f"  Local units:        {stats['local_units']}")
    logger.info(f"Output: {args.output}")


if __name__ == "__main__":
    main()
