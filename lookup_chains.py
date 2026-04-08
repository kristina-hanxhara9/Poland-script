#!/usr/bin/env python3
"""
Look up major Polish retail chain stores (DIY & Mobile Phone) via GUS API.

Searches for known chain retailers by name, retrieves full business details
including NIP, REGON, KRS, addresses, PKD codes, etc.

Usage:
    python lookup_chains.py --output chains.xlsx
    python lookup_chains.py --category diy --output diy_chains.xlsx
    python lookup_chains.py --category mobile --output mobile_chains.xlsx
    python lookup_chains.py --category all --output all_chains.xlsx
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

# ─── Known Polish retail chains ─────────────────────────────────────────────

DIY_CHAINS = [
    # Major international chains in Poland
    {"name": "Castorama Polska", "search_names": ["Castorama", "CASTORAMA POLSKA"], "category": "DIY"},
    {"name": "Leroy Merlin Polska", "search_names": ["Leroy Merlin", "LEROY MERLIN POLSKA"], "category": "DIY"},
    {"name": "OBI Polska", "search_names": ["OBI", "OBI POLSKA"], "category": "DIY"},
    {"name": "Bricomarché", "search_names": ["Bricomarche", "BRICOMARCHE"], "category": "DIY"},
    {"name": "Bricoman Polska", "search_names": ["Bricoman", "BRICOMAN POLSKA"], "category": "DIY"},

    # Polish DIY chains
    {"name": "PSB Mrówka", "search_names": ["PSB Mrowka", "PSB HANDEL", "GRUPA PSB"], "category": "DIY"},
    {"name": "Nomi", "search_names": ["NOMI"], "category": "DIY"},
    {"name": "Majster Plus", "search_names": ["Majster Plus", "MAJSTER"], "category": "DIY"},
    {"name": "Polbruk", "search_names": ["POLBRUK"], "category": "DIY"},
    {"name": "Praktiker Polska", "search_names": ["Praktiker", "PRAKTIKER POLSKA"], "category": "DIY"},

    # Building materials / hardware
    {"name": "Merkury Market", "search_names": ["Merkury Market", "MERKURY MARKET"], "category": "DIY"},
    {"name": "Abra Meble", "search_names": ["ABRA", "ABRA MEBLE"], "category": "DIY"},
    {"name": "Jula Poland", "search_names": ["JULA", "JULA POLAND"], "category": "DIY"},
    {"name": "Würth Polska", "search_names": ["Wurth Polska", "WURTH"], "category": "DIY"},
    {"name": "Brico Depot", "search_names": ["BRICO DEPOT", "BRICODEPOT"], "category": "DIY"},

    # Specialist building/home
    {"name": "Leroymerlin", "search_names": ["LEROYMERLIN"], "category": "DIY"},
    {"name": "Komfort", "search_names": ["KOMFORT"], "category": "DIY"},
    {"name": "Atlas", "search_names": ["ATLAS FIRMA"], "category": "DIY"},
    {"name": "MAT-BUD", "search_names": ["MAT-BUD", "MATBUD"], "category": "DIY"},
    {"name": "Brico Market", "search_names": ["BRICO MARKET"], "category": "DIY"},
]

MOBILE_CHAINS = [
    # Major operators / retailers
    {"name": "Orange Polska", "search_names": ["Orange Polska", "ORANGE POLSKA"], "category": "Mobile"},
    {"name": "T-Mobile Polska", "search_names": ["T-Mobile Polska", "T-MOBILE POLSKA"], "category": "Mobile"},
    {"name": "Play (P4)", "search_names": ["P4 Sp. z o.o.", "P4"], "category": "Mobile"},
    {"name": "Plus (Polkomtel)", "search_names": ["Polkomtel", "POLKOMTEL"], "category": "Mobile"},
    {"name": "Vectra", "search_names": ["VECTRA"], "category": "Mobile"},

    # Electronics / phone retailers
    {"name": "Media Expert", "search_names": ["Media Expert", "MEDIA EXPERT", "TERG"], "category": "Mobile"},
    {"name": "Media Markt", "search_names": ["Media Markt", "MEDIA SATURN", "MEDIAMARKT"], "category": "Mobile"},
    {"name": "RTV Euro AGD", "search_names": ["Euro AGD", "RTV EURO AGD", "EURO-NET"], "category": "Mobile"},
    {"name": "x-kom", "search_names": ["x-kom", "X-KOM"], "category": "Mobile"},
    {"name": "Komputronik", "search_names": ["Komputronik", "KOMPUTRONIK"], "category": "Mobile"},
    {"name": "Morele.net", "search_names": ["Morele.net", "MORELE NET", "MORELE"], "category": "Mobile"},

    # Mobile-specific retailers
    {"name": "iSpot (Apple)", "search_names": ["iSpot", "ISPOT", "APPLE POLSKA"], "category": "Mobile"},
    {"name": "Samsung Polska", "search_names": ["Samsung Electronics Polska", "SAMSUNG POLSKA", "SAMSUNG ELECTRONICS"], "category": "Mobile"},
    {"name": "Xiaomi Polska", "search_names": ["Xiaomi", "XIAOMI"], "category": "Mobile"},
    {"name": "Huawei Polska", "search_names": ["Huawei Polska", "HUAWEI"], "category": "Mobile"},
    {"name": "Mi-Home", "search_names": ["MI-HOME", "MI HOME"], "category": "Mobile"},

    # MVNO / virtual operators
    {"name": "Virgin Mobile Polska", "search_names": ["Virgin Mobile", "VIRGIN MOBILE"], "category": "Mobile"},
    {"name": "Nju Mobile (Orange)", "search_names": ["NJU MOBILE"], "category": "Mobile"},
    {"name": "Lajt Mobile", "search_names": ["LAJT MOBILE", "LAJT"], "category": "Mobile"},
    {"name": "Lycamobile", "search_names": ["Lycamobile", "LYCAMOBILE"], "category": "Mobile"},

    # Phone repair / accessories chains
    {"name": "GSM Service", "search_names": ["GSM SERVICE"], "category": "Mobile"},
    {"name": "TelForceOne", "search_names": ["TelForceOne", "TELFORCEONE"], "category": "Mobile"},
    {"name": "mBank (mobile banking)", "search_names": ["MBANK"], "category": "Mobile"},
]

OUTPUT_COLUMNS = [
    "chain_name",
    "category",
    "search_method",
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
    "full_address",
    "pkd_main",
    "pkd_codes",
    "pkd_descriptions",
    "date_started",
    "date_ended",
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


def simplify_name(name):
    """Remove legal form suffixes."""
    lower = name.lower().strip()
    for suffix in [
        " sp. z o.o. sp. k.", " sp. z o.o. sp.k.",
        " sp. z o.o.", " sp.z o.o.", " s.a.", " sp.j.", " sp.k.",
        " z o.o.", " z.o.o.",
    ]:
        if lower.endswith(suffix):
            return name[:len(name) - len(suffix)].strip()
    return name.strip()


def search_chain(api_client, chain_info):
    """Search for a chain retailer using multiple name variants."""
    results = []
    seen_regons = set()

    for search_name in chain_info["search_names"]:
        logger.info(f"  Trying: '{search_name}'")
        gus_results = api_client.search_gus_by_name(search_name)
        if gus_results:
            for r in gus_results:
                regon = r.get("regon", "")
                if regon and regon not in seen_regons:
                    seen_regons.add(regon)
                    r["_search_term"] = search_name
                    results.append(r)
            logger.info(f"    Found {len(gus_results)} (new: {len(results)})")

        # Also try simplified
        simplified = simplify_name(search_name)
        if simplified.lower() != search_name.lower():
            gus_results = api_client.search_gus_by_name(simplified)
            if gus_results:
                for r in gus_results:
                    regon = r.get("regon", "")
                    if regon and regon not in seen_regons:
                        seen_regons.add(regon)
                        r["_search_term"] = simplified
                        results.append(r)

        if results:
            break  # Found with this variant, no need to try more

    return results


def build_row(chain_info, result, search_term):
    pkd_codes = result.get("pkd_codes", [])
    api_descs = result.get("pkd_descriptions", {})
    pkd_descs = [api_descs.get(c, "") or get_pkd_description(c) for c in pkd_codes]

    return {
        "chain_name": chain_info["name"],
        "category": chain_info["category"],
        "search_method": f"GUS ({search_term})",
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
        "full_address": result.get("address_str", ""),
        "pkd_main": result.get("pkd_main", ""),
        "pkd_codes": "; ".join(pkd_codes),
        "pkd_descriptions": "; ".join(pkd_descs),
        "date_started": result.get("date_started", ""),
        "date_ended": result.get("date_ended", ""),
        "email": result.get("email", ""),
        "phone": result.get("phone", ""),
        "website": result.get("website", ""),
        "api_source": result.get("source", ""),
    }


def main():
    parser = argparse.ArgumentParser(
        description="Look up Polish DIY and mobile phone retail chains via GUS API."
    )
    parser.add_argument("--output", "-o", default="chains.xlsx", help="Output Excel file")
    parser.add_argument("--category", "-c", default="all",
                        choices=["all", "diy", "mobile"],
                        help="Which category to search (default: all)")
    parser.add_argument("--gus-key", help="GUS API key (or set GUS_API_KEY env var)")
    parser.add_argument("--sandbox", action="store_true", help="Use GUS sandbox")
    args = parser.parse_args()

    # Build chain list
    chains = []
    if args.category in ("all", "diy"):
        chains.extend(DIY_CHAINS)
    if args.category in ("all", "mobile"):
        chains.extend(MOBILE_CHAINS)

    logger.info(f"Searching {len(chains)} retail chains ({args.category})")

    # Set up API
    api_client = PolandAPIClient(gus_api_key=args.gus_key, use_sandbox=args.sandbox)
    if not api_client.gus_api_key:
        logger.error("No GUS API key. Set GUS_API_KEY env var or use --gus-key.")
        sys.exit(1)

    # Process
    all_rows = []
    stats = {"found": 0, "not_found": 0, "total_entities": 0}

    for idx, chain in enumerate(chains):
        logger.info(f"[{idx + 1}/{len(chains)}] {chain['name']} ({chain['category']})")

        results = search_chain(api_client, chain)

        if results:
            stats["found"] += 1
            stats["total_entities"] += len(results)
            for r in results:
                search_term = r.pop("_search_term", chain["search_names"][0])
                # Enrich with KRS
                if r.get("krs"):
                    krs_data = api_client.search_krs(r["krs"])
                    if krs_data:
                        r = api_client._merge_results(r, krs_data)
                all_rows.append(build_row(chain, r, search_term))
            logger.info(f"  -> Found {len(results)} entities")
        else:
            stats["not_found"] += 1
            all_rows.append({
                "chain_name": chain["name"],
                "category": chain["category"],
                "search_method": "NOT FOUND",
            })
            logger.info(f"  -> Not found")

    # Output
    df_all = pd.DataFrame(all_rows, columns=OUTPUT_COLUMNS).fillna("")

    logger.info(f"Saving to {args.output}")
    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        # Sheet 1: All results
        df_all.to_excel(writer, index=False, sheet_name="All Chains")

        # Sheet 2: DIY only
        df_diy = df_all[df_all["category"] == "DIY"]
        if not df_diy.empty:
            df_diy.to_excel(writer, index=False, sheet_name="DIY Chains")

        # Sheet 3: Mobile only
        df_mobile = df_all[df_all["category"] == "Mobile"]
        if not df_mobile.empty:
            df_mobile.to_excel(writer, index=False, sheet_name="Mobile Chains")

        # Sheet 4: Summary
        summary = []
        for cat in ["DIY", "Mobile"]:
            cat_rows = [r for r in all_rows if r.get("category") == cat
                        and r.get("search_method", "") != "NOT FOUND"]
            not_found = [r for r in all_rows if r.get("category") == cat
                         and r.get("search_method", "") == "NOT FOUND"]
            summary.append({
                "category": cat,
                "chains_found": len(set(r["chain_name"] for r in cat_rows)),
                "chains_not_found": len(not_found),
                "total_entities": len(cat_rows),
                "active": len([r for r in cat_rows if r.get("is_active") == True]),
                "not_found_names": ", ".join(r["chain_name"] for r in not_found),
            })
        pd.DataFrame(summary).to_excel(writer, index=False, sheet_name="Summary")

    logger.info("=" * 55)
    logger.info(f"DONE  |  Chains searched: {len(chains)}")
    logger.info(f"  Found:           {stats['found']}")
    logger.info(f"  Not found:       {stats['not_found']}")
    logger.info(f"  Total entities:  {stats['total_entities']}")
    logger.info(f"Results: {args.output}")


if __name__ == "__main__":
    main()
