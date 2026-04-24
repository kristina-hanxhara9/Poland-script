#!/usr/bin/env python3
"""
Find mobile phone specialists and accessories shops in Poland.

Combines multiple search methods:
  1. GUS API — search by business name keywords (needs GUS key)
  2. CEIDG API — search by name keywords, filter by mobile PKD codes (needs CEIDG key)
  3. KRS API — look up known mobile chains by KRS number (FREE, no key needed!)
  4. KRS enrichment — add full company data from KRS

Use --krs-only to find chains WITHOUT any API keys (uses the built-in list of
known mobile phone chains with their KRS numbers).

Relevant PKD codes for mobile phone shops:
  47.42.Z — Retail sale of telecommunications equipment
  46.52.Z — Wholesale of electronic & telecom equipment
  95.12.Z — Repair of communication equipment
  47.41.Z — Retail sale of computers, software (overlap)
  61.10.Z — Wired telecom activities
  61.20.Z — Wireless telecom activities

Outputs ALL available API fields.

Usage:
    python find_mobile_shops.py --krs-only --output mobile_chains.xlsx
    python find_mobile_shops.py --output mobile_specialists.xlsx
    python find_mobile_shops.py --gus-key YOUR_KEY --output mobile_specialists.xlsx
    python find_mobile_shops.py --active-only --output mobile_active.xlsx
"""

import argparse
import logging
import sys

from dotenv import load_dotenv
import pandas as pd

load_dotenv()

from poland_api import PolandAPIClient
from lookup_chains import MOBILE_CHAINS

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

# ─── PKD codes for mobile phone / telecom retail ───────────────────────────

MOBILE_PKD_CODES = [
    "47.42.Z",  # Retail sale of telecommunications equipment in specialised stores
    "46.52.Z",  # Wholesale of electronic and telecommunications equipment
    "95.12.Z",  # Repair of communication equipment
    "47.41.Z",  # Retail sale of computers, peripheral units and software
    "61.10.Z",  # Wired telecommunications activities
    "61.20.Z",  # Wireless telecommunications activities
    "46.43.Z",  # Wholesale of electrical household appliances (includes phones)
    "47.43.Z",  # Retail sale of audio and video equipment (overlap)
    "77.22.Z",  # Renting of video tapes, CDs, phones
]

# ─── Search keywords for GUS name search ───────────────────────────────────

GUS_KEYWORDS = [
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
    "akcesoria komórkowe",
    "etui telefon",
    "serwis GSM",
    "serwis telefonów",
    "naprawa telefonów",
    "serwis komórkowy",
    "hurtownia GSM",
    "hurtownia telefonów",
    "dystrybutor GSM",
    "telekomunikacja",
    "telekomunikacyjny",
    "Teletorium",
    "Teleakces",
    "iSpot",
    "Mi Store",
    "Xiaomi",
    "myPhone",
    "MPTECH",
    "Maxcom",
    "TelForceOne",
    "Cortland",
    "Komputronik",
    "x-kom",
    "Media Expert",
    "TERG",
    "Euro AGD",
    "EURO-NET",
    "Samsung Electronics Polska",
    "Huawei Polska",
    "Orange Polska",
    "T-Mobile Polska",
    "P4",
    "Polkomtel",
]

# ─── CEIDG keywords (paired with PKD filtering) ───────────────────────────

CEIDG_KEYWORDS = [
    "telefon",
    "GSM",
    "mobile",
    "smartfon",
    "akcesoria",
    "serwis telefon",
    "komórkowy",
    "telekomunikac",
]

OUTPUT_COLUMNS = [
    "search_method",
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
    "ceidg_link",
    "api_source",
]


def get_pkd_description(code):
    try:
        from pkd_codes import get_pkd_description as _get_desc
        return _get_desc(code)
    except ImportError:
        return ""


def is_mobile_related(result):
    """Check if a result's PKD codes or name indicate mobile/telecom business."""
    pkd_main = result.get("pkd_main", "")
    pkd_codes = result.get("pkd_codes", [])
    all_codes = set(pkd_codes)
    if pkd_main:
        all_codes.add(pkd_main)

    mobile_prefixes = ["47.42", "46.52", "95.12", "47.41", "61.10", "61.20", "46.43", "47.43", "77.22"]
    for code in all_codes:
        for prefix in mobile_prefixes:
            if code.startswith(prefix):
                return True

    name = (result.get("name", "") or "").upper()
    hints = ["TELEFON", "GSM", "MOBILE", "SMARTFON", "KOMÓRK", "KOMORK",
             "AKCESORI", "SERWIS", "TELEKOM", "SAMSUNG", "HUAWEI", "XIAOMI",
             "IPHONE", "APPLE", "NOKIA"]
    for hint in hints:
        if hint in name:
            return True

    return False


def build_row(method, keyword, result):
    pkd_codes = result.get("pkd_codes", [])
    api_descs = result.get("pkd_descriptions", {})
    if isinstance(api_descs, dict):
        pkd_descs = [api_descs.get(c, "") or get_pkd_description(c) for c in pkd_codes]
    else:
        pkd_descs = []

    return {
        "search_method": method,
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
        "ceidg_link": result.get("ceidg_link", ""),
        "api_source": result.get("source", ""),
    }


def add_result(result, method, keyword, seen_ids, all_rows, api_client, stats, active_only):
    """Deduplicate by REGON/NIP, enrich with KRS, add to rows."""
    regon = result.get("regon", "")
    nip = result.get("nip", "")
    dedup_key = regon or nip
    if not dedup_key:
        return False

    if dedup_key in seen_ids:
        stats["duplicates"] += 1
        return False

    seen_ids.add(dedup_key)
    if regon:
        seen_ids.add(regon)
    if nip:
        seen_ids.add(nip)

    if active_only and not result.get("is_active"):
        return False

    # KRS enrichment
    if result.get("krs"):
        try:
            krs_data = api_client.search_krs(result["krs"])
            if krs_data:
                result = api_client._merge_results(result, krs_data)
                stats["krs_enriched"] += 1
        except Exception as e:
            logger.debug(f"  KRS failed: {e}")

    if result.get("is_active"):
        stats["active"] += 1

    all_rows.append(build_row(method, keyword, result))
    stats["unique"] += 1
    return True


def main():
    parser = argparse.ArgumentParser(
        description="Find mobile phone specialists and accessories shops via GUS + CEIDG + KRS."
    )
    parser.add_argument("--output", "-o", default="mobile_specialists.xlsx",
                        help="Output Excel file (default: mobile_specialists.xlsx)")
    parser.add_argument("--gus-key", help="GUS API key (or set GUS_API_KEY env var)")
    parser.add_argument("--ceidg-key", help="CEIDG API key (or set CEIDG_API_KEY env var)")
    parser.add_argument("--sandbox", action="store_true", help="Use GUS sandbox")
    parser.add_argument("--active-only", action="store_true",
                        help="Only keep active businesses")
    parser.add_argument("--skip-gus", action="store_true", help="Skip GUS name search")
    parser.add_argument("--skip-ceidg", action="store_true", help="Skip CEIDG search")
    parser.add_argument("--krs-only", action="store_true",
                        help="Only use the FREE KRS API with known mobile chain KRS numbers. "
                        "No GUS or CEIDG key needed!")
    parser.add_argument("--local-units", action="store_true",
                        help="Also fetch local units (branch locations) for chains. "
                        "Requires GUS API key.")
    args = parser.parse_args()

    api_client = PolandAPIClient(
        gus_api_key=args.gus_key,
        ceidg_api_key=args.ceidg_key,
        use_sandbox=args.sandbox,
    )

    has_gus = bool(api_client.gus_api_key) and not args.krs_only
    has_ceidg = bool(api_client.ceidg_api_key) and not args.krs_only

    if not has_gus and not has_ceidg and not args.krs_only:
        logger.warning("No GUS or CEIDG API key found. Switching to KRS-only mode.")
        logger.info("KRS API is FREE — will look up known mobile chains by KRS number.")
        args.krs_only = True

    logger.info(f"APIs: GUS={'YES' if has_gus else 'NO'}, CEIDG={'YES' if has_ceidg else 'NO'}, KRS=YES (free)")

    all_rows = []
    local_unit_rows = []
    seen_ids = set()
    stats = {"total_found": 0, "unique": 0, "duplicates": 0,
             "krs_enriched": 0, "active": 0}

    # ─── METHOD 1: GUS search by business name keywords ────────────────
    if has_gus and not args.skip_gus:
        logger.info(f"")
        logger.info(f"{'='*55}")
        logger.info(f"METHOD 1: GUS name search ({len(GUS_KEYWORDS)} keywords)")
        logger.info(f"{'='*55}")
        for idx, keyword in enumerate(GUS_KEYWORDS):
            logger.info(f"[GUS {idx+1}/{len(GUS_KEYWORDS)}] '{keyword}'")
            try:
                results = api_client.search_gus_by_name(keyword)
            except Exception as e:
                logger.error(f"  GUS error: {e}")
                continue

            if not results:
                logger.info(f"  -> 0 results")
                continue

            new_count = 0
            for r in results:
                stats["total_found"] += 1
                if add_result(r, "GUS_NAME", keyword, seen_ids, all_rows,
                              api_client, stats, args.active_only):
                    new_count += 1
            logger.info(f"  -> {len(results)} results, {new_count} new unique")

    # ─── METHOD 2: CEIDG search by name, filter by mobile PKD ──────────
    if has_ceidg and not args.skip_ceidg:
        logger.info(f"")
        logger.info(f"{'='*55}")
        logger.info(f"METHOD 2: CEIDG name search ({len(CEIDG_KEYWORDS)} keywords)")
        logger.info(f"{'='*55}")
        for idx, keyword in enumerate(CEIDG_KEYWORDS):
            logger.info(f"[CEIDG {idx+1}/{len(CEIDG_KEYWORDS)}] '{keyword}'")
            try:
                results = api_client.search_ceidg(keyword)
            except Exception as e:
                logger.error(f"  CEIDG error: {e}")
                continue

            if not results:
                logger.info(f"  -> 0 results")
                continue

            mobile_results = [r for r in results if is_mobile_related(r)]
            logger.info(f"  -> {len(results)} total, {len(mobile_results)} mobile-related")

            new_count = 0
            for r in mobile_results:
                stats["total_found"] += 1
                if add_result(r, "CEIDG_NAME", keyword, seen_ids, all_rows,
                              api_client, stats, args.active_only):
                    new_count += 1
            if new_count:
                logger.info(f"  -> {new_count} new unique added")

    # ─── METHOD 3: GUS search with PKD activity descriptions ───────────
    if has_gus:
        logger.info(f"")
        logger.info(f"{'='*55}")
        logger.info(f"METHOD 3: GUS PKD-targeted name search")
        logger.info(f"{'='*55}")
        pkd_descriptions = [
            "sprzedaż detaliczna sprzętu telekomunikacyjnego",
            "sprzedaż hurtowa sprzętu telekomunikacyjnego",
            "naprawa sprzętu komunikacyjnego",
            "sprzedaż detaliczna komputerów",
            "działalność w zakresie telekomunikacji",
            "telekomunikacja bezprzewodowa",
        ]
        for idx, desc in enumerate(pkd_descriptions):
            logger.info(f"[PKD {idx+1}/{len(pkd_descriptions)}] '{desc}'")
            try:
                results = api_client.search_gus_by_name(desc)
            except Exception as e:
                logger.error(f"  GUS error: {e}")
                continue

            if not results:
                logger.info(f"  -> 0 results")
                continue

            matched = [r for r in results if is_mobile_related(r)]
            logger.info(f"  -> {len(results)} total, {len(matched)} with mobile PKD/name")

            new_count = 0
            for r in matched:
                stats["total_found"] += 1
                if add_result(r, "GUS_PKD", desc, seen_ids, all_rows,
                              api_client, stats, args.active_only):
                    new_count += 1
            if new_count:
                logger.info(f"  -> {new_count} new unique added")

    # ─── METHOD 4: CEIDG direct PKD code search ───────────────────────
    if has_ceidg and not args.skip_ceidg:
        logger.info(f"")
        logger.info(f"{'='*55}")
        logger.info(f"METHOD 4: CEIDG direct PKD code search ({len(MOBILE_PKD_CODES)} codes)")
        logger.info(f"{'='*55}")
        for idx, pkd_code in enumerate(MOBILE_PKD_CODES):
            logger.info(f"[PKD {idx+1}/{len(MOBILE_PKD_CODES)}] Searching CEIDG for PKD: {pkd_code}")
            try:
                results = api_client.search_ceidg_by_pkd(pkd_code, status=1)
            except Exception as e:
                logger.error(f"  CEIDG PKD error: {e}")
                continue

            if not results:
                logger.info(f"  -> 0 results")
                continue

            new_count = 0
            for r in results:
                stats["total_found"] += 1
                if add_result(r, "CEIDG_PKD", pkd_code, seen_ids, all_rows,
                              api_client, stats, args.active_only):
                    new_count += 1
            logger.info(f"  -> {len(results)} results, {new_count} new unique")

    # ─── METHOD 5: KRS lookup of known mobile chains ────────────────────
    # This works WITHOUT any API key — KRS is free
    krs_chains = [c for c in MOBILE_CHAINS if c.get("krs") and c["krs"].isdigit()]
    if krs_chains:
        logger.info(f"")
        logger.info(f"{'='*55}")
        logger.info(f"METHOD 5: KRS lookup of known mobile chains ({len(krs_chains)} with KRS numbers)")
        logger.info(f"{'='*55}")

        for idx, chain in enumerate(krs_chains):
            krs_num = chain["krs"]
            logger.info(f"[KRS {idx+1}/{len(krs_chains)}] {chain['name']} — KRS: {krs_num}")
            try:
                result = api_client.search_krs(krs_num)
            except Exception as e:
                logger.error(f"  KRS error: {e}")
                continue

            if not result:
                logger.info(f"  -> Not found in KRS")
                continue

            stats["total_found"] += 1

            # If we have GUS, enrich with full GUS data
            if has_gus and result.get("nip"):
                try:
                    gus_data = api_client.search_gus_by_nip(result["nip"])
                    if gus_data:
                        result = api_client._merge_results(gus_data, result)
                        logger.info(f"  GUS enrichment OK")
                except Exception as e:
                    logger.debug(f"  GUS enrichment failed: {e}")

            added = add_result(result, "KRS_CHAIN", chain["name"], seen_ids, all_rows,
                               api_client, stats, args.active_only)
            if added:
                logger.info(f"  -> {result.get('name', '?')} | {result.get('city', '')} | "
                             f"PKD: {result.get('pkd_main', '')}")

                # Fetch local units if requested and we have GUS
                if args.local_units and has_gus and result.get("regon") and result.get("entity_type"):
                    units = api_client.fetch_local_units(result["regon"], result["entity_type"])
                    for u in units:
                        local_unit_rows.append({
                            "chain_name": chain["name"],
                            "parent_nip": result.get("nip", ""),
                            "parent_regon": result.get("regon", ""),
                            "parent_name": result.get("name", ""),
                            **u,
                        })
                    if units:
                        logger.info(f"  -> {len(units)} local units (branches)")
            else:
                logger.info(f"  -> Already found via other method")

    # ─── Save to Excel ─────────────────────────────────────────────────
    df_out = pd.DataFrame(all_rows, columns=OUTPUT_COLUMNS).fillna("")

    logger.info(f"")
    logger.info(f"Saving {len(df_out)} rows to {args.output}")
    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False, sheet_name="All Mobile Shops")

        active = df_out[df_out["is_active"].astype(str).str.lower().isin(["true", "1", "yes"])]
        if not active.empty:
            active.to_excel(writer, index=False, sheet_name="Active Only")

        # Local units (branch locations) sheet
        if args.local_units and local_unit_rows:
            df_units = pd.DataFrame(local_unit_rows).fillna("")
            df_units.to_excel(writer, index=False, sheet_name="Local Units")
            logger.info(f"  Sheet 'Local Units': {len(df_units)} branch locations")

        pkd_summary = []
        for pkd in df_out["pkd_main"].unique():
            if not pkd:
                continue
            pkd_rows = df_out[df_out["pkd_main"] == pkd]
            pkd_active = pkd_rows[pkd_rows["is_active"].astype(str).str.lower().isin(["true", "1", "yes"])]
            pkd_summary.append({
                "pkd_code": pkd,
                "description": get_pkd_description(pkd),
                "total": len(pkd_rows),
                "active": len(pkd_active),
            })
        if pkd_summary:
            pd.DataFrame(pkd_summary).sort_values("total", ascending=False).to_excel(
                writer, index=False, sheet_name="PKD Summary")

        prov_summary = []
        for prov in sorted(df_out["province"].unique()):
            if not prov:
                continue
            prov_rows = df_out[df_out["province"] == prov]
            prov_active = prov_rows[prov_rows["is_active"].astype(str).str.lower().isin(["true", "1", "yes"])]
            prov_summary.append({"province": prov, "total": len(prov_rows), "active": len(prov_active)})
        if prov_summary:
            pd.DataFrame(prov_summary).to_excel(writer, index=False, sheet_name="By Province")

        city_summary = []
        for city in df_out["city"].unique():
            if not city:
                continue
            city_summary.append({"city": city, "total": len(df_out[df_out["city"] == city])})
        if city_summary:
            pd.DataFrame(city_summary).sort_values("total", ascending=False).head(50).to_excel(
                writer, index=False, sheet_name="Top Cities")

        method_summary = []
        for method in df_out["search_method"].unique():
            if not method:
                continue
            method_summary.append({"method": method, "count": len(df_out[df_out["search_method"] == method])})
        if method_summary:
            pd.DataFrame(method_summary).to_excel(writer, index=False, sheet_name="By Method")

        suspended = df_out[df_out["date_suspended"].astype(str).str.len() > 0]
        if not suspended.empty:
            suspended.to_excel(writer, index=False, sheet_name="Suspended")

        has_units = df_out[df_out["num_local_units"].astype(str).str.match(r'^[1-9]')]
        if not has_units.empty:
            has_units.to_excel(writer, index=False, sheet_name="Has Local Units")

        has_contact = df_out[
            (df_out["phone"].astype(str).str.len() > 5) |
            (df_out["email"].astype(str).str.len() > 5) |
            (df_out["website"].astype(str).str.len() > 5)
        ]
        if not has_contact.empty:
            has_contact.to_excel(writer, index=False, sheet_name="With Contact")

    logger.info("=" * 55)
    logger.info(f"DONE  |  Total API hits: {stats['total_found']}")
    logger.info(f"  Unique businesses:  {stats['unique']}")
    logger.info(f"  Duplicates skipped: {stats['duplicates']}")
    logger.info(f"  KRS enriched:       {stats['krs_enriched']}")
    logger.info(f"  Active:             {stats['active']}")
    logger.info(f"Output: {args.output}")


if __name__ == "__main__":
    main()
