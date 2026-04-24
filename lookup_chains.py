#!/usr/bin/env python3
"""
Look up major Polish retail chain stores (DIY & Mobile Phone) via GUS + KRS APIs.

Searches for known chain retailers by name, retrieves full business details
including NIP, REGON, KRS, addresses, PKD codes, local units, etc.

Three modes:
  1. Full mode (default): GUS name search + KRS enrichment (needs GUS API key)
  2. KRS-only mode (--krs-only): Uses FREE KRS API with known KRS numbers (no key needed!)
  3. Local units (--local-units): Also fetches all branch locations (needs GUS key)

Usage:
    python lookup_chains.py --krs-only --category mobile --output mobile_chains.xlsx
    python lookup_chains.py --category mobile --output mobile_chains.xlsx
    python lookup_chains.py --local-units --category mobile --output mobile_with_units.xlsx
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
    # ─── Major international DIY chains in Poland ────────────────────────
    {"name": "Castorama Polska", "search_names": ["Castorama", "CASTORAMA POLSKA"], "category": "DIY"},
    {"name": "Leroy Merlin Polska", "search_names": ["Leroy Merlin", "LEROY MERLIN POLSKA"], "category": "DIY"},
    {"name": "OBI Polska", "search_names": ["OBI", "OBI POLSKA"], "category": "DIY"},
    {"name": "Bricomarché (Grupa Muszkieterów)", "search_names": ["Bricomarche", "BRICOMARCHE", "ITM POLSKA"], "category": "DIY"},
    {"name": "Bricoman Polska", "search_names": ["Bricoman", "BRICOMAN POLSKA"], "category": "DIY"},
    {"name": "Jula Poland", "search_names": ["JULA POLAND", "JULA"], "category": "DIY"},
    {"name": "Würth Polska", "search_names": ["Wurth Polska", "WURTH POLSKA"], "category": "DIY"},
    {"name": "Hilti Polska", "search_names": ["Hilti Polska", "HILTI POLSKA"], "category": "DIY"},
    {"name": "MR.DIY Polska", "search_names": ["MR.DIY", "MRDIY", "MR DIY POLSKA"], "category": "DIY"},
    {"name": "Flügger Polska", "search_names": ["FLUGGER", "FLUGGER POLSKA"], "category": "DIY"},

    # ─── Polish DIY / building materials chains ──────────────────────────
    {"name": "Grupa PSB (Mrówka, Profi)", "search_names": ["GRUPA PSB", "PSB HANDEL", "PSB Mrowka"], "category": "DIY"},
    {"name": "PSB Mrówka", "search_names": ["PSB MROWKA", "MROWKA"], "category": "DIY"},
    {"name": "PSB Profi", "search_names": ["PSB PROFI"], "category": "DIY"},
    {"name": "Nomi", "search_names": ["NOMI"], "category": "DIY"},
    {"name": "Majster Plus (Grupa PMB)", "search_names": ["Majster Plus", "MAJSTER", "GRUPA PMB"], "category": "DIY"},
    {"name": "Praktiker Polska", "search_names": ["Praktiker", "PRAKTIKER POLSKA"], "category": "DIY"},
    {"name": "Merkury Market", "search_names": ["Merkury Market", "MERKURY MARKET"], "category": "DIY"},
    {"name": "Budokop", "search_names": ["BUDOKOP"], "category": "DIY"},
    {"name": "Mat-Bud", "search_names": ["MAT-BUD", "MATBUD"], "category": "DIY"},

    # ─── Purchasing groups / franchise networks ──────────────────────────
    {"name": "Grupa Handlo-Budowa (GHB)", "search_names": ["GRUPA HANDLO-BUDOWA", "GHB"], "category": "DIY"},
    {"name": "Majster Budowlane ABC", "search_names": ["MAJSTER BUDOWLANE ABC"], "category": "DIY"},
    {"name": "Pszczółka (Grupa Majster)", "search_names": ["PSZCZOLKA", "INVEST DEVELOPMENT"], "category": "DIY"},
    {"name": "Sieć Budowlana", "search_names": ["SIEC BUDOWLANA"], "category": "DIY"},
    {"name": "3W Dystrybucja Budowlana", "search_names": ["3W DYSTRYBUCJA"], "category": "DIY"},
    {"name": "Grupa PHMB", "search_names": ["POLSKIE HURTOWNIE MATERIALOW BUDOWLANYCH", "PHMB"], "category": "DIY"},
    {"name": "BIMS Plus (Abex)", "search_names": ["BIMS PLUS", "ABEX"], "category": "DIY"},
    {"name": "Grupa ABG", "search_names": ["GRUPA ABG"], "category": "DIY"},
    {"name": "Hadex (Grupa PHMB)", "search_names": ["HADEX"], "category": "DIY"},

    # ─── Professional distributors ───────────────────────────────────────
    {"name": "Saint-Gobain Dystrybucja (Raab Karcher)", "search_names": ["SAINT-GOBAIN POLSKA", "RAAB KARCHER"], "category": "DIY"},
    {"name": "Onninen Polska (Kesko)", "search_names": ["ONNINEN"], "category": "DIY"},
    {"name": "Enexon (ex-Elektroskandia)", "search_names": ["ENEXON"], "category": "DIY"},

    # ─── Building materials wholesalers / distributors ───────────────────
    {"name": "Grupa Polskie Składy Budowlane", "search_names": ["POLSKIE SKLADY BUDOWLANE"], "category": "DIY"},
    {"name": "AB Bechcicki (Bygma)", "search_names": ["AB BECHCICKI", "BECHCICKI"], "category": "DIY"},
    {"name": "Grupa Dekpol", "search_names": ["DEKPOL"], "category": "DIY"},
    {"name": "Grupa Atlas", "search_names": ["ATLAS", "GRUPA ATLAS"], "category": "DIY"},
    {"name": "Śnieżka (paints)", "search_names": ["SNIEZKA", "FABRYKA FARB SNIEZKA"], "category": "DIY"},
    {"name": "Dekoral", "search_names": ["DEKORAL", "PPG DECO POLSKA"], "category": "DIY"},
    {"name": "Selena Group (Tytan)", "search_names": ["SELENA", "SELENA GROUP"], "category": "DIY"},
    {"name": "SIG Polska (roofing)", "search_names": ["SIG POLSKA", "SIG"], "category": "DIY"},
    {"name": "Blachy Pruszyński", "search_names": ["BLACHY PRUSZYNSKI", "PRUSZYNSKI"], "category": "DIY"},
    {"name": "Budmat", "search_names": ["BUDMAT"], "category": "DIY"},
    {"name": "Fakro (windows/skylights)", "search_names": ["FAKRO"], "category": "DIY"},
    {"name": "Velux Polska", "search_names": ["VELUX POLSKA", "VELUX"], "category": "DIY"},
    {"name": "Wiśniowski (garage doors/gates)", "search_names": ["WISNIOWSKI"], "category": "DIY"},

    # ─── Flooring / tiles / bathrooms ────────────────────────────────────
    {"name": "Komfort", "search_names": ["KOMFORT"], "category": "DIY"},
    {"name": "Cersanit", "search_names": ["CERSANIT"], "category": "DIY"},
    {"name": "Paradyż (tiles)", "search_names": ["CERAMIKA PARADYZ", "PARADYZ"], "category": "DIY"},
    {"name": "Tubądzin (tiles)", "search_names": ["TUBADZIN"], "category": "DIY"},
    {"name": "Opoczno (tiles)", "search_names": ["OPOCZNO", "CERAMIKA OPOCZNO"], "category": "DIY"},
    {"name": "Deante (bathrooms)", "search_names": ["DEANTE"], "category": "DIY"},
    {"name": "Sanitec / Koło", "search_names": ["SANITEC KOŁO", "KOLO"], "category": "DIY"},
    {"name": "Roca Polska", "search_names": ["ROCA POLSKA"], "category": "DIY"},
    {"name": "Grohe Polska", "search_names": ["GROHE POLSKA", "GROHE"], "category": "DIY"},
    {"name": "Hansgrohe Polska", "search_names": ["HANSGROHE POLSKA", "HANSGROHE"], "category": "DIY"},
    {"name": "Baltic Wood (flooring)", "search_names": ["BALTIC WOOD"], "category": "DIY"},
    {"name": "Barlinek (flooring)", "search_names": ["BARLINEK"], "category": "DIY"},
    {"name": "Kronopol (panels)", "search_names": ["KRONOPOL", "SWISS KRONO"], "category": "DIY"},
    {"name": "Classen Polska (panels)", "search_names": ["CLASSEN POLSKA", "CLASSEN"], "category": "DIY"},

    # ─── Garden / outdoor ────────────────────────────────────────────────
    {"name": "Stihl Polska", "search_names": ["STIHL POLSKA", "STIHL"], "category": "DIY"},
    {"name": "Husqvarna Polska", "search_names": ["HUSQVARNA POLSKA", "HUSQVARNA"], "category": "DIY"},
    {"name": "Gardena Polska", "search_names": ["GARDENA POLSKA", "GARDENA"], "category": "DIY"},

    # ─── Tools / power tools distributors ────────────────────────────────
    {"name": "Bosch Polska (power tools)", "search_names": ["ROBERT BOSCH", "BOSCH POLSKA"], "category": "DIY"},
    {"name": "Makita Polska", "search_names": ["MAKITA POLSKA", "MAKITA"], "category": "DIY"},
    {"name": "DeWalt (Stanley Black & Decker)", "search_names": ["STANLEY BLACK DECKER", "DEWALT POLSKA"], "category": "DIY"},
    {"name": "Milwaukee Tool (TTI)", "search_names": ["TECHTRONIC INDUSTRIES", "MILWAUKEE"], "category": "DIY"},
    {"name": "Festool Polska", "search_names": ["FESTOOL POLSKA", "FESTOOL"], "category": "DIY"},

    # ─── Heating / HVAC / insulation ─────────────────────────────────────
    {"name": "Viessmann Polska", "search_names": ["VIESSMANN"], "category": "DIY"},
    {"name": "Vaillant Polska", "search_names": ["VAILLANT SAUNIER DUVAL"], "category": "DIY"},
    {"name": "Buderus Polska", "search_names": ["BUDERUS POLSKA", "BUDERUS"], "category": "DIY"},
    {"name": "Rockwool Polska", "search_names": ["ROCKWOOL POLSKA"], "category": "DIY"},
    {"name": "Isover (Saint-Gobain)", "search_names": ["SAINT-GOBAIN CONSTRUCTION", "ISOVER"], "category": "DIY"},
    {"name": "Knauf Polska", "search_names": ["KNAUF", "KNAUF POLSKA"], "category": "DIY"},
    {"name": "Rigips (Saint-Gobain)", "search_names": ["RIGIPS"], "category": "DIY"},

    # ─── Electrical / lighting chains ────────────────────────────────────
    {"name": "Elektroskandia Polska", "search_names": ["ELEKTROSKANDIA"], "category": "DIY"},
    {"name": "TIM SA (electrical wholesaler)", "search_names": ["TIM"], "category": "DIY"},
    {"name": "Kanlux (lighting)", "search_names": ["KANLUX"], "category": "DIY"},
    {"name": "Nowodvorski Lighting", "search_names": ["NOWODVORSKI"], "category": "DIY"},
    {"name": "Philips Lighting Polska", "search_names": ["SIGNIFY POLAND"], "category": "DIY"},

    # ─── Furniture / home (with DIY sections) ────────────────────────────
    {"name": "IKEA Polska", "search_names": ["IKEA RETAIL", "IKEA POLSKA", "INGKA"], "category": "DIY"},
    {"name": "Jysk Polska", "search_names": ["JYSK"], "category": "DIY"},
    {"name": "Agata Meble", "search_names": ["AGATA", "AGATA MEBLE"], "category": "DIY"},
    {"name": "Black Red White", "search_names": ["BLACK RED WHITE", "BRW"], "category": "DIY"},
    {"name": "VOX Meble", "search_names": ["VOX INDUSTRIE", "VOX MEBLE"], "category": "DIY"},
    {"name": "Abra Meble", "search_names": ["ABRA", "ABRA MEBLE"], "category": "DIY"},
    {"name": "Bodzio Meble", "search_names": ["BODZIO", "FABRYKA MEBLI BODZIO"], "category": "DIY"},
    {"name": "Forte Meble", "search_names": ["FORTE", "FABRYKA MEBLI FORTE"], "category": "DIY"},
    {"name": "Sklepy Komfort (Sołowow)", "search_names": ["SKLEPY KOMFORT"], "category": "DIY"},

    # ─── Discount/variety with DIY overlap ───────────────────────────────
    {"name": "Action Polska", "search_names": ["ACTION POLSKA", "ACTION S.A. DISCOUNT"], "category": "DIY"},
    {"name": "Pepco Polska", "search_names": ["PEPCO POLAND"], "category": "DIY"},
    {"name": "TEDi Polska", "search_names": ["TEDI"], "category": "DIY"},
    {"name": "Dealz (Pepco Group)", "search_names": ["DEALZ"], "category": "DIY"},

    # ─── Regional / smaller DIY ──────────────────────────────────────────
    {"name": "Marma Polskie Folie", "search_names": ["MARMA POLSKIE FOLIE"], "category": "DIY"},
    {"name": "Grupa SBS (składy budowlane)", "search_names": ["GRUPA SBS", "SBS"], "category": "DIY"},
    {"name": "Hurtownia Grodno", "search_names": ["GRODNO"], "category": "DIY"},
    {"name": "Magnat (farby)", "search_names": ["MAGNAT", "FFiL SNIEZKA"], "category": "DIY"},
    {"name": "Polbruk", "search_names": ["POLBRUK"], "category": "DIY"},
    {"name": "Styropmin (insulation)", "search_names": ["STYROPMIN"], "category": "DIY"},
    {"name": "Centrobud", "search_names": ["CENTROBUD"], "category": "DIY"},
]

MOBILE_CHAINS = [
    # ─── Mobile network operators (MNO) with retail stores ───────────────
    {"name": "Orange Polska", "search_names": ["Orange Polska", "ORANGE POLSKA"], "category": "Mobile", "krs": "0000010681", "nip": "5260250995"},
    {"name": "T-Mobile Polska", "search_names": ["T-Mobile Polska", "T-MOBILE POLSKA"], "category": "Mobile", "krs": "0000391193", "nip": "5261040567"},
    {"name": "Play (P4 Sp. z o.o.)", "search_names": ["P4 Sp. z o.o.", "P4"], "category": "Mobile", "krs": "0000217207", "nip": "5213236440"},
    {"name": "Plus (Polkomtel)", "search_names": ["Polkomtel", "POLKOMTEL"], "category": "Mobile", "krs": "0000419430", "nip": "5271037727"},
    {"name": "Vectra", "search_names": ["VECTRA"], "category": "Mobile", "krs": "0000076684"},
    {"name": "UPC Polska (Play)", "search_names": ["UPC POLSKA"], "category": "Mobile", "krs": "0000273136"},
    {"name": "Netia", "search_names": ["NETIA"], "category": "Mobile", "krs": "0000041649", "nip": "5262361475"},
    {"name": "Inea", "search_names": ["INEA"], "category": "Mobile", "krs": "0000063469"},

    # ─── MVNOs (Virtual operators) ───────────────────────────────────────
    {"name": "Virgin Mobile Polska", "search_names": ["Virgin Mobile", "VIRGIN MOBILE POLSKA"], "category": "Mobile"},
    {"name": "Lycamobile", "search_names": ["Lycamobile", "LYCAMOBILE POLSKA"], "category": "Mobile"},
    {"name": "Premium Mobile", "search_names": ["PREMIUM MOBILE"], "category": "Mobile", "krs": "0000337781"},
    {"name": "Multimedia Polska", "search_names": ["MULTIMEDIA POLSKA"], "category": "Mobile", "krs": "0000065242"},
    {"name": "Cyfrowy Polsat (Grupa Polsat Plus)", "search_names": ["CYFROWY POLSAT"], "category": "Mobile", "krs": "0000010078", "nip": "7961810732"},

    # ─── Major electronics retail chains ─────────────────────────────────
    {"name": "Media Expert (TERG)", "search_names": ["Media Expert", "TERG", "MEDIA EXPERT"], "category": "Mobile", "krs": "0000048007", "nip": "7251752913"},
    {"name": "Media Markt (MediaMarktSaturn)", "search_names": ["Media Markt", "MEDIA SATURN HOLDING POLSKA", "MEDIAMARKT"], "category": "Mobile", "krs": "0000028842"},
    {"name": "RTV Euro AGD (Euro-net)", "search_names": ["Euro AGD", "RTV EURO AGD", "EURO-NET"], "category": "Mobile", "krs": "0000117710", "nip": "5270004159"},
    {"name": "Neonet (x-kom group)", "search_names": ["NEONET"], "category": "Mobile", "krs": "0000263628"},
    {"name": "Max Elektro", "search_names": ["MAX ELEKTRO"], "category": "Mobile"},
    {"name": "Kakto", "search_names": ["KAKTO"], "category": "Mobile"},
    {"name": "Rebel Electro", "search_names": ["REBEL ELECTRO"], "category": "Mobile"},
    {"name": "Avans", "search_names": ["AVANS"], "category": "Mobile"},
    {"name": "Mix Electronics", "search_names": ["MIX ELECTRONICS"], "category": "Mobile"},

    # ─── IT / computer / phone specialist chains ─────────────────────────
    {"name": "x-kom", "search_names": ["x-kom", "X-KOM"], "category": "Mobile", "nip": "6443492570"},
    {"name": "Komputronik", "search_names": ["Komputronik", "KOMPUTRONIK"], "category": "Mobile", "krs": "0000270885", "nip": "6342525416"},
    {"name": "Morele.net", "search_names": ["Morele.net", "MORELE NET", "MORELE"], "category": "Mobile", "krs": "0000393967"},
    {"name": "Sferis (Action)", "search_names": ["SFERIS"], "category": "Mobile"},
    {"name": "iSpot (Apple Premium Reseller)", "search_names": ["ISPOT"], "category": "Mobile"},
    {"name": "Cortland (Apple reseller)", "search_names": ["CORTLAND"], "category": "Mobile"},

    # ─── Phone manufacturer official stores ──────────────────────────────
    {"name": "Samsung Polska", "search_names": ["Samsung Electronics Polska", "SAMSUNG ELECTRONICS POLSKA"], "category": "Mobile", "krs": "0000128080", "nip": "5260152505"},
    {"name": "Xiaomi Polska (Mi Store)", "search_names": ["XIAOMI POLSKA", "MI STORE", "XIAOMI"], "category": "Mobile"},
    {"name": "Huawei Polska", "search_names": ["Huawei Polska", "HUAWEI POLSKA"], "category": "Mobile", "krs": "0000140955"},
    {"name": "Sony Polska", "search_names": ["SONY POLSKA", "SONY EUROPE"], "category": "Mobile"},
    {"name": "Motorola / Lenovo Polska", "search_names": ["MOTOROLA POLSKA", "LENOVO POLSKA"], "category": "Mobile", "krs": "0000187782"},
    {"name": "Nokia (HMD Global)", "search_names": ["HMD GLOBAL"], "category": "Mobile"},
    {"name": "Oppo Polska", "search_names": ["OPPO POLSKA", "OPPO"], "category": "Mobile"},
    {"name": "realme Polska", "search_names": ["REALME POLSKA", "REALME"], "category": "Mobile"},
    {"name": "Google Polska", "search_names": ["GOOGLE POLAND", "GOOGLE POLSKA"], "category": "Mobile", "krs": "0000267967"},
    {"name": "ASBIS Polska", "search_names": ["ASBIS POLSKA", "ASBIS"], "category": "Mobile", "krs": "0000037705"},
    {"name": "Mi-Home.pl (Xiaomi stores)", "search_names": ["MI-HOME", "MI HOME"], "category": "Mobile"},

    # ─── Phone / IT distributors / wholesalers ───────────────────────────
    {"name": "ABC Data", "search_names": ["ABC DATA"], "category": "Mobile", "krs": "0000287132", "nip": "5242617178"},
    {"name": "Also Polska", "search_names": ["ALSO POLSKA"], "category": "Mobile"},
    {"name": "AB SA (IT distributor)", "search_names": ["AB SA", "AB SPOLKA AKCYJNA"], "category": "Mobile", "krs": "0000053834", "nip": "8951615977"},
    {"name": "Ingram Micro Polska", "search_names": ["INGRAM MICRO"], "category": "Mobile", "krs": "0000038636"},
    {"name": "Action SA", "search_names": ["ACTION", "ACTION SA"], "category": "Mobile", "krs": "0000214038", "nip": "5271107221"},
    {"name": "TelForceOne", "search_names": ["TelForceOne", "TELFORCEONE"], "category": "Mobile", "krs": "0000252002", "nip": "8982030993"},

    # ─── Phone accessories / repair / Polish brands ──────────────────────
    {"name": "GSM Service", "search_names": ["GSM SERVICE"], "category": "Mobile"},
    {"name": "Teletorium", "search_names": ["TELETORIUM"], "category": "Mobile"},
    {"name": "Teleakces.com", "search_names": ["TELEAKCES"], "category": "Mobile"},
    {"name": "CCS (Cyfrowe Centrum Serwisowe)", "search_names": ["CYFROWE CENTRUM SERWISOWE", "CCS"], "category": "Mobile"},
    {"name": "Cordon Electronics (sbe-online)", "search_names": ["CORDON ELECTRONICS"], "category": "Mobile"},
    {"name": "MyPhone / mPTech (Polish brand)", "search_names": ["MPTECH", "MYPHONE"], "category": "Mobile", "krs": "0000371014"},
    {"name": "Maxcom (Polish brand)", "search_names": ["MAXCOM"], "category": "Mobile", "krs": "0000076759"},
    {"name": "Kruger&Matz", "search_names": ["KRUGER MATZ", "KRUGER"], "category": "Mobile"},

    # ─── E-commerce / marketplaces ───────────────────────────────────────
    {"name": "Allegro", "search_names": ["ALLEGRO"], "category": "Mobile", "krs": "0000635012", "nip": "5272806857"},
    {"name": "Amazon Polska", "search_names": ["AMAZON POLSKA"], "category": "Mobile"},
    {"name": "Empik (electronics)", "search_names": ["EMPIK"], "category": "Mobile", "krs": "0000024654"},

    # ─── Hypermarkets with electronics/phone sections ────────────────────
    {"name": "Auchan Polska", "search_names": ["AUCHAN POLSKA"], "category": "Mobile", "krs": "0000028508"},
    {"name": "Carrefour Polska", "search_names": ["CARREFOUR POLSKA"], "category": "Mobile", "krs": "0000049087"},
    {"name": "Kaufland Polska", "search_names": ["KAUFLAND POLSKA"], "category": "Mobile", "krs": "0000091702"},
    {"name": "E.Leclerc Polska", "search_names": ["E.LECLERC", "LECLERC POLSKA"], "category": "Mobile"},
]

OUTPUT_COLUMNS = [
    "chain_name",
    "category",
    "search_method",
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

LOCAL_UNIT_COLUMNS = [
    "chain_name",
    "parent_nip",
    "parent_regon",
    "parent_name",
    "regon14",
    "unit_name",
    "street",
    "building",
    "unit_number",
    "zip_code",
    "city",
    "province",
    "county",
    "municipality",
    "pkd_main",
    "full_address",
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


def lookup_by_krs(api_client, chain_info):
    """Look up a chain using its known KRS number (free API, no auth needed)."""
    krs = chain_info.get("krs", "")
    if not krs or not krs.isdigit():
        return None

    logger.info(f"  KRS lookup: {krs}")
    result = api_client.search_krs(krs)
    if result:
        result["_search_term"] = f"KRS:{krs}"
    return result


def lookup_by_nip_gus(api_client, chain_info):
    """Look up a chain using its known NIP via GUS."""
    nip = chain_info.get("nip", "")
    if not nip:
        return None

    logger.info(f"  GUS NIP lookup: {nip}")
    result = api_client.search_gus_by_nip(nip)
    if result:
        result["_search_term"] = f"NIP:{nip}"
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Look up Polish DIY and mobile phone retail chains via GUS + KRS APIs."
    )
    parser.add_argument("--output", "-o", default="chains.xlsx", help="Output Excel file")
    parser.add_argument("--category", "-c", default="all",
                        choices=["all", "diy", "mobile"],
                        help="Which category to search (default: all)")
    parser.add_argument("--gus-key", help="GUS API key (or set GUS_API_KEY env var)")
    parser.add_argument("--sandbox", action="store_true", help="Use GUS sandbox")
    parser.add_argument("--krs-only", action="store_true",
                        help="Use ONLY the free KRS API (no GUS key needed). "
                        "Only finds chains with known KRS numbers.")
    parser.add_argument("--local-units", action="store_true",
                        help="Also fetch local units (branch locations) for each chain found. "
                        "Requires GUS API key.")
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
    has_gus = bool(api_client.gus_api_key) and not args.krs_only

    if args.krs_only:
        krs_chains = [c for c in chains if c.get("krs") and c["krs"].isdigit()]
        logger.info(f"KRS-only mode: {len(krs_chains)} chains have known KRS numbers "
                     f"(out of {len(chains)} total)")
        logger.info("KRS API is FREE — no API key needed.")
    elif not has_gus:
        krs_chains = [c for c in chains if c.get("krs") and c["krs"].isdigit()]
        logger.warning(f"No GUS API key. Falling back to KRS-only mode.")
        logger.info(f"  {len(krs_chains)} chains have known KRS numbers.")
        args.krs_only = True

    if args.local_units and not has_gus:
        logger.warning("--local-units requires GUS API key. Local units will be skipped.")
        args.local_units = False

    # Process
    all_rows = []
    local_unit_rows = []
    stats = {"found": 0, "not_found": 0, "total_entities": 0, "local_units": 0}

    for idx, chain in enumerate(chains):
        logger.info(f"[{idx + 1}/{len(chains)}] {chain['name']} ({chain['category']})")

        result = None
        results = []

        if args.krs_only:
            # KRS-only mode: use known KRS number
            if chain.get("krs") and chain["krs"].isdigit():
                r = lookup_by_krs(api_client, chain)
                if r:
                    results = [r]
            else:
                logger.info(f"  -> No KRS number known, skipping (KRS-only mode)")
        else:
            # Full mode: try GUS name search first
            results = search_chain(api_client, chain)

            # If GUS name search failed, try by known NIP
            if not results and chain.get("nip"):
                r = lookup_by_nip_gus(api_client, chain)
                if r:
                    results = [r]

            # If still nothing, try KRS as fallback
            if not results and chain.get("krs") and chain["krs"].isdigit():
                r = lookup_by_krs(api_client, chain)
                if r:
                    results = [r]

        if results:
            stats["found"] += 1
            stats["total_entities"] += len(results)
            for r in results:
                search_term = r.pop("_search_term", chain["search_names"][0])

                # Enrich with KRS if not already from KRS
                if r.get("krs") and r.get("source") != "KRS":
                    krs_data = api_client.search_krs(r["krs"])
                    if krs_data:
                        r = api_client._merge_results(r, krs_data)
                elif not r.get("krs") and chain.get("krs") and chain["krs"].isdigit():
                    krs_data = api_client.search_krs(chain["krs"])
                    if krs_data:
                        r = api_client._merge_results(r, krs_data)

                # Enrich with GUS by NIP for extra details (if we have GUS)
                if has_gus and r.get("source") == "KRS" and r.get("nip"):
                    gus_data = api_client.search_gus_by_nip(r["nip"])
                    if gus_data:
                        r = api_client._merge_results(gus_data, r)
                        logger.info(f"  GUS enrichment OK for NIP {r['nip']}")

                all_rows.append(build_row(chain, r, search_term))

                # Fetch local units if requested
                if args.local_units and r.get("regon") and r.get("entity_type"):
                    units = api_client.fetch_local_units(r["regon"], r["entity_type"])
                    for u in units:
                        local_unit_rows.append({
                            "chain_name": chain["name"],
                            "parent_nip": r.get("nip", ""),
                            "parent_regon": r.get("regon", ""),
                            "parent_name": r.get("name", ""),
                            "regon14": u.get("regon14", ""),
                            "unit_name": u.get("name", ""),
                            "street": u.get("street", ""),
                            "building": u.get("building", ""),
                            "unit_number": u.get("unit_number", ""),
                            "zip_code": u.get("zip_code", ""),
                            "city": u.get("city", ""),
                            "province": u.get("province", ""),
                            "county": u.get("county", ""),
                            "municipality": u.get("municipality", ""),
                            "pkd_main": u.get("pkd_main", ""),
                            "full_address": u.get("full_address", ""),
                        })
                    stats["local_units"] += len(units)

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
        df_all.to_excel(writer, index=False, sheet_name="All Chains")

        df_diy = df_all[df_all["category"] == "DIY"]
        if not df_diy.empty:
            df_diy.to_excel(writer, index=False, sheet_name="DIY Chains")

        df_mobile = df_all[df_all["category"] == "Mobile"]
        if not df_mobile.empty:
            df_mobile.to_excel(writer, index=False, sheet_name="Mobile Chains")

        # Local units sheet
        if local_unit_rows:
            df_units = pd.DataFrame(local_unit_rows, columns=LOCAL_UNIT_COLUMNS).fillna("")
            df_units.to_excel(writer, index=False, sheet_name="Local Units")
            logger.info(f"  Sheet 'Local Units': {len(df_units)} branch locations")

        # Has local units (chains with num_local_units > 0)
        has_units = df_all[df_all["num_local_units"].astype(str).str.match(r'^[1-9]')]
        if not has_units.empty:
            has_units.to_excel(writer, index=False, sheet_name="Has Local Units")

        # Summary
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
    if args.local_units:
        logger.info(f"  Local units:     {stats['local_units']}")
    logger.info(f"Results: {args.output}")
    if args.krs_only:
        logger.info(f"NOTE: Used KRS-only mode. For more results, provide a GUS API key.")


if __name__ == "__main__":
    main()
