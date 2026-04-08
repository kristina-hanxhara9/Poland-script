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
]

MOBILE_CHAINS = [
    # ─── Mobile network operators (MNO) with retail stores ───────────────
    {"name": "Orange Polska", "search_names": ["Orange Polska", "ORANGE POLSKA"], "category": "Mobile"},
    {"name": "T-Mobile Polska", "search_names": ["T-Mobile Polska", "T-MOBILE POLSKA"], "category": "Mobile"},
    {"name": "Play (P4 Sp. z o.o.)", "search_names": ["P4 Sp. z o.o.", "P4"], "category": "Mobile"},
    {"name": "Plus (Polkomtel)", "search_names": ["Polkomtel", "POLKOMTEL"], "category": "Mobile"},
    {"name": "Vectra", "search_names": ["VECTRA"], "category": "Mobile"},
    {"name": "UPC Polska (Play)", "search_names": ["UPC POLSKA"], "category": "Mobile"},
    {"name": "Netia", "search_names": ["NETIA"], "category": "Mobile"},
    {"name": "Inea", "search_names": ["INEA"], "category": "Mobile"},

    # ─── MVNOs (Virtual operators) ───────────────────────────────────────
    {"name": "Virgin Mobile Polska", "search_names": ["Virgin Mobile", "VIRGIN MOBILE POLSKA"], "category": "Mobile"},
    {"name": "Nju Mobile (Orange)", "search_names": ["NJU MOBILE"], "category": "Mobile"},
    {"name": "Lajt Mobile", "search_names": ["LAJT MOBILE", "LAJT"], "category": "Mobile"},
    {"name": "Lycamobile", "search_names": ["Lycamobile", "LYCAMOBILE POLSKA"], "category": "Mobile"},
    {"name": "Premium Mobile", "search_names": ["PREMIUM MOBILE"], "category": "Mobile"},
    {"name": "Aero2", "search_names": ["AERO2"], "category": "Mobile"},
    {"name": "Heyah (T-Mobile)", "search_names": ["HEYAH"], "category": "Mobile"},
    {"name": "Red Bull Mobile", "search_names": ["RED BULL MOBILE"], "category": "Mobile"},
    {"name": "Plush (Plus)", "search_names": ["PLUSH"], "category": "Mobile"},
    {"name": "Fakt Mobile", "search_names": ["FAKT MOBILE"], "category": "Mobile"},
    {"name": "Klucz Mobile", "search_names": ["KLUCZ MOBILE"], "category": "Mobile"},
    {"name": "Lemon Mobile", "search_names": ["LEMON MOBILE"], "category": "Mobile"},
    {"name": "a2mobile", "search_names": ["A2MOBILE"], "category": "Mobile"},
    {"name": "Tuya Mobile", "search_names": ["TUYA"], "category": "Mobile"},

    # ─── Additional MVNOs ──────────────────────────────────────────────
    {"name": "Mobile Vikings", "search_names": ["MOBILE VIKINGS"], "category": "Mobile"},
    {"name": "Vectra Mobile", "search_names": ["VECTRA MOBILE"], "category": "Mobile"},
    {"name": "FM GROUP Mobile", "search_names": ["FM GROUP MOBILE", "FM GROUP"], "category": "Mobile"},
    {"name": "Otvarta", "search_names": ["OTVARTA"], "category": "Mobile"},
    {"name": "Sat Film", "search_names": ["SAT FILM"], "category": "Mobile"},
    {"name": "Fonia Telekom", "search_names": ["FONIA TELEKOM"], "category": "Mobile"},
    {"name": "Multimedia Polska", "search_names": ["MULTIMEDIA POLSKA"], "category": "Mobile"},
    {"name": "Canal+ Polska (nc+ Mobile)", "search_names": ["CANAL PLUS POLSKA", "CANAL+"], "category": "Mobile"},

    # ─── Major electronics retail chains ─────────────────────────────────
    {"name": "Media Expert (TERG)", "search_names": ["Media Expert", "TERG", "MEDIA EXPERT"], "category": "Mobile"},
    {"name": "Media Markt (MediaMarktSaturn)", "search_names": ["Media Markt", "MEDIA SATURN HOLDING POLSKA", "MEDIAMARKT"], "category": "Mobile"},
    {"name": "RTV Euro AGD (Euro-net)", "search_names": ["Euro AGD", "RTV EURO AGD", "EURO-NET"], "category": "Mobile"},
    {"name": "Neonet (x-kom group)", "search_names": ["NEONET"], "category": "Mobile"},
    {"name": "Max Elektro", "search_names": ["MAX ELEKTRO"], "category": "Mobile"},
    {"name": "Kakto", "search_names": ["KAKTO"], "category": "Mobile"},
    {"name": "Rebel Electro", "search_names": ["REBEL ELECTRO"], "category": "Mobile"},
    {"name": "Electro (Neonet group)", "search_names": ["ELECTRO"], "category": "Mobile"},
    {"name": "Avans", "search_names": ["AVANS"], "category": "Mobile"},
    {"name": "Mix Electronics", "search_names": ["MIX ELECTRONICS"], "category": "Mobile"},
    {"name": "OleOle (Euro-net online)", "search_names": ["OLEOLE", "OLE OLE"], "category": "Mobile"},

    # ─── IT / computer / phone specialist chains ─────────────────────────
    {"name": "x-kom", "search_names": ["x-kom", "X-KOM"], "category": "Mobile"},
    {"name": "Komputronik", "search_names": ["Komputronik", "KOMPUTRONIK"], "category": "Mobile"},
    {"name": "Morele.net", "search_names": ["Morele.net", "MORELE NET", "MORELE"], "category": "Mobile"},
    {"name": "al.to (x-kom group)", "search_names": ["AL.TO", "ALTO"], "category": "Mobile"},
    {"name": "Sferis", "search_names": ["SFERIS"], "category": "Mobile"},
    {"name": "Vobis", "search_names": ["VOBIS"], "category": "Mobile"},
    {"name": "iSource (Apple reseller)", "search_names": ["ISOURCE"], "category": "Mobile"},
    {"name": "Cortland (Apple reseller)", "search_names": ["CORTLAND"], "category": "Mobile"},

    # ─── Phone manufacturer official stores ──────────────────────────────
    {"name": "Samsung Polska", "search_names": ["Samsung Electronics Polska", "SAMSUNG ELECTRONICS POLSKA"], "category": "Mobile"},
    {"name": "Apple Polska (iSpot, iDream)", "search_names": ["APPLE POLSKA", "ISPOT", "IDREAM"], "category": "Mobile"},
    {"name": "Xiaomi Polska (Mi Store)", "search_names": ["XIAOMI POLSKA", "MI STORE", "XIAOMI"], "category": "Mobile"},
    {"name": "Huawei Polska", "search_names": ["Huawei Polska", "HUAWEI POLSKA"], "category": "Mobile"},
    {"name": "Sony Polska", "search_names": ["SONY POLSKA", "SONY EUROPE"], "category": "Mobile"},
    {"name": "Motorola / Lenovo Polska", "search_names": ["MOTOROLA POLSKA", "LENOVO POLSKA"], "category": "Mobile"},
    {"name": "Nokia (HMD Global)", "search_names": ["HMD GLOBAL"], "category": "Mobile"},
    {"name": "Oppo Polska", "search_names": ["OPPO POLSKA", "OPPO"], "category": "Mobile"},
    {"name": "realme Polska", "search_names": ["REALME POLSKA", "REALME"], "category": "Mobile"},
    {"name": "Honor Polska", "search_names": ["HONOR POLSKA", "HONOR"], "category": "Mobile"},
    {"name": "Google Polska", "search_names": ["GOOGLE POLAND", "GOOGLE POLSKA"], "category": "Mobile"},

    # ─── Phone / IT distributors / wholesalers ───────────────────────────
    {"name": "ABC Data", "search_names": ["ABC DATA"], "category": "Mobile"},
    {"name": "Also Polska", "search_names": ["ALSO POLSKA"], "category": "Mobile"},
    {"name": "AB SA (IT distributor)", "search_names": ["AB SA", "AB SPOLKA AKCYJNA"], "category": "Mobile"},
    {"name": "Ingram Micro Polska", "search_names": ["INGRAM MICRO"], "category": "Mobile"},
    {"name": "Tech Data / TD Synnex", "search_names": ["TECH DATA POLSKA", "TD SYNNEX"], "category": "Mobile"},
    {"name": "Action SA", "search_names": ["ACTION", "ACTION SA"], "category": "Mobile"},
    {"name": "TelForceOne", "search_names": ["TelForceOne", "TELFORCEONE"], "category": "Mobile"},
    {"name": "EET Europarts Polska", "search_names": ["EET EUROPARTS"], "category": "Mobile"},

    # ─── Phone accessories / repair / Polish brands ──────────────────────
    {"name": "GSM Service", "search_names": ["GSM SERVICE"], "category": "Mobile"},
    {"name": "Teletorium", "search_names": ["TELETORIUM"], "category": "Mobile"},
    {"name": "MyPhone (Mobiway)", "search_names": ["MYPHONE", "MOBIWAY"], "category": "Mobile"},
    {"name": "Maxcom (Polish brand)", "search_names": ["MAXCOM"], "category": "Mobile"},
    {"name": "Kruger&Matz", "search_names": ["KRUGER MATZ", "KRUGER"], "category": "Mobile"},

    # ─── E-commerce / marketplaces ───────────────────────────────────────
    {"name": "Allegro", "search_names": ["ALLEGRO"], "category": "Mobile"},
    {"name": "Amazon Polska", "search_names": ["AMAZON POLSKA"], "category": "Mobile"},
    {"name": "Empik (electronics)", "search_names": ["EMPIK"], "category": "Mobile"},

    # ─── Hypermarkets with electronics/phone sections ────────────────────
    {"name": "Auchan Polska", "search_names": ["AUCHAN POLSKA"], "category": "Mobile"},
    {"name": "Carrefour Polska", "search_names": ["CARREFOUR POLSKA"], "category": "Mobile"},
    {"name": "Kaufland Polska", "search_names": ["KAUFLAND POLSKA"], "category": "Mobile"},
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
