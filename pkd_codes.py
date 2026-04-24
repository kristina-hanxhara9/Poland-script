"""
PKD Code definitions, category classification, and keyword matching for Polish businesses.

PKD (Polska Klasyfikacja Działalności) is Poland's official business activity classification system.
"""

# Comprehensive PKD code dictionary for construction, retail, and wholesale sectors
PKD_CODES = {
    # Retail sale
    "47.52.Z": "Retail sale of hardware, paints and glass in specialised stores",
    "47.53.Z": "Retail sale of carpets, rugs, wall and floor coverings in specialised stores",
    "47.59.Z": "Retail sale of furniture, lighting equipment and other household articles in specialised stores",
    "47.54.Z": "Retail sale of electrical household appliances in specialised stores",
    "47.78.Z": "Other retail sale of new goods in specialised stores",
    "47.19.Z": "Other retail sale in non-specialised stores",
    "47.91.Z": "Retail sale via mail order houses or via Internet",
    # Wholesale
    "46.73.Z": "Wholesale of wood, construction materials and sanitary equipment",
    "46.74.Z": "Wholesale of hardware, plumbing and heating equipment and supplies",
    "46.44.Z": "Wholesale of china and glassware and cleaning materials",
    "46.49.Z": "Wholesale of other household goods",
    "46.90.Z": "Non-specialised wholesale trade",
    "46.13.Z": "Agents involved in the sale of timber and building materials",
    "46.15.Z": "Agents involved in the sale of furniture, household goods, hardware and ironmongery",
    # Manufacturing
    "20.30.Z": "Manufacture of paints, varnishes and similar coatings, printing ink and mastics",
    "23.61.Z": "Manufacture of concrete products for construction purposes",
    "23.63.Z": "Manufacture of ready-mixed concrete",
    "23.32.Z": "Manufacture of bricks, tiles and construction products, in baked clay",
    "16.10.Z": "Sawmilling and planing of wood",
    "16.23.Z": "Manufacture of other builders' carpentry and joinery",
    "25.72.Z": "Manufacture of locks and hinges",
    # Construction
    "41.20.Z": "Construction of residential and non-residential buildings",
    "43.11.Z": "Demolition",
    "43.12.Z": "Site preparation",
    "43.29.Z": "Other construction installation",
    "43.31.Z": "Plastering",
    "43.32.Z": "Joinery installation",
    "43.33.Z": "Floor and wall covering",
    "43.34.Z": "Painting and glazing",
    "43.39.Z": "Other building completion and finishing",
    "43.91.Z": "Roofing activities",
    "43.99.Z": "Other specialised construction activities n.e.c.",
    # Mobile phone / telecom retail
    "47.42.Z": "Retail sale of telecommunications equipment in specialised stores",
    "46.52.Z": "Wholesale of electronic and telecommunications equipment and parts",
    "95.12.Z": "Repair of communication equipment",
    "47.41.Z": "Retail sale of computers, peripheral units and software in specialised stores",
    "47.43.Z": "Retail sale of audio and video equipment in specialised stores",
    "46.43.Z": "Wholesale of electrical household appliances",
    "61.10.Z": "Wired telecommunications activities",
    "61.20.Z": "Wireless telecommunications activities",
    "61.30.Z": "Satellite telecommunications activities",
    "61.90.Z": "Other telecommunications activities",
    "77.22.Z": "Renting of video tapes and disks (includes phone rental)",
    "26.30.Z": "Manufacture of communication equipment",
    "26.12.Z": "Manufacture of loaded electronic boards",
}

# PKD codes mapped to each target channel category
CATEGORY_PKD_MAPPING = {
    "DIY shops": {
        "primary": ["47.52.Z"],
        "secondary": ["47.53.Z", "47.59.Z", "47.54.Z", "47.78.Z", "47.19.Z", "47.91.Z"],
    },
    "Paint specialists": {
        "primary": ["47.52.Z", "20.30.Z"],
        "secondary": ["46.44.Z", "43.34.Z", "46.73.Z"],
    },
    "Mobile phone specialists": {
        "primary": ["47.42.Z", "95.12.Z"],
        "secondary": ["46.52.Z", "47.41.Z", "47.43.Z", "61.10.Z", "61.20.Z", "61.90.Z", "26.30.Z"],
    },
    "Builders merchants": {
        "primary": ["46.73.Z", "46.74.Z"],
        "secondary": ["47.52.Z", "46.13.Z", "46.15.Z", "46.90.Z"],
    },
}

# Keywords for matching business names to categories (Polish + English)
CATEGORY_KEYWORDS = {
    "DIY shops": {
        "strong": [
            "castorama", "leroy merlin", "obi", "bricomarché", "bricomarche",
            "bricoman", "mrówka", "mrowka", "nomi", "majster",
            "market budowlany", "dom i ogród", "dom i ogrod",
            "majsterkowanie", "diy", "do it yourself",
        ],
        "moderate": [
            "narzędzia", "narzedzia", "budowlany", "hobby", "ogrodniczy",
            "techniczny", "domowy", "hardware", "home improvement",
            "tools", "garden", "artykuły budowlane", "artykuly budowlane",
        ],
    },
    "Paint specialists": {
        "strong": [
            "farby", "lakiery", "farba", "lakier", "malowanie",
            "dulux", "śnieżka", "sniezka", "dekoral", "beckers",
            "tikkurila", "caparol", "sigma", "hempel",
            "paint", "paints", "coatings",
        ],
        "moderate": [
            "dekoracje", "tynki", "tynk", "pigment", "emulsja",
            "gruntowanie", "grunt", "impregnat", "lazura",
            "kolorystyka", "barwniki", "rozpuszczalnik",
            "varnish", "stain", "primer", "coating",
        ],
    },
    "Builders merchants": {
        "strong": [
            "materiały budowlane", "materialy budowlane",
            "skład budowlany", "sklad budowlany",
            "hurtownia budowlana", "centrum budowlane",
            "builders merchant", "building materials", "building supplies",
            "psg", "grupa psb", "bat",
        ],
        "moderate": [
            "cement", "beton", "cegła", "cegla", "dachówka", "dachowka",
            "izolacja", "styropian", "wełna mineralna", "welna mineralna",
            "klej", "zaprawa", "piasek", "żwir", "zwir", "stal", "drewno",
            "rury", "kanalizacja", "hydraulika", "ogrzewanie",
            "insulation", "timber", "lumber", "plumbing", "roofing",
        ],
    },
}


def classify_by_pkd(pkd_code):
    """
    Classify a PKD code into channel categories.

    Returns a list of dicts with category name and match strength ('primary' or 'secondary').
    """
    matches = []
    for category, codes in CATEGORY_PKD_MAPPING.items():
        if pkd_code in codes["primary"]:
            matches.append({"category": category, "strength": "primary"})
        elif pkd_code in codes["secondary"]:
            matches.append({"category": category, "strength": "secondary"})
    return matches


def get_pkd_description(pkd_code):
    """Return the description for a PKD code, or 'Unknown' if not in our dictionary."""
    return PKD_CODES.get(pkd_code, "Unknown PKD code")


def match_keywords(business_name, category):
    """
    Match a business name against keywords for a given category.

    Returns a dict with:
      - matched_keywords: list of matched keywords
      - score: float 0-1 indicating match confidence
      - strength: 'strong', 'moderate', or 'none'
    """
    if category not in CATEGORY_KEYWORDS:
        return {"matched_keywords": [], "score": 0.0, "strength": "none"}

    name_lower = business_name.lower().strip()
    strong_matches = [kw for kw in CATEGORY_KEYWORDS[category]["strong"] if kw in name_lower]
    moderate_matches = [kw for kw in CATEGORY_KEYWORDS[category]["moderate"] if kw in name_lower]

    if strong_matches:
        score = min(1.0, 0.7 + 0.1 * len(strong_matches) + 0.05 * len(moderate_matches))
        strength = "strong"
    elif moderate_matches:
        score = min(0.6, 0.2 + 0.1 * len(moderate_matches))
        strength = "moderate"
    else:
        score = 0.0
        strength = "none"

    return {
        "matched_keywords": strong_matches + moderate_matches,
        "score": score,
        "strength": strength,
    }


def get_all_target_categories():
    """Return the list of target channel categories."""
    return list(CATEGORY_PKD_MAPPING.keys())


def get_expected_pkd_codes(category):
    """Return all expected PKD codes (primary + secondary) for a category."""
    if category not in CATEGORY_PKD_MAPPING:
        return []
    mapping = CATEGORY_PKD_MAPPING[category]
    return mapping["primary"] + mapping["secondary"]
