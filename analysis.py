"""
PKD code analysis and reporting for Polish businesses.

Looks up ALL businesses via Poland government APIs, extracts full details,
analyzes which PKD codes correspond to each channel category
(DIY shops, Paint specialists, Builders merchants), and performs
keyword matching on all business names.
"""

import logging
from collections import Counter, defaultdict

import pandas as pd

from pkd_codes import (
    classify_by_pkd,
    get_pkd_description,
    match_keywords,
    get_all_target_categories,
    get_expected_pkd_codes,
    CATEGORY_PKD_MAPPING,
    CATEGORY_KEYWORDS,
)

logger = logging.getLogger(__name__)

TARGET_CATEGORIES = ["DIY shops", "Paint specialists", "Builders merchants"]

# All enriched fields we extract from APIs
API_FIELDS = [
    "api_name", "nip", "regon", "krs",
    "address_str", "street", "building", "unit", "zip_code_api", "city_api",
    "municipality", "county", "province", "country",
    "mailing_address_str",
    "status", "is_active",
    "date_started", "date_ended", "date_suspended", "date_resumed", "date_deleted",
    "pkd_main", "pkd_codes", "pkd_descriptions",
    "owner_first_name", "owner_last_name",
    "email", "phone", "website", "ceidg_link",
    "api_source",
]


def load_excel(filepath, channel_column="channel"):
    """
    Load an Excel file with business data.

    Expects columns: name, zip, and optionally channel/channel online/channel offline.

    Returns:
        Tuple of (full_df, channel_col_online, channel_col_offline)
    """
    df = pd.read_excel(filepath)

    # Normalize column names (lowercase, strip whitespace)
    df.columns = [c.strip().lower() for c in df.columns]

    logger.info(f"Loaded {len(df)} rows. Columns: {list(df.columns)}")

    # Find channel columns
    channel_col_online = None
    channel_col_offline = None
    channel_col_generic = None

    for col in df.columns:
        col_lower = col.lower()
        if "channel" in col_lower and "online" in col_lower:
            channel_col_online = col
        elif "channel" in col_lower and "offline" in col_lower:
            channel_col_offline = col
        elif "channel" in col_lower or "kategori" in col_lower or "category" in col_lower:
            channel_col_generic = col

    # If no online/offline split, use generic channel column
    if not channel_col_online and not channel_col_offline:
        channel_col_generic = channel_col_generic or channel_column.lower()

    if channel_col_online:
        logger.info(f"Found channel online column: '{channel_col_online}'")
    if channel_col_offline:
        logger.info(f"Found channel offline column: '{channel_col_offline}'")
    if channel_col_generic and not channel_col_online and not channel_col_offline:
        logger.info(f"Found channel column: '{channel_col_generic}'")

    # Determine the best category for each row
    def classify_row(row):
        """Check all channel columns and classify into target categories."""
        channels = []
        for col in [channel_col_online, channel_col_offline, channel_col_generic]:
            if col and col in row.index:
                val = str(row[col]).lower().strip()
                if val and val != "nan":
                    channels.append(val)

        for cat in TARGET_CATEGORIES:
            if any(cat.lower() in ch for ch in channels):
                return cat
        return ", ".join(channels) if channels else "Unknown"

    df["channel_normalized"] = df.apply(classify_row, axis=1)

    channel_cols = {
        "online": channel_col_online,
        "offline": channel_col_offline,
        "generic": channel_col_generic,
    }

    logger.info(f"Channel distribution:\n{df['channel_normalized'].value_counts().to_string()}")

    return df, channel_cols


def enrich_with_api(df, api_client):
    """
    Enrich ALL rows in DataFrame with full business data from Poland APIs.

    Uses name and zip columns to look up each business.
    """
    enriched_data = {field: [] for field in API_FIELDS}

    total = len(df)
    for i, (idx, row) in enumerate(df.iterrows()):
        name = str(row.get("name", "")).strip()
        zip_code = str(row.get("zip", "")).strip()
        if zip_code == "nan":
            zip_code = ""

        # Try to extract city from the data if available
        city = ""
        for col in ["city", "miasto", "location"]:
            if col in row.index:
                city = str(row.get(col, "")).strip()
                if city and city != "nan":
                    break
                city = ""

        if not name or name == "nan":
            for field in API_FIELDS:
                enriched_data[field].append("" if field != "pkd_codes" else [])
            enriched_data["pkd_descriptions"][-1] = []
            enriched_data["is_active"][-1] = ""
            continue

        logger.info(f"[{i+1}/{total}] Looking up: {name} ({zip_code})")
        result = api_client.lookup_business(name, zip_code, city)

        if result:
            enriched_data["api_name"].append(result.get("name", ""))
            enriched_data["nip"].append(result.get("nip", ""))
            enriched_data["regon"].append(result.get("regon", ""))
            enriched_data["krs"].append(result.get("krs", ""))
            enriched_data["address_str"].append(result.get("address_str", ""))
            enriched_data["street"].append(result.get("street", ""))
            enriched_data["building"].append(result.get("building", ""))
            enriched_data["unit"].append(result.get("unit", ""))
            enriched_data["zip_code_api"].append(result.get("zip_code", ""))
            enriched_data["city_api"].append(result.get("city", ""))
            enriched_data["municipality"].append(result.get("municipality", ""))
            enriched_data["county"].append(result.get("county", ""))
            enriched_data["province"].append(result.get("province", ""))
            enriched_data["country"].append(result.get("country", ""))
            enriched_data["mailing_address_str"].append(result.get("mailing_address_str", ""))
            enriched_data["status"].append(result.get("status", ""))
            enriched_data["is_active"].append(result.get("is_active", ""))
            enriched_data["date_started"].append(result.get("date_started", ""))
            enriched_data["date_ended"].append(result.get("date_ended", ""))
            enriched_data["date_suspended"].append(result.get("date_suspended", ""))
            enriched_data["date_resumed"].append(result.get("date_resumed", ""))
            enriched_data["date_deleted"].append(result.get("date_deleted", ""))
            enriched_data["pkd_main"].append(result.get("pkd_main", ""))
            pkd_codes = result.get("pkd_codes", [])
            enriched_data["pkd_codes"].append(pkd_codes)
            enriched_data["pkd_descriptions"].append(
                [get_pkd_description(c) for c in pkd_codes]
            )
            enriched_data["owner_first_name"].append(result.get("owner_first_name", ""))
            enriched_data["owner_last_name"].append(result.get("owner_last_name", ""))
            enriched_data["email"].append(result.get("email", ""))
            enriched_data["phone"].append(result.get("phone", ""))
            enriched_data["website"].append(result.get("website", ""))
            enriched_data["ceidg_link"].append(result.get("ceidg_link", ""))
            enriched_data["api_source"].append(result.get("source", ""))

            logger.info(
                f"  Found: {result.get('name', '')} | Status: {result.get('status', '')} | "
                f"PKD: {pkd_codes} | Source: {result.get('source', '')}"
            )
        else:
            for field in API_FIELDS:
                if field == "pkd_codes":
                    enriched_data[field].append([])
                elif field == "pkd_descriptions":
                    enriched_data[field].append([])
                elif field == "is_active":
                    enriched_data[field].append("")
                else:
                    enriched_data[field].append("")
            enriched_data["api_source"][-1] = "not_found"
            logger.info(f"  Not found")

    # Add all enriched columns to the dataframe
    for field in API_FIELDS:
        df[field] = enriched_data[field]

    return df


def enrich_with_keywords(df):
    """
    Enrich DataFrame with keyword matching results for ALL businesses
    against ALL target categories.
    """
    # Match each business against ALL categories
    best_category = []
    best_score = []
    best_strength = []
    best_keywords = []

    # Also store per-category results
    for cat in TARGET_CATEGORIES:
        df[f"kw_{cat}_score"] = 0.0
        df[f"kw_{cat}_strength"] = "none"
        df[f"kw_{cat}_keywords"] = ""

    for idx, row in df.iterrows():
        name = str(row.get("name", "")).strip()

        top_cat = "none"
        top_score = 0.0
        top_strength = "none"
        top_keywords = []

        for cat in TARGET_CATEGORIES:
            result = match_keywords(name, cat)
            df.at[idx, f"kw_{cat}_score"] = result["score"]
            df.at[idx, f"kw_{cat}_strength"] = result["strength"]
            df.at[idx, f"kw_{cat}_keywords"] = ", ".join(result["matched_keywords"])

            if result["score"] > top_score:
                top_score = result["score"]
                top_cat = cat
                top_strength = result["strength"]
                top_keywords = result["matched_keywords"]

        best_category.append(top_cat if top_score > 0 else "none")
        best_score.append(top_score)
        best_strength.append(top_strength)
        best_keywords.append(", ".join(top_keywords))

    df["keyword_best_category"] = best_category
    df["keyword_score"] = best_score
    df["keyword_strength"] = best_strength
    df["matched_keywords"] = best_keywords

    return df


def enrich_with_pkd_classification(df):
    """
    For each business, classify its PKD codes into target categories.
    """
    pkd_category = []
    pkd_category_strength = []

    for idx, row in df.iterrows():
        codes = row.get("pkd_codes", [])
        if not isinstance(codes, list):
            codes = []

        best_cat = "Unclassified"
        best_str = "none"

        for code in codes:
            classifications = classify_by_pkd(code)
            for cl in classifications:
                if cl["strength"] == "primary":
                    best_cat = cl["category"]
                    best_str = "primary"
                    break
                elif cl["strength"] == "secondary" and best_str != "primary":
                    best_cat = cl["category"]
                    best_str = "secondary"
            if best_str == "primary":
                break

        pkd_category.append(best_cat)
        pkd_category_strength.append(best_str)

    df["pkd_suggested_category"] = pkd_category
    df["pkd_category_strength"] = pkd_category_strength

    return df


def analyze_pkd_distribution(df):
    """Analyze PKD code distribution across all businesses and per category."""
    analysis = {}

    # Overall
    all_pkd = []
    for codes in df["pkd_codes"]:
        if isinstance(codes, list):
            all_pkd.extend(codes)
    overall_counter = Counter(all_pkd)
    analysis["_overall"] = {
        "count": len(df),
        "pkd_distribution": dict(overall_counter.most_common(30)),
        "has_pkd_data": sum(1 for codes in df["pkd_codes"] if codes),
        "no_pkd_data": sum(1 for codes in df["pkd_codes"] if not codes),
    }

    # Per target category (based on channel_normalized)
    for category in TARGET_CATEGORIES:
        cat_df = df[df["channel_normalized"] == category]
        if cat_df.empty:
            analysis[category] = {
                "count": 0, "pkd_distribution": {},
                "has_pkd_data": 0, "no_pkd_data": 0,
            }
            continue

        cat_pkd = []
        for codes in cat_df["pkd_codes"]:
            if isinstance(codes, list):
                cat_pkd.extend(codes)

        pkd_counter = Counter(cat_pkd)
        expected_pkd = get_expected_pkd_codes(category)

        analysis[category] = {
            "count": len(cat_df),
            "pkd_distribution": dict(pkd_counter.most_common()),
            "has_pkd_data": sum(1 for codes in cat_df["pkd_codes"] if codes),
            "no_pkd_data": sum(1 for codes in cat_df["pkd_codes"] if not codes),
            "expected_pkd_codes": expected_pkd,
            "expected_match_count": sum(
                pkd_counter.get(code, 0) for code in expected_pkd
            ),
            "unexpected_pkd_codes": {
                code: count for code, count in pkd_counter.items()
                if code not in expected_pkd
            },
        }

    return analysis


def analyze_cross_channel(df):
    """Analyze which PKD codes appear across multiple channel categories."""
    pkd_by_channel = defaultdict(set)
    for _, row in df.iterrows():
        channel = row.get("channel_normalized", "")
        codes = row.get("pkd_codes", [])
        if isinstance(codes, list):
            for code in codes:
                pkd_by_channel[code].add(channel)

    return {
        code: list(channels)
        for code, channels in pkd_by_channel.items()
        if len(channels) > 1
    }


def analyze_keyword_matching(df):
    """Analyze keyword matching results overall and per category."""
    analysis = {}

    # Per target category
    for category in TARGET_CATEGORIES:
        cat_df = df[df["channel_normalized"] == category]
        if cat_df.empty:
            analysis[category] = {"count": 0}
            continue

        analysis[category] = {
            "count": len(cat_df),
            "strong_matches": sum(1 for s in cat_df[f"kw_{category}_strength"] if s == "strong"),
            "moderate_matches": sum(1 for s in cat_df[f"kw_{category}_strength"] if s == "moderate"),
            "no_matches": sum(1 for s in cat_df[f"kw_{category}_strength"] if s == "none"),
            "avg_score": cat_df[f"kw_{category}_score"].mean(),
        }

    # Overall keyword detection across all businesses
    analysis["_overall"] = {
        "count": len(df),
        "any_match": sum(1 for s in df["keyword_strength"] if s != "none"),
        "strong_matches": sum(1 for s in df["keyword_strength"] if s == "strong"),
        "moderate_matches": sum(1 for s in df["keyword_strength"] if s == "moderate"),
        "no_matches": sum(1 for s in df["keyword_strength"] if s == "none"),
    }

    return analysis


def generate_report(df, pkd_analysis, cross_channel, keyword_analysis):
    """Generate a formatted text report of the analysis."""
    lines = []
    lines.append("=" * 70)
    lines.append("POLAND BUSINESS PKD/SIC CODE ANALYSIS REPORT")
    lines.append("=" * 70)
    lines.append("")

    # Summary
    total = len(df)
    lines.append("## SUMMARY")
    lines.append(f"Total businesses analyzed: {total}")
    has_pkd = pkd_analysis["_overall"]["has_pkd_data"]
    no_pkd = pkd_analysis["_overall"]["no_pkd_data"]
    lines.append(f"Businesses with PKD data from API: {has_pkd} ({has_pkd/total*100:.1f}%)")
    lines.append(f"Businesses without PKD data: {no_pkd} ({no_pkd/total*100:.1f}%)")
    lines.append("")

    # Active vs inactive
    active = sum(1 for v in df["is_active"] if v is True)
    inactive = sum(1 for v in df["is_active"] if v is False)
    unknown = total - active - inactive
    lines.append(f"Active businesses: {active}")
    lines.append(f"Inactive/deleted businesses: {inactive}")
    lines.append(f"Status unknown: {unknown}")
    lines.append("")

    # API sources
    if "api_source" in df.columns:
        source_counts = df["api_source"].value_counts()
        lines.append("Data sources:")
        for src, count in source_counts.items():
            lines.append(f"  - {src}: {count}")
        lines.append("")

    # Channel distribution
    lines.append("## CHANNEL DISTRIBUTION")
    lines.append("-" * 50)
    for channel, count in df["channel_normalized"].value_counts().items():
        lines.append(f"  {channel}: {count}")
    lines.append("")

    # Overall PKD distribution (top 30)
    lines.append("## TOP PKD CODES (ALL BUSINESSES)")
    lines.append("-" * 50)
    dist = pkd_analysis["_overall"]["pkd_distribution"]
    if dist:
        lines.append("  PKD Code   | Count | Description")
        lines.append("  " + "-" * 60)
        for code, count in sorted(dist.items(), key=lambda x: x[1], reverse=True):
            desc = get_pkd_description(code)
            lines.append(f"  {code:10s} | {count:5d} | {desc}")
    lines.append("")

    # PKD Distribution per target channel
    lines.append("## PKD CODE DISTRIBUTION BY TARGET CHANNEL")
    lines.append("-" * 50)

    for category in TARGET_CATEGORIES:
        data = pkd_analysis.get(category, {})
        count = data.get("count", 0)
        lines.append(f"\n### {category} ({count} businesses)")
        lines.append(f"  With PKD data: {data.get('has_pkd_data', 0)}")
        lines.append(f"  Without PKD data: {data.get('no_pkd_data', 0)}")

        cat_dist = data.get("pkd_distribution", {})
        if cat_dist:
            lines.append("\n  PKD Code   | Count | Description")
            lines.append("  " + "-" * 60)
            for code, cnt in sorted(cat_dist.items(), key=lambda x: x[1], reverse=True):
                desc = get_pkd_description(code)
                lines.append(f"  {code:10s} | {cnt:5d} | {desc}")

            expected = data.get("expected_pkd_codes", [])
            lines.append(f"\n  Expected PKD codes: {', '.join(expected)}")
            lines.append(f"  Matches with expected: {data.get('expected_match_count', 0)}")

            unexpected = data.get("unexpected_pkd_codes", {})
            if unexpected:
                lines.append(f"  Unexpected PKD codes:")
                for code, cnt in unexpected.items():
                    desc = get_pkd_description(code)
                    lines.append(f"    {code}: {cnt}x - {desc}")
        else:
            lines.append("  No PKD data available")

    lines.append("")

    # Cross-channel analysis
    lines.append("## CROSS-CHANNEL PKD CODE ANALYSIS")
    lines.append("-" * 50)
    if cross_channel:
        lines.append("PKD codes appearing in multiple channel categories:")
        for code, channels in cross_channel.items():
            desc = get_pkd_description(code)
            lines.append(f"  {code} ({desc})")
            lines.append(f"    Found in: {', '.join(channels)}")
    else:
        lines.append("No PKD codes found across multiple categories.")
    lines.append("")

    # Keyword matching
    lines.append("## KEYWORD MATCHING ANALYSIS")
    lines.append("-" * 50)
    overall_kw = keyword_analysis.get("_overall", {})
    lines.append(f"Total businesses: {overall_kw.get('count', 0)}")
    lines.append(f"Any keyword match: {overall_kw.get('any_match', 0)}")
    lines.append(f"Strong matches: {overall_kw.get('strong_matches', 0)}")
    lines.append(f"Moderate matches: {overall_kw.get('moderate_matches', 0)}")
    lines.append(f"No matches: {overall_kw.get('no_matches', 0)}")
    lines.append("")

    for category in TARGET_CATEGORIES:
        data = keyword_analysis.get(category, {})
        count = data.get("count", 0)
        if count == 0:
            continue
        lines.append(f"### {category} ({count} businesses in this channel)")
        lines.append(f"  Strong keyword matches: {data.get('strong_matches', 0)} ({data.get('strong_matches', 0)/count*100:.1f}%)")
        lines.append(f"  Moderate keyword matches: {data.get('moderate_matches', 0)} ({data.get('moderate_matches', 0)/count*100:.1f}%)")
        lines.append(f"  No keyword matches: {data.get('no_matches', 0)} ({data.get('no_matches', 0)/count*100:.1f}%)")
        lines.append(f"  Average confidence score: {data.get('avg_score', 0):.2f}")
        lines.append("")

    # PKD-channel alignment
    lines.append("## PKD-CHANNEL ALIGNMENT ANALYSIS")
    lines.append("-" * 50)
    lines.append("How well do retrieved PKD codes align with assigned channels?")
    lines.append("")

    for category in TARGET_CATEGORIES:
        cat_df = df[df["channel_normalized"] == category]
        if cat_df.empty:
            continue

        expected = set(get_expected_pkd_codes(category))
        aligned = 0
        misaligned = 0
        no_data = 0

        for _, row in cat_df.iterrows():
            codes = row.get("pkd_codes", [])
            if not codes:
                no_data += 1
            elif any(c in expected for c in codes):
                aligned += 1
            else:
                misaligned += 1

        cat_total = len(cat_df)
        lines.append(f"### {category}")
        lines.append(f"  Aligned (PKD matches expected): {aligned} ({aligned/cat_total*100:.1f}%)")
        lines.append(f"  Misaligned (unexpected PKD): {misaligned} ({misaligned/cat_total*100:.1f}%)")
        lines.append(f"  No PKD data: {no_data} ({no_data/cat_total*100:.1f}%)")
        lines.append("")

    # PKD-based suggested reclassification
    lines.append("## PKD-BASED CATEGORY SUGGESTIONS")
    lines.append("-" * 50)
    lines.append("Businesses whose PKD codes suggest a different category than assigned:")
    lines.append("")

    reclassified = 0
    for _, row in df.iterrows():
        assigned = row.get("channel_normalized", "")
        suggested = row.get("pkd_suggested_category", "Unclassified")
        name = row.get("name", "")

        if suggested != "Unclassified" and assigned in TARGET_CATEGORIES and suggested != assigned:
            lines.append(f"  {name}")
            lines.append(f"    Assigned: {assigned} -> PKD suggests: {suggested}")
            lines.append(f"    PKD codes: {row.get('pkd_codes', [])}")
            reclassified += 1

    if reclassified == 0:
        lines.append("  No reclassification suggestions found.")
    else:
        lines.append(f"\n  Total potential reclassifications: {reclassified}")
    lines.append("")

    lines.append("=" * 70)
    lines.append("END OF REPORT")
    lines.append("=" * 70)

    return "\n".join(lines)


def run_full_analysis(df, api_client=None):
    """
    Run the complete analysis pipeline on ALL businesses.

    Args:
        df: Full DataFrame (all rows, not filtered)
        api_client: PolandAPIClient instance

    Returns:
        Tuple of (enriched_df, report_text)
    """
    # Step 1: Enrich ALL businesses with API data
    if api_client and api_client.has_api_access():
        logger.info(f"Looking up {len(df)} businesses via Poland government APIs...")
        df = enrich_with_api(df, api_client)
    else:
        logger.info("No API access configured. Using keyword matching only.")
        for field in API_FIELDS:
            if field in ("pkd_codes", "pkd_descriptions"):
                df[field] = [[] for _ in range(len(df))]
            elif field == "is_active":
                df[field] = ""
            else:
                df[field] = ""

    # Step 2: Keyword matching against ALL categories for every business
    logger.info("Running keyword matching analysis...")
    df = enrich_with_keywords(df)

    # Step 3: PKD-based classification
    logger.info("Classifying businesses by PKD codes...")
    df = enrich_with_pkd_classification(df)

    # Step 4: Run analyses
    logger.info("Analyzing PKD code distribution...")
    pkd_analysis = analyze_pkd_distribution(df)
    cross_channel = analyze_cross_channel(df)
    keyword_analysis = analyze_keyword_matching(df)

    # Step 5: Generate report
    report = generate_report(df, pkd_analysis, cross_channel, keyword_analysis)

    return df, report
