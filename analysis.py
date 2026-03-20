"""
PKD code analysis and reporting for Polish businesses.

Analyzes which PKD codes correspond to each channel category
(DIY shops, Paint specialists, Builders merchants) and performs
keyword matching on business names.
"""

import logging
from collections import Counter, defaultdict

import pandas as pd
from tabulate import tabulate

from pkd_codes import (
    classify_by_pkd,
    get_pkd_description,
    match_keywords,
    get_all_target_categories,
    get_expected_pkd_codes,
    CATEGORY_PKD_MAPPING,
)

logger = logging.getLogger(__name__)

TARGET_CATEGORIES = ["DIY shops", "Paint specialists", "Builders merchants"]


def load_excel(filepath, channel_column="channel"):
    """
    Load an Excel file and filter for target channel categories.

    Args:
        filepath: Path to the Excel file
        channel_column: Name of the column containing channel/category info

    Returns:
        Tuple of (full_df, filtered_df) where filtered_df only has target categories.
    """
    df = pd.read_excel(filepath)

    # Normalize column names (lowercase, strip whitespace)
    df.columns = [c.strip().lower() for c in df.columns]

    # Try to find the channel column
    channel_col = None
    for col in df.columns:
        if "channel" in col or "kategori" in col or "category" in col:
            channel_col = col
            break

    if channel_col is None:
        # Fall back to the provided column name
        channel_col = channel_column.lower()

    if channel_col not in df.columns:
        raise ValueError(
            f"Channel column '{channel_col}' not found. "
            f"Available columns: {list(df.columns)}"
        )

    logger.info(f"Loaded {len(df)} rows. Using channel column: '{channel_col}'")
    logger.info(f"Unique channels: {df[channel_col].unique().tolist()}")

    # Filter for target categories (case-insensitive partial match)
    mask = df[channel_col].str.lower().str.strip().apply(
        lambda x: any(cat.lower() in str(x) for cat in TARGET_CATEGORIES)
    )
    filtered = df[mask].copy()

    # Normalize the channel values to standard names
    def normalize_channel(val):
        val_lower = str(val).lower().strip()
        for cat in TARGET_CATEGORIES:
            if cat.lower() in val_lower:
                return cat
        return val

    filtered["channel_normalized"] = filtered[channel_col].apply(normalize_channel)

    logger.info(f"Filtered to {len(filtered)} rows matching target categories")

    return df, filtered, channel_col


def enrich_with_api(df, api_client):
    """
    Enrich DataFrame with PKD codes from Poland government APIs.

    Args:
        df: DataFrame with 'name' and 'zip' columns
        api_client: PolandAPIClient instance

    Returns:
        DataFrame with added pkd_codes, pkd_main, and pkd_descriptions columns.
    """
    pkd_codes_list = []
    pkd_main_list = []
    api_source_list = []

    for idx, row in df.iterrows():
        name = str(row.get("name", "")).strip()
        zip_code = str(row.get("zip", "")).strip()

        if not name:
            pkd_codes_list.append([])
            pkd_main_list.append("")
            api_source_list.append("none")
            continue

        result = api_client.lookup_business(name, zip_code)

        if result:
            pkd_codes_list.append(result.get("pkd_codes", []))
            pkd_main_list.append(result.get("pkd_main", ""))
            api_source_list.append(result.get("source", "API"))
            logger.info(
                f"Found: {name} -> PKD: {result.get('pkd_codes', [])} ({result.get('source', '')})"
            )
        else:
            pkd_codes_list.append([])
            pkd_main_list.append("")
            api_source_list.append("not_found")
            logger.info(f"Not found: {name} ({zip_code})")

    df["pkd_codes"] = pkd_codes_list
    df["pkd_main"] = pkd_main_list
    df["pkd_descriptions"] = df["pkd_codes"].apply(
        lambda codes: [get_pkd_description(c) for c in codes]
    )
    df["api_source"] = api_source_list

    return df


def enrich_with_keywords(df):
    """
    Enrich DataFrame with keyword matching results.

    For each business, checks how well its name matches keywords
    for each target category.
    """
    keyword_results = []

    for idx, row in df.iterrows():
        name = str(row.get("name", "")).strip()
        channel = row.get("channel_normalized", "")

        # Match against the assigned category
        result = match_keywords(name, channel)
        keyword_results.append(result)

    df["keyword_score"] = [r["score"] for r in keyword_results]
    df["keyword_strength"] = [r["strength"] for r in keyword_results]
    df["matched_keywords"] = [", ".join(r["matched_keywords"]) for r in keyword_results]

    return df


def analyze_pkd_distribution(df):
    """
    Analyze PKD code distribution across channel categories.

    Returns a dict with analysis results per category.
    """
    analysis = {}

    for category in TARGET_CATEGORIES:
        cat_df = df[df["channel_normalized"] == category]
        if cat_df.empty:
            analysis[category] = {
                "count": 0,
                "pkd_distribution": {},
                "has_pkd_data": 0,
                "no_pkd_data": 0,
            }
            continue

        # Flatten all PKD codes for this category
        all_pkd = []
        for codes in cat_df["pkd_codes"]:
            if isinstance(codes, list):
                all_pkd.extend(codes)

        pkd_counter = Counter(all_pkd)
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
                code: count
                for code, count in pkd_counter.items()
                if code not in expected_pkd
            },
        }

    return analysis


def analyze_cross_channel(df):
    """
    Analyze which PKD codes appear across multiple channel categories.
    """
    pkd_by_channel = defaultdict(set)

    for _, row in df.iterrows():
        channel = row.get("channel_normalized", "")
        codes = row.get("pkd_codes", [])
        if isinstance(codes, list):
            for code in codes:
                pkd_by_channel[code].add(channel)

    cross_channel = {
        code: list(channels)
        for code, channels in pkd_by_channel.items()
        if len(channels) > 1
    }

    return cross_channel


def analyze_keyword_matching(df):
    """
    Analyze keyword matching results per category.
    """
    analysis = {}

    for category in TARGET_CATEGORIES:
        cat_df = df[df["channel_normalized"] == category]
        if cat_df.empty:
            analysis[category] = {"count": 0}
            continue

        analysis[category] = {
            "count": len(cat_df),
            "strong_matches": sum(1 for s in cat_df["keyword_strength"] if s == "strong"),
            "moderate_matches": sum(1 for s in cat_df["keyword_strength"] if s == "moderate"),
            "no_matches": sum(1 for s in cat_df["keyword_strength"] if s == "none"),
            "avg_score": cat_df["keyword_score"].mean(),
        }

    return analysis


def generate_report(df, pkd_analysis, cross_channel, keyword_analysis):
    """
    Generate a formatted text report of the analysis.

    Returns the report as a string.
    """
    lines = []
    lines.append("=" * 70)
    lines.append("POLAND BUSINESS PKD CODE ANALYSIS REPORT")
    lines.append("=" * 70)
    lines.append("")

    # Summary
    lines.append("## SUMMARY")
    lines.append(f"Total businesses analyzed: {len(df)}")
    for cat in TARGET_CATEGORIES:
        count = len(df[df["channel_normalized"] == cat])
        lines.append(f"  - {cat}: {count}")
    lines.append("")

    # PKD Distribution per Channel
    lines.append("## PKD CODE DISTRIBUTION BY CHANNEL")
    lines.append("-" * 50)

    for category in TARGET_CATEGORIES:
        data = pkd_analysis.get(category, {})
        lines.append(f"\n### {category} ({data.get('count', 0)} businesses)")
        lines.append(f"  Businesses with PKD data: {data.get('has_pkd_data', 0)}")
        lines.append(f"  Businesses without PKD data: {data.get('no_pkd_data', 0)}")

        dist = data.get("pkd_distribution", {})
        if dist:
            lines.append("\n  PKD Code | Count | Description")
            lines.append("  " + "-" * 60)
            for code, count in sorted(dist.items(), key=lambda x: x[1], reverse=True):
                desc = get_pkd_description(code)
                lines.append(f"  {code:10s} | {count:5d} | {desc}")

            # Expected vs actual
            expected = data.get("expected_pkd_codes", [])
            lines.append(f"\n  Expected PKD codes for this category: {', '.join(expected)}")
            lines.append(f"  Matches with expected codes: {data.get('expected_match_count', 0)}")

            unexpected = data.get("unexpected_pkd_codes", {})
            if unexpected:
                lines.append(f"  Unexpected PKD codes found:")
                for code, count in unexpected.items():
                    desc = get_pkd_description(code)
                    lines.append(f"    {code}: {count}x - {desc}")
        else:
            lines.append("  No PKD data available (API lookup needed)")

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

    # Keyword matching analysis
    lines.append("## KEYWORD MATCHING ANALYSIS")
    lines.append("-" * 50)
    for category in TARGET_CATEGORIES:
        data = keyword_analysis.get(category, {})
        count = data.get("count", 0)
        if count == 0:
            continue
        lines.append(f"\n### {category} ({count} businesses)")
        lines.append(f"  Strong keyword matches: {data.get('strong_matches', 0)} ({data.get('strong_matches', 0)/count*100:.1f}%)")
        lines.append(f"  Moderate keyword matches: {data.get('moderate_matches', 0)} ({data.get('moderate_matches', 0)/count*100:.1f}%)")
        lines.append(f"  No keyword matches: {data.get('no_matches', 0)} ({data.get('no_matches', 0)/count*100:.1f}%)")
        lines.append(f"  Average confidence score: {data.get('avg_score', 0):.2f}")

    lines.append("")

    # PKD-to-category match quality
    lines.append("## PKD-CHANNEL ALIGNMENT ANALYSIS")
    lines.append("-" * 50)
    lines.append("How well do the retrieved PKD codes align with the assigned channel categories?")
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

        total = len(cat_df)
        lines.append(f"### {category}")
        lines.append(f"  Aligned (PKD matches expected): {aligned} ({aligned/total*100:.1f}%)")
        lines.append(f"  Misaligned (unexpected PKD): {misaligned} ({misaligned/total*100:.1f}%)")
        lines.append(f"  No PKD data: {no_data} ({no_data/total*100:.1f}%)")
        lines.append("")

    lines.append("=" * 70)
    lines.append("END OF REPORT")
    lines.append("=" * 70)

    return "\n".join(lines)


def run_full_analysis(df, api_client=None):
    """
    Run the complete analysis pipeline on a filtered DataFrame.

    Args:
        df: DataFrame filtered to target categories (must have channel_normalized column)
        api_client: Optional PolandAPIClient for PKD lookups

    Returns:
        Tuple of (enriched_df, report_text)
    """
    # Step 1: Enrich with API data if available
    if api_client and api_client.has_api_access():
        logger.info("Looking up businesses via Poland government APIs...")
        df = enrich_with_api(df, api_client)
    else:
        logger.info("No API access configured. Using keyword matching only.")
        df["pkd_codes"] = [[] for _ in range(len(df))]
        df["pkd_main"] = ""
        df["pkd_descriptions"] = [[] for _ in range(len(df))]
        df["api_source"] = "none"

    # Step 2: Enrich with keyword matching
    logger.info("Running keyword matching analysis...")
    df = enrich_with_keywords(df)

    # Step 3: Run analyses
    logger.info("Analyzing PKD code distribution...")
    pkd_analysis = analyze_pkd_distribution(df)
    cross_channel = analyze_cross_channel(df)
    keyword_analysis = analyze_keyword_matching(df)

    # Step 4: Generate report
    report = generate_report(df, pkd_analysis, cross_channel, keyword_analysis)

    return df, report
