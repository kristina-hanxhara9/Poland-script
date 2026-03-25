#!/usr/bin/env python3
"""
Analyze PKD code distribution and business name keywords from results file.

Reads an Excel results file and generates insights about:
- PKD main code distribution (counts, percentages, top codes)
- All PKD codes frequency analysis
- Business name keyword frequency
- Cross-analysis: top keywords per PKD category

Usage:
    python analyze_pkd.py resultsmobile.xlsx --output pkd_insights.xlsx
"""

import argparse
import logging
import re
import sys
from collections import Counter

import pandas as pd

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

# Common Polish stopwords to exclude from keyword analysis
STOPWORDS = {
    "sp", "z", "o", "oo", "zo", "zoo", "sa", "sc", "sj", "sk",
    "spółka", "spolka", "jawna", "akcyjna", "cywilna", "komandytowa",
    "i", "w", "na", "do", "od", "dla", "po", "ze", "we", "nie",
    "ul", "nr", "im",
    "the", "and", "of", "in", "for", "a", "an",
    "sp.", "z.o.o.", "s.a.", "s.c.",
}


def extract_keywords(name):
    """Extract meaningful keywords from a business name."""
    if not name or str(name).lower() in ("nan", "none", ""):
        return []
    name = str(name).lower()
    # Remove punctuation except hyphens
    name = re.sub(r'[^\w\s\-]', ' ', name)
    words = name.split()
    # Filter out stopwords, short words, and numbers
    return [
        w for w in words
        if w not in STOPWORDS
        and len(w) > 1
        and not w.isdigit()
        and w not in ("sp", "zo", "oo")
    ]


def main():
    parser = argparse.ArgumentParser(description="Analyze PKD distribution and name keywords.")
    parser.add_argument("input_file", help="Excel results file (e.g. resultsmobile.xlsx)")
    parser.add_argument("--output", "-o", default="pkd_insights.xlsx", help="Output Excel file")
    parser.add_argument("--name-col", default="B", help="Column letter for business name (default: B)")
    parser.add_argument("--pkd-main-col", default="Q", help="Column letter for pkd_main (default: Q)")
    parser.add_argument("--pkd-codes-col", default=None, help="Column letter for all pkd_codes (optional)")
    parser.add_argument("--pkd-desc-col", default=None, help="Column letter for pkd_descriptions (optional)")
    args = parser.parse_args()

    # Load file
    logger.info(f"Reading {args.input_file}")
    try:
        df = pd.read_excel(args.input_file, header=0)
    except FileNotFoundError:
        logger.error(f"File not found: {args.input_file}")
        sys.exit(1)

    logger.info(f"Loaded {len(df)} rows, {len(df.columns)} columns")
    logger.info(f"Columns: {list(df.columns)}")

    # Resolve column letters to indices (A=0, B=1, ..., Q=16)
    def col_letter_to_idx(letter):
        letter = letter.upper().strip()
        idx = 0
        for ch in letter:
            idx = idx * 26 + (ord(ch) - ord('A') + 1)
        return idx - 1

    name_idx = col_letter_to_idx(args.name_col)
    pkd_main_idx = col_letter_to_idx(args.pkd_main_col)

    name_col = df.columns[name_idx]
    pkd_main_col = df.columns[pkd_main_idx]

    pkd_codes_col = None
    if args.pkd_codes_col:
        pkd_codes_col = df.columns[col_letter_to_idx(args.pkd_codes_col)]

    pkd_desc_col = None
    if args.pkd_desc_col:
        pkd_desc_col = df.columns[col_letter_to_idx(args.pkd_desc_col)]

    logger.info(f"Name column ({args.name_col}): '{name_col}'")
    logger.info(f"PKD main column ({args.pkd_main_col}): '{pkd_main_col}'")

    total = len(df)

    # ========================================
    # 1. PKD MAIN DISTRIBUTION
    # ========================================
    logger.info("Analyzing PKD main distribution...")

    pkd_values = df[pkd_main_col].astype(str).str.strip()
    pkd_values = pkd_values.replace(["nan", "None", ""], pd.NA).dropna()

    has_pkd = len(pkd_values)
    no_pkd = total - has_pkd

    pkd_counter = Counter(pkd_values)
    pkd_dist = []
    for code, count in pkd_counter.most_common():
        pct = count / total * 100
        pkd_dist.append({
            "pkd_code": code,
            "count": count,
            "percentage": round(pct, 1),
            "pct_of_those_with_pkd": round(count / has_pkd * 100, 1) if has_pkd else 0,
        })

    df_pkd_dist = pd.DataFrame(pkd_dist)

    # Try to add descriptions
    if pkd_desc_col:
        # Build a code->description map from the data
        desc_map = {}
        for _, row in df.iterrows():
            codes_str = str(row.get(pkd_codes_col, "")) if pkd_codes_col else str(row[pkd_main_col])
            descs_str = str(row.get(pkd_desc_col, ""))
            codes = [c.strip() for c in codes_str.split(";") if c.strip() and c.strip() != "nan"]
            descs = [d.strip() for d in descs_str.split(";") if d.strip() and d.strip() != "nan"]
            for c, d in zip(codes, descs):
                if c and d and c not in desc_map:
                    desc_map[c] = d
        df_pkd_dist["description"] = df_pkd_dist["pkd_code"].map(desc_map).fillna("")
    else:
        # Try pkd_codes module
        try:
            from pkd_codes import get_pkd_description
            df_pkd_dist["description"] = df_pkd_dist["pkd_code"].apply(get_pkd_description)
        except ImportError:
            df_pkd_dist["description"] = ""

    # ========================================
    # 2. ALL PKD CODES FREQUENCY (if available)
    # ========================================
    df_all_pkd = pd.DataFrame()
    if pkd_codes_col:
        logger.info("Analyzing all PKD codes frequency...")
        all_codes = Counter()
        desc_map_all = {}

        for _, row in df.iterrows():
            codes_str = str(row[pkd_codes_col])
            descs_str = str(row.get(pkd_desc_col, "")) if pkd_desc_col else ""
            codes = [c.strip() for c in codes_str.split(";") if c.strip() and c.strip() != "nan"]
            descs = [d.strip() for d in descs_str.split(";") if d.strip() and d.strip() != "nan"]
            for c in codes:
                all_codes[c] += 1
            for c, d in zip(codes, descs):
                if c and d and c not in desc_map_all:
                    desc_map_all[c] = d

        all_pkd_list = []
        for code, count in all_codes.most_common():
            all_pkd_list.append({
                "pkd_code": code,
                "businesses_with_code": count,
                "percentage": round(count / total * 100, 1),
                "description": desc_map_all.get(code, ""),
            })
        df_all_pkd = pd.DataFrame(all_pkd_list)

    # ========================================
    # 3. KEYWORD ANALYSIS FROM BUSINESS NAMES
    # ========================================
    logger.info("Analyzing business name keywords...")

    all_keywords = Counter()
    bigrams = Counter()

    for name in df[name_col]:
        words = extract_keywords(name)
        all_keywords.update(words)
        # Bigrams (2-word combos)
        for i in range(len(words) - 1):
            bigrams[f"{words[i]} {words[i+1]}"] += 1

    kw_list = []
    for word, count in all_keywords.most_common(10):
        kw_list.append({
            "keyword": word,
            "count": count,
            "percentage": round(count / total * 100, 1),
        })
    df_keywords = pd.DataFrame(kw_list)

    bigram_list = []
    for bg, count in bigrams.most_common(50):
        if count >= 2:  # Only show bigrams appearing 2+ times
            bigram_list.append({
                "keyword_pair": bg,
                "count": count,
                "percentage": round(count / total * 100, 1),
            })
    df_bigrams = pd.DataFrame(bigram_list)

    # ========================================
    # 4. TOP KEYWORDS PER PKD MAIN CODE
    # ========================================
    logger.info("Analyzing keywords per PKD category...")

    pkd_keywords = {}
    for _, row in df.iterrows():
        pkd = str(row[pkd_main_col]).strip()
        if pkd in ("nan", "None", ""):
            continue
        words = extract_keywords(row[name_col])
        if pkd not in pkd_keywords:
            pkd_keywords[pkd] = Counter()
        pkd_keywords[pkd].update(words)

    # Build a flat table: top 10 keywords per PKD code (for top 20 PKD codes)
    pkd_kw_rows = []
    top_pkd_codes = [code for code, _ in pkd_counter.most_common(20)]
    for pkd_code in top_pkd_codes:
        kw_counter = pkd_keywords.get(pkd_code, Counter())
        for word, count in kw_counter.most_common(10):
            pkd_kw_rows.append({
                "pkd_code": pkd_code,
                "keyword": word,
                "count": count,
            })
    df_pkd_kw = pd.DataFrame(pkd_kw_rows)

    # ========================================
    # 5. SUMMARY STATS
    # ========================================
    summary_rows = [
        {"metric": "Total businesses", "value": total},
        {"metric": "With PKD main code", "value": has_pkd},
        {"metric": "Without PKD main code", "value": no_pkd},
        {"metric": "PKD coverage %", "value": f"{has_pkd/total*100:.1f}%"},
        {"metric": "Unique PKD main codes", "value": len(pkd_counter)},
        {"metric": "Top PKD code", "value": pkd_counter.most_common(1)[0][0] if pkd_counter else "N/A"},
        {"metric": "Top PKD count", "value": pkd_counter.most_common(1)[0][1] if pkd_counter else 0},
        {"metric": "Unique keywords in names", "value": len(all_keywords)},
        {"metric": "Top keyword", "value": all_keywords.most_common(1)[0][0] if all_keywords else "N/A"},
        {"metric": "Top keyword count", "value": all_keywords.most_common(1)[0][1] if all_keywords else 0},
    ]

    if pkd_codes_col and df_all_pkd is not None and not df_all_pkd.empty:
        summary_rows.append({"metric": "Unique PKD codes (all)", "value": len(df_all_pkd)})
        summary_rows.append({"metric": "Most common PKD (all)", "value": df_all_pkd.iloc[0]["pkd_code"] if len(df_all_pkd) > 0 else "N/A"})

    df_summary = pd.DataFrame(summary_rows)

    # ========================================
    # SAVE TO EXCEL (multiple sheets)
    # ========================================
    logger.info(f"Saving insights to {args.output}")

    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        df_summary.to_excel(writer, sheet_name="Summary", index=False)
        df_pkd_dist.to_excel(writer, sheet_name="PKD Main Distribution", index=False)
        if not df_all_pkd.empty:
            df_all_pkd.to_excel(writer, sheet_name="All PKD Codes", index=False)
        df_keywords.to_excel(writer, sheet_name="Name Keywords Top10", index=False)
        if not df_bigrams.empty:
            df_bigrams.to_excel(writer, sheet_name="Name Keyword Pairs", index=False)
        if not df_pkd_kw.empty:
            df_pkd_kw.to_excel(writer, sheet_name="Keywords per PKD", index=False)

    # ========================================
    # ADD CHARTS
    # ========================================
    logger.info("Adding charts...")
    from openpyxl import load_workbook
    from openpyxl.chart import BarChart, PieChart, Reference

    wb = load_workbook(args.output)

    # --- Chart 1: Top 15 PKD Main Codes (bar chart) ---
    ws_pkd = wb["PKD Main Distribution"]
    chart_rows = min(len(df_pkd_dist), 15)

    bar1 = BarChart()
    bar1.type = "col"
    bar1.title = "Top 15 PKD Main Codes"
    bar1.x_axis.title = "PKD Code"
    bar1.y_axis.title = "Count"
    bar1.style = 10
    bar1.width = 25
    bar1.height = 14

    data_ref = Reference(ws_pkd, min_col=2, min_row=1, max_row=chart_rows + 1)
    cats_ref = Reference(ws_pkd, min_col=1, min_row=2, max_row=chart_rows + 1)
    bar1.add_data(data_ref, titles_from_data=True)
    bar1.set_categories(cats_ref)
    bar1.shape = 4
    ws_pkd.add_chart(bar1, f"A{len(df_pkd_dist) + 4}")

    # --- Chart 2: PKD Main Distribution (pie chart, top 10) ---
    pie1 = PieChart()
    pie1.title = "PKD Main Code Distribution (Top 10)"
    pie1.style = 10
    pie1.width = 20
    pie1.height = 14

    pie_rows = min(len(df_pkd_dist), 10)
    pie_data = Reference(ws_pkd, min_col=2, min_row=1, max_row=pie_rows + 1)
    pie_cats = Reference(ws_pkd, min_col=1, min_row=2, max_row=pie_rows + 1)
    pie1.add_data(pie_data, titles_from_data=True)
    pie1.set_categories(pie_cats)
    ws_pkd.add_chart(pie1, f"J{len(df_pkd_dist) + 4}")

    # --- Chart 3: Top 10 Keywords (bar chart) ---
    ws_kw = wb["Name Keywords Top10"]
    kw_rows = min(len(df_keywords), 10)

    bar2 = BarChart()
    bar2.type = "col"
    bar2.title = "Top 10 Keywords in Business Names"
    bar2.x_axis.title = "Keyword"
    bar2.y_axis.title = "Count"
    bar2.style = 10
    bar2.width = 20
    bar2.height = 14

    kw_data = Reference(ws_kw, min_col=2, min_row=1, max_row=kw_rows + 1)
    kw_cats = Reference(ws_kw, min_col=1, min_row=2, max_row=kw_rows + 1)
    bar2.add_data(kw_data, titles_from_data=True)
    bar2.set_categories(kw_cats)
    bar2.shape = 4
    ws_kw.add_chart(bar2, f"A{len(df_keywords) + 4}")

    # --- Chart 4: PKD Coverage pie (has PKD vs no PKD) ---
    ws_summary = wb["Summary"]
    pie2 = PieChart()
    pie2.title = "PKD Data Coverage"
    pie2.style = 10
    pie2.width = 16
    pie2.height = 12

    # Write mini table for the pie chart
    start_row = len(summary_rows) + 4
    ws_summary.cell(row=start_row, column=1, value="Category")
    ws_summary.cell(row=start_row, column=2, value="Count")
    ws_summary.cell(row=start_row + 1, column=1, value="With PKD")
    ws_summary.cell(row=start_row + 1, column=2, value=has_pkd)
    ws_summary.cell(row=start_row + 2, column=1, value="Without PKD")
    ws_summary.cell(row=start_row + 2, column=2, value=no_pkd)

    cov_data = Reference(ws_summary, min_col=2, min_row=start_row, max_row=start_row + 2)
    cov_cats = Reference(ws_summary, min_col=1, min_row=start_row + 1, max_row=start_row + 2)
    pie2.add_data(cov_data, titles_from_data=True)
    pie2.set_categories(cov_cats)
    ws_summary.add_chart(pie2, f"D2")

    wb.save(args.output)
    logger.info("Charts added.")

    # Print summary to console
    print("\n" + "=" * 60)
    print("PKD & KEYWORD INSIGHTS")
    print("=" * 60)
    print(f"\nTotal businesses: {total}")
    print(f"With PKD data: {has_pkd} ({has_pkd/total*100:.1f}%)")
    print(f"Unique PKD main codes: {len(pkd_counter)}")

    print(f"\n--- TOP 15 PKD MAIN CODES ---")
    for i, row in enumerate(pkd_dist[:15]):
        desc = row.get("description", "") or ""
        desc_str = f" - {desc}" if desc else ""
        print(f"  {i+1:2d}. {row['pkd_code']:12s}  {row['count']:4d} ({row['percentage']}%){desc_str}")

    print(f"\n--- TOP 10 KEYWORDS IN BUSINESS NAMES ---")
    for i, row in enumerate(kw_list[:10]):
        print(f"  {i+1:2d}. {row['keyword']:20s}  {row['count']:4d} ({row['percentage']}%)")

    if bigram_list:
        print(f"\n--- TOP 10 KEYWORD PAIRS ---")
        for i, row in enumerate(bigram_list[:10]):
            print(f"  {i+1:2d}. {row['keyword_pair']:30s}  {row['count']:4d} ({row['percentage']}%)")

    print(f"\n{'=' * 60}")
    print(f"Full insights saved to: {args.output}")
    print(f"Sheets: Summary, PKD Main Distribution, Name Keywords, Keywords per PKD")
    print("=" * 60)


if __name__ == "__main__":
    main()
