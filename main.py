#!/usr/bin/env python3
"""
Poland Business PKD Code Analysis

Main entry point. Reads an Excel file with business data (name, zip, channel),
looks up PKD codes via Poland government APIs, and analyzes which PKD codes
correspond to each channel category (DIY shops, Paint specialists, Builders merchants).

Usage:
    python main.py input.xlsx [--output results.xlsx] [--ceidg-key KEY] [--gus-key KEY] [--sandbox]
"""

import argparse
import logging
import sys

import pandas as pd

from poland_api import PolandAPIClient
from analysis import load_excel, run_full_analysis

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)


def main():
    parser = argparse.ArgumentParser(
        description="Analyze Polish businesses and their PKD codes by channel category."
    )
    parser.add_argument("input_file", help="Path to input Excel file (.xlsx)")
    parser.add_argument(
        "--output", "-o",
        default="results.xlsx",
        help="Path for output Excel file (default: results.xlsx)",
    )
    parser.add_argument(
        "--report", "-r",
        default="report.txt",
        help="Path for text report (default: report.txt)",
    )
    parser.add_argument(
        "--ceidg-key",
        help="CEIDG API key (or set CEIDG_API_KEY env var)",
    )
    parser.add_argument(
        "--gus-key",
        help="GUS REGON API key (or set GUS_API_KEY env var)",
    )
    parser.add_argument(
        "--sandbox",
        action="store_true",
        help="Use GUS sandbox API (test data, no real key needed)",
    )
    parser.add_argument(
        "--channel-column",
        default="channel",
        help="Name of the channel/category column in the Excel file (default: channel)",
    )
    args = parser.parse_args()

    # Load Excel
    logger.info(f"Loading input file: {args.input_file}")
    try:
        full_df, filtered_df, channel_col = load_excel(args.input_file, args.channel_column)
    except FileNotFoundError:
        logger.error(f"File not found: {args.input_file}")
        sys.exit(1)
    except ValueError as e:
        logger.error(str(e))
        sys.exit(1)

    logger.info(f"Total rows: {len(full_df)}, Filtered (target categories): {len(filtered_df)}")

    if filtered_df.empty:
        logger.warning("No rows match target categories (DIY shops, Paint specialists, Builders merchants).")
        logger.info(f"Available channel values: {full_df[channel_col].unique().tolist()}")
        sys.exit(0)

    # Set up API client
    api_client = PolandAPIClient(
        ceidg_api_key=args.ceidg_key,
        gus_api_key=args.gus_key,
        use_sandbox=args.sandbox,
    )

    if api_client.has_api_access():
        logger.info("API access configured. Will look up PKD codes from government databases.")
    else:
        logger.info(
            "No API keys configured. Running keyword-based analysis only.\n"
            "To enable API lookups, set CEIDG_API_KEY or GUS_API_KEY environment variables,\n"
            "or use --ceidg-key / --gus-key / --sandbox flags."
        )

    # Run analysis
    enriched_df, report = run_full_analysis(filtered_df, api_client)

    # Save output Excel
    logger.info(f"Saving enriched data to: {args.output}")

    # Prepare output columns - convert lists to strings for Excel
    output_df = enriched_df.copy()
    output_df["pkd_codes"] = output_df["pkd_codes"].apply(
        lambda x: "; ".join(x) if isinstance(x, list) else str(x)
    )
    output_df["pkd_descriptions"] = output_df["pkd_descriptions"].apply(
        lambda x: "; ".join(x) if isinstance(x, list) else str(x)
    )
    output_df.to_excel(args.output, index=False)

    # Save text report
    logger.info(f"Saving analysis report to: {args.report}")
    with open(args.report, "w", encoding="utf-8") as f:
        f.write(report)

    # Print report to console
    print("\n" + report)

    logger.info("Done!")


if __name__ == "__main__":
    main()
