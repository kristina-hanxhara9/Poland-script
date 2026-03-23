#!/usr/bin/env python3
"""
Poland Business PKD Code Analysis

Main entry point. Reads an Excel file with business data (name, zip, channel columns),
looks up ALL businesses via Poland government APIs (CEIDG, KRS, GUS),
extracts full details (address, status, PKD codes, etc.), and analyzes
which PKD/SIC codes correspond to DIY shops, Paint specialists, Builders merchants.

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
        description="Analyze Polish businesses: look up PKD/SIC codes and classify by channel category."
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

    # Load Excel (ALL rows, not filtered)
    logger.info(f"Loading input file: {args.input_file}")
    try:
        df, channel_cols = load_excel(args.input_file, args.channel_column)
    except FileNotFoundError:
        logger.error(f"File not found: {args.input_file}")
        sys.exit(1)
    except ValueError as e:
        logger.error(str(e))
        sys.exit(1)

    logger.info(f"Total rows to process: {len(df)}")

    if df.empty:
        logger.warning("No data found in the input file.")
        sys.exit(0)

    # Set up API client
    api_client = PolandAPIClient(
        ceidg_api_key=args.ceidg_key,
        gus_api_key=args.gus_key,
        use_sandbox=args.sandbox,
    )

    if args.ceidg_key or api_client.ceidg_api_key:
        logger.info("CEIDG API key configured.")
    if args.gus_key or api_client.gus_api_key:
        logger.info("GUS API key configured.")
    logger.info("KRS API: always available (free, no auth needed).")

    # Run analysis on ALL businesses
    enriched_df, report = run_full_analysis(df, api_client)

    # Save output Excel
    logger.info(f"Saving enriched data to: {args.output}")

    output_df = enriched_df.copy()
    # Convert lists to strings for Excel compatibility
    for col in output_df.columns:
        if output_df[col].apply(lambda x: isinstance(x, list)).any():
            output_df[col] = output_df[col].apply(
                lambda x: "; ".join(str(i) for i in x) if isinstance(x, list) else str(x)
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
