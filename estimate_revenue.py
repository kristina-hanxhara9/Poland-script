#!/usr/bin/env python3
"""
Estimate revenue/turnover for Polish companies using BDL national statistics.

Reads an Excel file with company data (output from lookup scripts) and
estimates turnover based on:
  - PKD code -> NACE section mapping
  - Company size scoring (legal form, capital, directors, PKD count, age)
  - BDL national revenue-per-size-band statistics
  - Share capital scaling, age adjustment
  - Sector average as secondary signal

Input Excel expected columns (auto-detected):
  - pkd_main (PKD code like "47.11.Z")
  - api_name or name
  - date_started (for age calculation)
  - nip, regon, krs (identifiers)
  - Optional: share_capital, employee_count

Usage:
    python estimate_revenue.py results.xlsx --output revenue_estimates.xlsx
    python estimate_revenue.py results.xlsx --limit 5
    python estimate_revenue.py results.xlsx --pkd-col Q --name-col B
"""

import argparse
import logging
import sys
from datetime import datetime

from dotenv import load_dotenv
import pandas as pd

load_dotenv()

from bdl_service import (
    get_rich_estimate,
    pkd_to_section,
    SECTION_NAMES,
    score_company_size,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)


def find_column(df, candidates, required=False, label=""):
    """Find first matching column name from candidates list."""
    for c in candidates:
        # Exact match
        if c in df.columns:
            return c
        # Case-insensitive
        for col in df.columns:
            if col.lower().strip() == c.lower().strip():
                return col
    if required:
        logger.error(f"Could not find {label} column. Tried: {candidates}")
        logger.error(f"Available columns: {list(df.columns)}")
        sys.exit(1)
    return None


def parse_date(val):
    """Try to parse a date string into a datetime."""
    if pd.isna(val) or not val:
        return None
    val = str(val).strip()
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(val, fmt)
        except ValueError:
            continue
    # Try pandas
    try:
        return pd.to_datetime(val)
    except Exception:
        return None


def guess_legal_form(name):
    """Guess legal form from company name."""
    if not name:
        return ""
    name_upper = str(name).upper()
    if "S.A." in name_upper or "AKCYJNA" in name_upper:
        return "SPÓŁKA AKCYJNA"
    elif "SP. Z O.O." in name_upper or "Z O.O." in name_upper or "Z.O.O." in name_upper:
        return "SPÓŁKA Z O.O."
    elif "SP. J." in name_upper or "JAWNA" in name_upper:
        return "SPÓŁKA JAWNA"
    elif "SP. K." in name_upper or "KOMANDYTOWA" in name_upper:
        return "SPÓŁKA KOMANDYTOWA"
    elif "S.C." in name_upper or "CYWILNA" in name_upper:
        return "SPÓŁKA CYWILNA"
    return ""


def parse_capital(val):
    """Parse share capital value — handle strings like '1,000,000.00 PLN'."""
    if pd.isna(val) or not val:
        return 0
    s = str(val).strip().upper().replace("PLN", "").replace("ZŁ", "").replace("ZL", "")
    s = s.replace(",", "").replace(" ", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0


def main():
    parser = argparse.ArgumentParser(
        description="Estimate revenue for Polish companies using BDL statistics."
    )
    parser.add_argument("input_file", help="Excel file with company data")
    parser.add_argument("--output", "-o", default="revenue_estimates.xlsx",
                        help="Output Excel file (default: revenue_estimates.xlsx)")
    parser.add_argument("--limit", "-n", type=int, default=0,
                        help="Only process first N rows (0 = all)")
    parser.add_argument("--pkd-col", help="Column name/letter for PKD main code")
    parser.add_argument("--name-col", help="Column name/letter for company name")
    parser.add_argument("--employees", type=int, default=0,
                        help="Default employee count to use for all companies (0 = skip)")
    args = parser.parse_args()

    logger.info(f"Reading {args.input_file}")
    try:
        df = pd.read_excel(args.input_file, header=0)
    except FileNotFoundError:
        logger.error(f"File not found: {args.input_file}")
        sys.exit(1)

    logger.info(f"Loaded {len(df)} rows, {len(df.columns)} columns")
    logger.info(f"Columns: {list(df.columns)}")

    # Find relevant columns
    if args.pkd_col:
        pkd_col = args.pkd_col
    else:
        pkd_col = find_column(df, ["pkd_main", "PKD_main", "pkd", "PKD", "main_pkd"],
                              required=True, label="PKD main")

    if args.name_col:
        name_col = args.name_col
    else:
        name_col = find_column(df, ["api_name", "name", "Name", "company_name", "input_name"],
                               required=True, label="Company name")

    date_col = find_column(df, ["date_started", "date_started", "registration_date",
                                "start_date", "founded"])
    capital_col = find_column(df, ["share_capital", "share_capital_pln", "capital",
                                   "kapital", "kapital_zakladowy"])
    krs_col = find_column(df, ["krs", "KRS"])
    nip_col = find_column(df, ["nip", "NIP", "input_nip"])
    regon_col = find_column(df, ["regon", "REGON"])
    pkd_codes_col = find_column(df, ["pkd_codes", "PKD_codes", "all_pkd"])
    status_col = find_column(df, ["status", "match_status"])
    zip_col = find_column(df, ["api_zip_code", "zip_code", "input_zip"])
    city_col = find_column(df, ["api_city", "city"])
    emp_col = find_column(df, ["employee_count", "employees", "num_employees"])

    logger.info(f"Using columns: PKD={pkd_col}, Name={name_col}, Date={date_col}, Capital={capital_col}")

    # Process each row
    results = []
    now = datetime.now()

    rows_to_process = df if args.limit == 0 else df.head(args.limit)

    for idx, row in rows_to_process.iterrows():
        pkd_code = str(row.get(pkd_col, "")).strip()
        company_name = str(row.get(name_col, "")).strip()

        if pkd_code.lower() in ("nan", "none", ""):
            logger.info(f"[{idx+1}] {company_name} — no PKD code, skipping estimation")
            results.append({
                "row": idx + 1,
                "company_name": company_name,
                "pkd_code": "",
                "estimate_status": "NO PKD CODE",
            })
            continue

        logger.info(f"[{idx+1}] {company_name} — PKD: {pkd_code}")

        # Build company_info dict for scoring
        company_info = {}

        # Legal form — guess from name
        company_info["legal_form"] = guess_legal_form(company_name)

        # Company age
        if date_col:
            started = parse_date(row.get(date_col))
            if started:
                company_info["company_age_years"] = (now - started).days / 365.25

        # Share capital
        if capital_col:
            company_info["share_capital_pln"] = parse_capital(row.get(capital_col))

        # PKD count
        if pkd_codes_col:
            codes_str = str(row.get(pkd_codes_col, ""))
            if codes_str.lower() not in ("nan", "none", ""):
                codes = [c.strip() for c in codes_str.split(";") if c.strip()]
                company_info["pkd_count"] = len(codes)
            else:
                company_info["pkd_count"] = 1
        else:
            company_info["pkd_count"] = 1

        # KRS presence suggests larger company
        if krs_col:
            krs_val = str(row.get(krs_col, "")).strip()
            if krs_val.lower() not in ("nan", "none", "") and len(krs_val) >= 5:
                company_info["has_supervisory_board"] = False  # conservative default
                company_info["director_count"] = 1  # conservative default

        # Employee count
        if emp_col:
            try:
                emp_val = int(row.get(emp_col, 0))
                if emp_val > 0:
                    company_info["user_employee_count"] = emp_val
            except (ValueError, TypeError):
                pass
        if args.employees > 0 and "user_employee_count" not in company_info:
            company_info["user_employee_count"] = args.employees

        # Get the estimate
        try:
            est = get_rich_estimate(pkd_code, company_info)
        except Exception as e:
            logger.error(f"  Estimation failed: {e}")
            results.append({
                "row": idx + 1,
                "company_name": company_name,
                "pkd_code": pkd_code,
                "estimate_status": f"ERROR: {e}",
            })
            continue

        # Build output row — combine original data with estimate
        out = {
            "row": idx + 1,
            "company_name": company_name,
            "nip": str(row.get(nip_col, "")).strip() if nip_col else "",
            "regon": str(row.get(regon_col, "")).strip() if regon_col else "",
            "krs": str(row.get(krs_col, "")).strip() if krs_col else "",
            "status": str(row.get(status_col, "")).strip() if status_col else "",
            "city": str(row.get(city_col, "")).strip() if city_col else "",
            "zip_code": str(row.get(zip_col, "")).strip() if zip_col else "",
            "pkd_code": pkd_code,
            "section": est.get("section", ""),
            "section_name": est.get("section_name", ""),
            "size_band": est.get("size_band", ""),
            "size_score": est.get("size_score", ""),
            "base_revenue_mln_pln": est.get("base_revenue_mln", ""),
            "adjusted_revenue_mln_pln": est.get("adjusted_revenue_mln", ""),
            "low_estimate_mln_pln": est.get("low_estimate_mln", ""),
            "high_estimate_mln_pln": est.get("high_estimate_mln", ""),
            "sector_avg_revenue_mln_pln": est.get("sector_avg_mln", ""),
            "employee_estimate_mln_pln": est.get("employee_estimate_mln", ""),
            "division_share_pct": est.get("division_share", ""),
            "capital_multiplier": est.get("capital_multiplier", ""),
            "age_multiplier": est.get("age_multiplier", ""),
            "adjustments": est.get("adjustments_applied", ""),
            "data_year": est.get("data_year", ""),
            "monthly_wage_zl": est.get("monthly_wage_zl", ""),
            "company_age_years": round(company_info.get("company_age_years", 0), 1)
                if company_info.get("company_age_years") else "",
            "legal_form_guessed": company_info.get("legal_form", ""),
            "estimate_status": "OK",
        }
        results.append(out)

        adj = est.get("adjusted_revenue_mln")
        low = est.get("low_estimate_mln")
        high = est.get("high_estimate_mln")
        logger.info(f"  -> {est['size_band'].upper()} | "
                     f"Est: {adj:.1f} mln PLN (range: {low:.1f} - {high:.1f})")

    # Save to Excel
    df_out = pd.DataFrame(results).fillna("")
    logger.info(f"Saving {len(df_out)} rows to {args.output}")

    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False, sheet_name="Revenue Estimates")

        # Add a summary sheet
        summary_rows = []
        ok_rows = [r for r in results if r.get("estimate_status") == "OK"]
        if ok_rows:
            for band in ["micro", "small", "medium", "large"]:
                band_rows = [r for r in ok_rows if r.get("size_band") == band]
                if band_rows:
                    revenues = [r["adjusted_revenue_mln_pln"] for r in band_rows
                                if r.get("adjusted_revenue_mln_pln")]
                    if revenues:
                        summary_rows.append({
                            "size_band": band,
                            "count": len(band_rows),
                            "avg_est_mln_pln": round(sum(revenues) / len(revenues), 2),
                            "min_est_mln_pln": round(min(revenues), 2),
                            "max_est_mln_pln": round(max(revenues), 2),
                            "total_est_mln_pln": round(sum(revenues), 2),
                        })

            # By section
            sections = {}
            for r in ok_rows:
                sec = r.get("section", "?")
                if sec not in sections:
                    sections[sec] = []
                sections[sec].append(r)

            section_summary = []
            for sec, sec_rows in sorted(sections.items()):
                revenues = [r["adjusted_revenue_mln_pln"] for r in sec_rows
                            if r.get("adjusted_revenue_mln_pln")]
                if revenues:
                    section_summary.append({
                        "section": sec,
                        "section_name": SECTION_NAMES.get(sec, "Unknown"),
                        "count": len(sec_rows),
                        "avg_est_mln_pln": round(sum(revenues) / len(revenues), 2),
                        "total_est_mln_pln": round(sum(revenues), 2),
                    })

            if summary_rows:
                pd.DataFrame(summary_rows).to_excel(writer, index=False,
                                                     sheet_name="Summary by Size")
            if section_summary:
                pd.DataFrame(section_summary).to_excel(writer, index=False,
                                                        sheet_name="Summary by Section")

    logger.info("=" * 55)
    logger.info(f"DONE | {len(ok_rows)}/{len(results)} companies estimated")
    logger.info(f"Output: {args.output}")
    if ok_rows:
        all_rev = [r["adjusted_revenue_mln_pln"] for r in ok_rows
                   if r.get("adjusted_revenue_mln_pln")]
        if all_rev:
            logger.info(f"Total estimated revenue: {sum(all_rev):.1f} mln PLN")
            logger.info(f"Average: {sum(all_rev)/len(all_rev):.1f} mln PLN")


if __name__ == "__main__":
    main()
