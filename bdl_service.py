"""
BDL (Bank Danych Lokalnych) service for revenue estimation of Polish companies.

Uses the free BDL API (https://bdl.stat.gov.pl/api/v1/) to fetch national
statistics on revenue, employment, and wages by enterprise size band.

Then combines those stats with company-level signals (legal form, capital,
directors, PKD codes, age) to produce a rough revenue estimate.
"""

import logging
import time
import math
import requests
import warnings
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

logger = logging.getLogger(__name__)

BDL_BASE = "https://bdl.stat.gov.pl/api/v1"

# ─── Simple TTL cache ───────────────────────────────────────────────────────

_cache = {}
_CACHE_TTL = 7200  # 2 hours


def _cached_get(key):
    if key in _cache:
        val, ts = _cache[key]
        if time.time() - ts < _CACHE_TTL:
            return val
        del _cache[key]
    return None


def _cached_set(key, val):
    _cache[key] = (val, time.time())


# ─── PKD division → NACE section mapping ────────────────────────────────────

PKD_TO_SECTION = {
    # A - Agriculture, forestry and fishing
    "01": "A", "02": "A", "03": "A",
    # B - Mining and quarrying
    "05": "B", "06": "B", "07": "B", "08": "B", "09": "B",
    # C - Manufacturing
    "10": "C", "11": "C", "12": "C", "13": "C", "14": "C", "15": "C",
    "16": "C", "17": "C", "18": "C", "19": "C", "20": "C", "21": "C",
    "22": "C", "23": "C", "24": "C", "25": "C", "26": "C", "27": "C",
    "28": "C", "29": "C", "30": "C", "31": "C", "32": "C", "33": "C",
    # D - Electricity, gas, steam
    "35": "D",
    # E - Water supply, sewerage, waste
    "36": "E", "37": "E", "38": "E", "39": "E",
    # F - Construction
    "41": "F", "42": "F", "43": "F",
    # G - Wholesale and retail trade
    "45": "G", "46": "G", "47": "G",
    # H - Transportation and storage
    "49": "H", "50": "H", "51": "H", "52": "H", "53": "H",
    # I - Accommodation and food service
    "55": "I", "56": "I",
    # J - Information and communication
    "58": "J", "59": "J", "60": "J", "61": "J", "62": "J", "63": "J",
    # K - Financial and insurance
    "64": "K", "65": "K", "66": "K",
    # L - Real estate
    "68": "L",
    # M - Professional, scientific, technical
    "69": "M", "70": "M", "71": "M", "72": "M", "73": "M", "74": "M", "75": "M",
    # N - Administrative and support service
    "77": "N", "78": "N", "79": "N", "80": "N", "81": "N", "82": "N",
    # O - Public administration
    "84": "O",
    # P - Education
    "85": "P",
    # Q - Human health and social work
    "86": "Q", "87": "Q", "88": "Q",
    # R - Arts, entertainment, recreation
    "90": "R", "91": "R", "92": "R", "93": "R",
    # S - Other service activities
    "94": "S", "95": "S", "96": "S",
    # T - Households as employers
    "97": "T", "98": "T",
    # U - Extraterritorial organisations
    "99": "U",
}

SECTION_NAMES = {
    "A": "Agriculture, forestry and fishing",
    "B": "Mining and quarrying",
    "C": "Manufacturing",
    "D": "Electricity, gas, steam",
    "E": "Water supply, sewerage, waste",
    "F": "Construction",
    "G": "Wholesale and retail trade",
    "H": "Transportation and storage",
    "I": "Accommodation and food service",
    "J": "Information and communication",
    "K": "Financial and insurance",
    "L": "Real estate",
    "M": "Professional, scientific, technical",
    "N": "Administrative and support service",
    "O": "Public administration",
    "P": "Education",
    "Q": "Human health and social work",
    "R": "Arts, entertainment, recreation",
    "S": "Other service activities",
    "T": "Households as employers",
    "U": "Extraterritorial organisations",
}


def pkd_to_section(pkd_code):
    """Map a PKD code like '19.20.Z' to its NACE section letter."""
    if not pkd_code:
        return None
    division = pkd_code.strip().split(".")[0].zfill(2)
    return PKD_TO_SECTION.get(division)


def pkd_to_division(pkd_code):
    """Extract division number from PKD code (e.g. '19.20.Z' -> '19')."""
    if not pkd_code:
        return None
    return pkd_code.strip().split(".")[0].zfill(2)


# ─── BDL API calls ──────────────────────────────────────────────────────────

def _bdl_fetch_variable(variable_id, year_from=2015, year_to=2026):
    """Fetch a single BDL variable (national level, unit-level=0).
    Returns list of {year, val} dicts for the latest result, or [].
    """
    cache_key = f"bdl_{variable_id}_{year_from}_{year_to}"
    cached = _cached_get(cache_key)
    if cached is not None:
        return cached

    url = f"{BDL_BASE}/data/by-variable/{variable_id}"
    params = {
        "format": "json",
        "unit-level": 0,
        "year-from": year_from,
        "year-to": year_to,
    }

    try:
        resp = requests.get(url, params=params, timeout=30, verify=False)
        resp.raise_for_status()
        data = resp.json()
        results = data.get("results", [])
        if results:
            values = results[0].get("values", [])
            _cached_set(cache_key, values)
            return values
    except Exception as e:
        logger.error(f"BDL API error for variable {variable_id}: {e}")

    _cached_set(cache_key, [])
    return []


def _latest_value(values):
    """Get the most recent non-null value from a BDL values list."""
    for v in sorted(values, key=lambda x: x.get("year", 0), reverse=True):
        val = v.get("val")
        if val is not None:
            return val, v.get("year")
    return None, None


def fetch_revenue_per_size_band():
    """Fetch revenue per entity for micro/small/medium/large (mln PLN/year).

    BDL variable IDs (P2895 group - national averages across all sectors):
      196213 = micro, 498839 = small, 196207 = medium, 196201 = large
    """
    var_ids = {
        "micro": 196213,
        "small": 498839,
        "medium": 196207,
        "large": 196201,
    }
    result = {}
    for band, vid in var_ids.items():
        values = _bdl_fetch_variable(vid)
        val, year = _latest_value(values)
        result[band] = {"value_mln_pln": val, "year": year, "all_years": values}
        if val is not None:
            logger.info(f"  Revenue per {band} entity: {val:.1f} mln PLN ({year})")
    return result


def fetch_workers_per_enterprise():
    """Fetch workers per enterprise by size band for section C (Manufacturing).

    BDL variable IDs (P2895 group):
      196230 = micro, 196236 = small, 196238 = medium, 196244 = large
    """
    var_ids = {
        "micro": 196230,
        "small": 196236,
        "medium": 196238,
        "large": 196244,
    }
    result = {}
    for band, vid in var_ids.items():
        values = _bdl_fetch_variable(vid)
        val, year = _latest_value(values)
        result[band] = {"value": val, "year": year}
        if val is not None:
            logger.info(f"  Workers per {band} enterprise: {val:.1f} ({year})")
    return result


def fetch_sector_totals(section_letter="C"):
    """Fetch total sector revenue (tys PLN) and enterprise count.

    BDL variable IDs:
      213799 = total revenue in tys PLN
      215187 = enterprise count
    """
    rev_values = _bdl_fetch_variable(213799)
    count_values = _bdl_fetch_variable(215187)

    rev_val, rev_year = _latest_value(rev_values)
    count_val, count_year = _latest_value(count_values)

    result = {
        "total_revenue_tys_pln": rev_val,
        "revenue_year": rev_year,
        "enterprise_count": count_val,
        "count_year": count_year,
    }
    if rev_val and count_val:
        result["avg_revenue_mln_pln"] = (rev_val / 1000) / count_val
        logger.info(f"  Sector avg revenue: {result['avg_revenue_mln_pln']:.2f} mln PLN "
                     f"({rev_val/1e6:.0f} bln / {count_val:.0f} enterprises)")
    return result


def fetch_wages_per_size_band():
    """Fetch monthly gross wage by size band (zł).

    BDL variable IDs:
      196217 = micro, 498835 = small, 196211 = medium, 196205 = large
    """
    var_ids = {
        "micro": 196217,
        "small": 498835,
        "medium": 196211,
        "large": 196205,
    }
    result = {}
    for band, vid in var_ids.items():
        values = _bdl_fetch_variable(vid)
        val, year = _latest_value(values)
        result[band] = {"value_zl": val, "year": year}
    return result


def fetch_division_employment(division, section_letter="C"):
    """Fetch division employment and section total employment for division ratio.

    BDL variable IDs:
      155614 = division-level employment
      155654 = section-level total employment
    """
    div_values = _bdl_fetch_variable(155614)
    sec_values = _bdl_fetch_variable(155654)

    div_val, div_year = _latest_value(div_values)
    sec_val, sec_year = _latest_value(sec_values)

    result = {
        "division_employment": div_val,
        "section_employment": sec_val,
        "division_share": None,
    }
    if div_val and sec_val and sec_val > 0:
        result["division_share"] = div_val / sec_val
        logger.info(f"  Division {division} share of section {section_letter}: "
                     f"{result['division_share']*100:.2f}%")
    return result


# ─── Size band scoring ──────────────────────────────────────────────────────

SIZE_BAND_THRESHOLDS = {
    "micro": (0, 2),
    "small": (3, 5),
    "medium": (6, 8),
    "large": (9, 999),
}

# Confidence spreads per size band (log-normal distribution shape)
CONFIDENCE_SPREADS = {
    "micro":  (0.30, 2.50),
    "small":  (0.40, 2.00),
    "medium": (0.45, 1.90),
    "large":  (0.50, 1.80),
}

# Median share capital per band (PLN) — used for capital scaling
MEDIAN_CAPITAL = {
    "micro":  5_000,
    "small":  50_000,
    "medium": 500_000,
    "large":  5_000_000,
}


def score_company_size(company_info):
    """Score a company to determine its size band.

    company_info dict may contain:
      legal_form, has_supervisory_board, director_count,
      pkd_count, share_capital_pln, company_age_years,
      user_employee_count
    """
    score = 0

    legal_form = str(company_info.get("legal_form", "")).upper()
    if "AKCYJNA" in legal_form or "S.A." in legal_form:
        score += 3
    elif "Z O.O." in legal_form or "Z.O.O." in legal_form:
        score += 2
    elif "JAWNA" in legal_form or "KOMANDYTOWA" in legal_form:
        score += 1

    if company_info.get("has_supervisory_board"):
        score += 2

    directors = company_info.get("director_count", 0) or 0
    if directors >= 7:
        score += 3
    elif directors >= 4:
        score += 2
    elif directors >= 2:
        score += 1

    pkd_count = company_info.get("pkd_count", 0) or 0
    if pkd_count >= 10:
        score += 2
    elif pkd_count >= 5:
        score += 1

    capital = company_info.get("share_capital_pln", 0) or 0
    if capital >= 1_000_000:
        score += 2
    elif capital >= 100_000:
        score += 1

    # Determine band
    for band, (low, high) in SIZE_BAND_THRESHOLDS.items():
        if low <= score <= high:
            return band, score

    return "micro", score


# ─── Age adjustment ─────────────────────────────────────────────────────────

AGE_ADJUSTMENTS = [
    (2, 0.40),
    (5, 0.65),
    (10, 0.85),
    (20, 1.00),
    (50, 1.10),
    (999, 1.15),
]


def age_multiplier(company_age_years):
    """Return age-based revenue adjustment multiplier."""
    if company_age_years is None or company_age_years < 0:
        return 1.0
    for threshold, mult in AGE_ADJUSTMENTS:
        if company_age_years < threshold:
            return mult
    return 1.15


# ─── Main estimation ────────────────────────────────────────────────────────

def get_rich_estimate(pkd_code, company_info=None):
    """
    Full revenue estimation pipeline.

    Args:
        pkd_code: e.g. "19.20.Z"
        company_info: dict with keys like legal_form, share_capital_pln,
                      director_count, pkd_count, company_age_years,
                      has_supervisory_board, user_employee_count

    Returns dict with:
        size_band, size_score, base_revenue_mln,
        adjusted_revenue_mln, low_estimate_mln, high_estimate_mln,
        sector_avg_mln, employee_estimate_mln (if employees given),
        division_share, adjustments_applied, data_year, section, section_name
    """
    if company_info is None:
        company_info = {}

    section = pkd_to_section(pkd_code)
    division = pkd_to_division(pkd_code)
    section_name = SECTION_NAMES.get(section, "Unknown")

    logger.info(f"Estimating revenue for PKD {pkd_code} (section {section} - {section_name})")

    # Step 1: Fetch BDL data
    logger.info("Fetching BDL data...")
    revenue_bands = fetch_revenue_per_size_band()
    workers_bands = fetch_workers_per_enterprise()
    sector_totals = fetch_sector_totals(section)
    wages_bands = fetch_wages_per_size_band()
    div_employment = fetch_division_employment(division, section) if division else {}

    # Step 2: Score company size
    size_band, size_score = score_company_size(company_info)
    logger.info(f"Size band: {size_band} (score={size_score})")

    # Step 3: Base estimate from BDL
    band_data = revenue_bands.get(size_band, {})
    base_revenue = band_data.get("value_mln_pln")
    data_year = band_data.get("year")

    if base_revenue is None:
        logger.warning(f"No BDL revenue data for {size_band} band")
        base_revenue = 0

    # Step 4: Adjustments
    adjustments = []
    adjusted = base_revenue

    # A) Share capital scaling
    capital = company_info.get("share_capital_pln", 0) or 0
    median_cap = MEDIAN_CAPITAL.get(size_band, 5_000_000)
    capital_multiplier = 1.0
    if capital > 0 and median_cap > 0:
        ratio = capital / median_cap
        if ratio > 1:
            raw_mult = ratio ** 0.3
            capital_multiplier = min(raw_mult, 3.0)
            adjustments.append(f"capital_scaling: ratio={ratio:.1f}, mult={capital_multiplier:.2f}")
    adjusted *= capital_multiplier

    # B) Age adjustment
    age_years = company_info.get("company_age_years")
    age_mult = age_multiplier(age_years)
    if age_mult != 1.0:
        adjustments.append(f"age_adj: {age_years:.0f}y -> x{age_mult:.2f}")
    adjusted *= age_mult

    # C) Market trends adjustment (inflation + sector trend + GDP)
    from market_trends import get_market_adjustment
    market_mult, market_details = get_market_adjustment(section, data_year)
    if market_mult != 1.0:
        adjustments.append(
            f"market: inflation x{market_details['inflation_mult']:.3f}, "
            f"sector x{market_details['sector_trend_mult']:.3f}, "
            f"GDP x{market_details['gdp_mult']:.3f} = x{market_mult:.4f}"
        )
    adjusted *= market_mult

    # Step 5: Confidence range
    low_spread, high_spread = CONFIDENCE_SPREADS.get(size_band, (0.50, 1.80))
    low_est = adjusted * low_spread
    high_est = adjusted * high_spread

    # Step 6: Sector average (simple estimate)
    sector_avg = sector_totals.get("avg_revenue_mln_pln")

    # Step 7: Employee-based estimate (if employee count provided)
    employee_est = None
    emp_count = company_info.get("user_employee_count")
    if emp_count and emp_count > 0:
        workers_data = workers_bands.get(size_band, {})
        workers_per_ent = workers_data.get("value")
        ent_count = sector_totals.get("enterprise_count")
        total_rev = sector_totals.get("total_revenue_tys_pln")
        if workers_per_ent and ent_count and total_rev and workers_per_ent > 0:
            rev_per_employee = (total_rev / 1000) / (workers_per_ent * ent_count)
            employee_est = emp_count * rev_per_employee
            adjustments.append(f"employee_est: {emp_count} emp x {rev_per_employee:.3f} mln = {employee_est:.1f} mln")

    # Step 8: Division ratio
    div_share = div_employment.get("division_share")

    # Wages for context
    wage_data = wages_bands.get(size_band, {})

    result = {
        "pkd_code": pkd_code,
        "section": section,
        "section_name": section_name,
        "division": division,
        "size_band": size_band,
        "size_score": size_score,
        "base_revenue_mln": round(base_revenue, 2) if base_revenue else None,
        "adjusted_revenue_mln": round(adjusted, 2) if adjusted else None,
        "low_estimate_mln": round(low_est, 2) if low_est else None,
        "high_estimate_mln": round(high_est, 2) if high_est else None,
        "sector_avg_mln": round(sector_avg, 2) if sector_avg else None,
        "employee_estimate_mln": round(employee_est, 2) if employee_est else None,
        "division_share": round(div_share * 100, 2) if div_share else None,
        "capital_multiplier": round(capital_multiplier, 2),
        "age_multiplier": round(age_mult, 2),
        "market_multiplier": round(market_mult, 4),
        "inflation_multiplier": round(market_details.get("inflation_mult", 1), 3),
        "sector_trend_multiplier": round(market_details.get("sector_trend_mult", 1), 3),
        "gdp_multiplier": round(market_details.get("gdp_mult", 1), 3),
        "sector_growth_pct": market_details.get("sector_growth_pct"),
        "gdp_growth_pct": market_details.get("gdp_growth_pct"),
        "eur_pln_rate": market_details.get("eur_pln_now"),
        "eur_pln_change_pct": market_details.get("eur_pln_change_pct"),
        "adjustments_applied": "; ".join(adjustments) if adjustments else "none",
        "data_year": data_year,
        "monthly_wage_zl": wage_data.get("value_zl"),
        "confidence_low_spread": low_spread,
        "confidence_high_spread": high_spread,
    }

    logger.info(f"  Base: {base_revenue} mln -> Adjusted: {adjusted:.1f} mln "
                f"(range: {low_est:.1f} - {high_est:.1f} mln)")

    return result


def estimate_simple(pkd_code):
    """Quick estimate without company details — just sector average."""
    section = pkd_to_section(pkd_code)
    sector = fetch_sector_totals(section)
    avg = sector.get("avg_revenue_mln_pln")
    return {
        "pkd_code": pkd_code,
        "section": section,
        "section_name": SECTION_NAMES.get(section, "Unknown"),
        "sector_avg_mln": round(avg, 2) if avg else None,
        "enterprise_count": sector.get("enterprise_count"),
    }
