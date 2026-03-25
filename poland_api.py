"""
Poland Government API client for looking up businesses and retrieving full details.

Supports:
- CEIDG API v2 (dane.biznes.gov.pl) - for sole proprietors
- KRS API (api-krs.ms.gov.pl) - for registered companies (free, no auth)
- GUS REGON/BIR1 API via gusregon library - for all entities (sandbox available)
"""

import os
import time
import logging
import requests

logger = logging.getLogger(__name__)

# API endpoints
CEIDG_API_BASE = "https://dane.biznes.gov.pl/api/ceidg/v2"
KRS_API_BASE = "https://api-krs.ms.gov.pl/api/krs"
GUS_SANDBOX_KEY = "abcde12345abcde12345"

# Rate limiting
CEIDG_RATE_LIMIT_DELAY = 3.6
KRS_RATE_LIMIT_DELAY = 9.0


def _build_address_dict(addr_obj):
    """Build a normalized address dict from a CEIDG address object."""
    if not addr_obj or not isinstance(addr_obj, dict):
        return {}
    return {
        "street": addr_obj.get("ulica", ""),
        "building": addr_obj.get("budynek", ""),
        "unit": addr_obj.get("lokal", ""),
        "zip_code": addr_obj.get("kod", ""),
        "city": addr_obj.get("miasto", ""),
        "municipality": addr_obj.get("gmina", ""),
        "county": addr_obj.get("powiat", ""),
        "province": addr_obj.get("wojewodztwo", ""),
        "country": addr_obj.get("kraj", ""),
    }


def _format_address(addr):
    """Format an address dict into a single-line string."""
    if not addr:
        return ""
    parts = []
    street = addr.get("street", "")
    building = addr.get("building", "")
    unit = addr.get("unit", "")
    if street:
        s = street
        if building:
            s += f" {building}"
        if unit:
            s += f"/{unit}"
        parts.append(s)
    zip_code = addr.get("zip_code", "")
    city = addr.get("city", "")
    if zip_code and city:
        parts.append(f"{zip_code} {city}")
    elif city:
        parts.append(city)
    province = addr.get("province", "")
    if province:
        parts.append(f"woj. {province}")
    return ", ".join(parts)


class PolandAPIClient:
    """Client for Poland government business APIs."""

    def __init__(self, ceidg_api_key=None, gus_api_key=None, use_sandbox=False):
        self.ceidg_api_key = (ceidg_api_key or os.environ.get("CEIDG_API_KEY", "")).strip().strip('"').strip("'")
        self.gus_api_key = (gus_api_key or os.environ.get("GUS_API_KEY", "")).strip().strip('"').strip("'")
        self.use_sandbox = use_sandbox
        self._last_ceidg_request = 0
        self._last_krs_request = 0
        self._gus_client = None

        if use_sandbox and not self.gus_api_key:
            self.gus_api_key = GUS_SANDBOX_KEY

    def _rate_limit_ceidg(self):
        elapsed = time.time() - self._last_ceidg_request
        if elapsed < CEIDG_RATE_LIMIT_DELAY:
            time.sleep(CEIDG_RATE_LIMIT_DELAY - elapsed)
        self._last_ceidg_request = time.time()

    def _rate_limit_krs(self):
        elapsed = time.time() - self._last_krs_request
        if elapsed < KRS_RATE_LIMIT_DELAY:
            time.sleep(KRS_RATE_LIMIT_DELAY - elapsed)
        self._last_krs_request = time.time()

    def has_api_access(self):
        """Check if any API credentials are configured. KRS is always available."""
        return True

    # ---- CEIDG API v2 ----

    def search_ceidg(self, name, zip_code=None, city=None):
        """Search CEIDG for sole proprietors by name and optional location."""
        if not self.ceidg_api_key:
            logger.warning("No CEIDG API key configured. Set CEIDG_API_KEY env var.")
            return []

        self._rate_limit_ceidg()

        headers = {
            "Authorization": f"Bearer {self.ceidg_api_key}",
            "Accept": "application/json",
        }

        params = {"nazwa": name}
        if zip_code:
            params["kod"] = zip_code
        if city:
            params["miasto"] = city

        try:
            response = requests.get(
                f"{CEIDG_API_BASE}/firmy",
                headers=headers,
                params=params,
                timeout=30,
            )
            response.raise_for_status()
            data = response.json()

            results = []
            firms = data.get("firmy", data) if isinstance(data, dict) else data
            if not isinstance(firms, list):
                firms = [firms] if firms else []

            for firm in firms:
                addr = firm.get("adresDzialalnosci", {})
                address = _build_address_dict(addr)
                owner = firm.get("wlasciciel", {}) or {}

                result = {
                    "name": firm.get("nazwa", ""),
                    "nip": owner.get("nip", "") or firm.get("nip", ""),
                    "regon": owner.get("regon", "") or firm.get("regon", ""),
                    "krs": "",
                    "address": address,
                    "address_str": _format_address(address),
                    "street": address.get("street", ""),
                    "building": address.get("building", ""),
                    "unit": address.get("unit", ""),
                    "zip_code": address.get("zip_code", ""),
                    "city": address.get("city", ""),
                    "municipality": address.get("municipality", ""),
                    "county": address.get("county", ""),
                    "province": address.get("province", ""),
                    "country": address.get("country", ""),
                    "mailing_address_str": _format_address(
                        _build_address_dict(firm.get("adresKorespondencyjny", {}))
                    ),
                    "status": firm.get("status", ""),
                    "is_active": firm.get("status", "") == "AKTYWNY",
                    "date_started": firm.get("dataRozpoczecia", ""),
                    "date_ended": firm.get("dataZakonczenia", ""),
                    "date_suspended": firm.get("dataZawieszenia", ""),
                    "date_resumed": firm.get("dataWznowienia", ""),
                    "date_deleted": firm.get("dataWykreslenia", ""),
                    "pkd_codes": [],
                    "pkd_main": firm.get("pkdGlowny", ""),
                    "owner_first_name": owner.get("imie", ""),
                    "owner_last_name": owner.get("nazwisko", ""),
                    "email": firm.get("email", ""),
                    "phone": firm.get("telefon", ""),
                    "website": firm.get("www", ""),
                    "ceidg_link": firm.get("link", ""),
                    "source": "CEIDG",
                }

                if firm.get("pkdGlowny"):
                    result["pkd_codes"].append(firm["pkdGlowny"])
                for pkd in firm.get("pkd", []):
                    code = pkd if isinstance(pkd, str) else pkd.get("kod", "") if isinstance(pkd, dict) else ""
                    if code and code != firm.get("pkdGlowny"):
                        result["pkd_codes"].append(code)

                results.append(result)

            return results

        except requests.exceptions.RequestException as e:
            logger.error(f"CEIDG API error for '{name}': {e}")
            return []

    # ---- KRS API (free, no auth) ----

    def search_krs(self, krs_number):
        """Look up a company in KRS by its KRS number."""
        if not krs_number:
            return None

        krs_number = str(krs_number).strip().zfill(10)
        self._rate_limit_krs()

        try:
            response = requests.get(
                f"{KRS_API_BASE}/OdpisAktualny/{krs_number}",
                headers={"Accept": "application/json"},
                timeout=30,
            )
            if response.status_code == 404:
                logger.info(f"KRS {krs_number} not found")
                return None
            response.raise_for_status()
            data = response.json()
            return self._parse_krs_response(data, krs_number)

        except requests.exceptions.RequestException as e:
            logger.error(f"KRS API error for '{krs_number}': {e}")
            return None

    def _parse_krs_response(self, data, krs_number):
        """Parse KRS OdpisAktualny JSON response into a normalized result dict."""
        odpis = data.get("odppisPelnyGrupa") or data.get("odppisTresciGrupa") or data
        dzial1 = odpis.get("dzial1", {})
        dzial3 = odpis.get("dzial3", {})

        dane_podmiotu = dzial1.get("danePodmiotu", {})
        siedzibaiadres = dzial1.get("siedzibaIAdres", {})
        siedziba = siedzibaiadres.get("siedziba", {})
        adres = siedzibaiadres.get("adres", {})

        name = dane_podmiotu.get("nazwa", "")
        nip = dane_podmiotu.get("identyfikatory", {}).get("nip", "")
        regon = dane_podmiotu.get("identyfikatory", {}).get("regon", "")

        address = {
            "street": adres.get("ulica", ""),
            "building": adres.get("nrDomu", ""),
            "unit": adres.get("nrLokalu", ""),
            "zip_code": adres.get("kodPocztowy", ""),
            "city": siedziba.get("miejscowosc", "") or adres.get("miejscowosc", ""),
            "municipality": siedziba.get("gmina", ""),
            "county": siedziba.get("powiat", ""),
            "province": siedziba.get("wojewodztwo", ""),
            "country": siedziba.get("kraj", adres.get("kraj", "")),
        }

        pkd_data = dzial3.get("przedmiotDzialalnosci", {})
        pkd_glowny_list = pkd_data.get("przedmiotPrzewazajacejDzialalnosci", [])
        pkd_pozostale_list = pkd_data.get("przedmiotPozostalejDzialalnosci", [])

        pkd_main = ""
        pkd_codes = []
        for p in pkd_glowny_list:
            code = p.get("kodDzial", "")
            if code:
                pkd_main = code
                pkd_codes.append(code)
        for p in pkd_pozostale_list:
            code = p.get("kodDzial", "")
            if code and code not in pkd_codes:
                pkd_codes.append(code)

        info_o_zakonczeniu = dane_podmiotu.get("czyOstatecznieWykreslony", False)
        is_active = not info_o_zakonczeniu

        return {
            "name": name,
            "nip": nip,
            "regon": regon,
            "krs": krs_number,
            "address": address,
            "address_str": _format_address(address),
            "street": address.get("street", ""),
            "building": address.get("building", ""),
            "unit": address.get("unit", ""),
            "zip_code": address.get("zip_code", ""),
            "city": address.get("city", ""),
            "municipality": address.get("municipality", ""),
            "county": address.get("county", ""),
            "province": address.get("province", ""),
            "country": address.get("country", ""),
            "mailing_address_str": "",
            "status": "AKTYWNY" if is_active else "WYKRESLONY",
            "is_active": is_active,
            "date_started": dane_podmiotu.get("dataOrzeczeniaOWpisaniu", ""),
            "date_ended": dane_podmiotu.get("dataOrzeczeniaOWykresleniu", ""),
            "date_suspended": "",
            "date_resumed": "",
            "date_deleted": dane_podmiotu.get("dataOrzeczeniaOWykresleniu", ""),
            "pkd_codes": pkd_codes,
            "pkd_main": pkd_main,
            "owner_first_name": "",
            "owner_last_name": "",
            "email": "",
            "phone": "",
            "website": "",
            "ceidg_link": "",
            "source": "KRS",
        }

    # ---- GUS REGON/BIR1 API (via gusregon library) ----

    def _gus_connect(self):
        """Connect to GUS BIR1 API using gusregon library."""
        if self._gus_client:
            return True

        if not self.gus_api_key:
            logger.warning("No GUS API key configured. Set GUS_API_KEY env var.")
            return False

        try:
            from gusregon import GUS

            if self.use_sandbox:
                self._gus_client = GUS(sandbox=True)
            else:
                self._gus_client = GUS(api_key=self.gus_api_key)

            logger.info("GUS API connected successfully.")
            return True

        except Exception as e:
            logger.error(f"GUS connection error: {e}")
            return False

    def _parse_gus_result(self, data):
        """Parse a gusregon search result dict into our normalized format."""
        if not data:
            return None

        address = {
            "street": data.get("Ulica", ""),
            "building": data.get("NrNieruchomosci", ""),
            "unit": data.get("NrLokalu", ""),
            "zip_code": data.get("KodPocztowy", ""),
            "city": data.get("Miejscowosc", ""),
            "municipality": data.get("Gmina", ""),
            "county": data.get("Powiat", ""),
            "province": data.get("Wojewodztwo", ""),
            "country": "PL",
        }

        status_end = data.get("DataZakonczeniaDzialalnosci", "")

        result = {
            "name": data.get("Nazwa", ""),
            "nip": data.get("Nip", ""),
            "regon": data.get("Regon", ""),
            "krs": "",
            "address": address,
            "address_str": _format_address(address),
            "street": address.get("street", ""),
            "building": address.get("building", ""),
            "unit": address.get("unit", ""),
            "zip_code": address.get("zip_code", ""),
            "city": address.get("city", ""),
            "municipality": address.get("municipality", ""),
            "county": address.get("county", ""),
            "province": address.get("province", ""),
            "country": address.get("country", ""),
            "mailing_address_str": "",
            "status": "WYKRESLONY" if status_end else "AKTYWNY",
            "is_active": not bool(status_end),
            "date_started": data.get("DataRozpoczeciaDzialalnosci", ""),
            "date_ended": status_end,
            "date_suspended": "",
            "date_resumed": "",
            "date_deleted": "",
            "pkd_codes": [],
            "pkd_main": "",
            "owner_first_name": "",
            "owner_last_name": "",
            "email": "",
            "phone": "",
            "website": "",
            "ceidg_link": "",
            "entity_type": data.get("Typ", ""),
            "source": "GUS",
        }

        return result

    def _enrich_with_full_report(self, result):
        """Fetch full GUS report to get PKD codes, email, phone, website."""
        if not self._gus_client or not result.get("regon"):
            return result

        try:
            full = self._gus_client.search(regon=result["regon"])
            if full:
                report = full
                if isinstance(report, list):
                    report = report[0] if report else {}

                # Get PKD codes
                regon = result["regon"]
                try:
                    pkd_list = self._gus_client.get_pkd(regon)
                    if pkd_list:
                        for pkd in pkd_list:
                            code = ""
                            is_main = False
                            if isinstance(pkd, dict):
                                code = pkd.get("Kod", "") or pkd.get("kod", "")
                                is_main = pkd.get("Przewazajace", "") == "1" or pkd.get("przewazajace", "") == "1"
                            elif isinstance(pkd, str):
                                code = pkd
                            if code and code not in result["pkd_codes"]:
                                result["pkd_codes"].append(code)
                            if is_main and code:
                                result["pkd_main"] = code
                        if not result["pkd_main"] and result["pkd_codes"]:
                            result["pkd_main"] = result["pkd_codes"][0]
                except Exception as e:
                    logger.debug(f"Could not get PKD for REGON {regon}: {e}")

                # Get address/contact from full report
                if isinstance(report, dict):
                    result["email"] = report.get("adresEmail", "") or report.get("AdresEmail", "") or result["email"]
                    result["phone"] = report.get("numerTelefonu", "") or report.get("NumerTelefonu", "") or result["phone"]
                    result["website"] = report.get("adresStronyinternetowej", "") or report.get("AdresStronyInternetowej", "") or result["website"]
                    krs = report.get("numerWRejestrzeEwidencji", "") or report.get("NumerWRejestrzeEwidencji", "")
                    if krs:
                        result["krs"] = krs

        except Exception as e:
            logger.debug(f"GUS full report error for REGON {result.get('regon')}: {e}")

        return result

    def search_gus_by_nip(self, nip):
        """Search GUS REGON by NIP number. Returns a result dict or None."""
        if not self._gus_connect():
            return None

        try:
            logger.info(f"GUS searching NIP: {nip}")
            data = self._gus_client.search(nip=nip)

            if not data:
                return None

            if isinstance(data, list):
                data = data[0] if data else None
            if not data:
                return None

            result = self._parse_gus_result(data)
            if result:
                result = self._enrich_with_full_report(result)

            return result

        except Exception as e:
            logger.error(f"GUS NIP search error for '{nip}': {e}")
            return None

    def search_gus_by_name(self, name):
        """Search GUS REGON by company name. Returns list of result dicts."""
        if not self._gus_connect():
            return []

        try:
            data = self._gus_client.search(name=name)

            if not data:
                return []

            if not isinstance(data, list):
                data = [data]

            results = []
            for item in data:
                result = self._parse_gus_result(item)
                if result:
                    result = self._enrich_with_full_report(result)
                    results.append(result)

            return results

        except Exception as e:
            logger.error(f"GUS search error for '{name}': {e}")
            return []

    # ---- Unified lookup ----

    def lookup_business(self, name, zip_code=None, city=None):
        """
        Look up a business by name using all available APIs.
        Cascade: CEIDG -> GUS -> KRS (if KRS number found).
        """
        best_result = None

        if self.ceidg_api_key:
            results = self.search_ceidg(name, zip_code, city)
            if results:
                best_result = self._best_match(results, name, zip_code, city)

        if not best_result and self.gus_api_key:
            results = self.search_gus_by_name(name)
            if results:
                best_result = self._best_match(results, name, zip_code, city)

        if best_result and best_result.get("krs"):
            krs_data = self.search_krs(best_result["krs"])
            if krs_data:
                best_result = self._merge_results(best_result, krs_data)

        if not best_result:
            logger.info(f"No API results for '{name}' ({zip_code}, {city})")

        return best_result

    @staticmethod
    def _merge_results(primary, secondary):
        """Merge two result dicts, preferring non-empty values from primary."""
        merged = dict(primary)
        for key, val in secondary.items():
            if key == "source":
                merged["source"] = f"{primary.get('source', '')},{secondary.get('source', '')}"
                continue
            if key in ("pkd_codes",):
                existing = set(merged.get(key, []))
                for item in val:
                    if item not in existing:
                        merged[key].append(item)
                        existing.add(item)
                continue
            if not merged.get(key) and val:
                merged[key] = val
        return merged

    @staticmethod
    def _best_match(results, name, zip_code=None, city=None):
        """Find the best matching result based on name, zip code, and city."""
        if not results:
            return None

        name_lower = name.lower().strip()

        scored = []
        for r in results:
            score = 0
            r_name = r.get("name", "").lower().strip()

            if r_name == name_lower:
                score += 10
            elif name_lower in r_name or r_name in name_lower:
                score += 5
            else:
                name_words = set(name_lower.split())
                result_words = set(r_name.split())
                overlap = name_words & result_words
                score += len(overlap) * 2

            if zip_code and r.get("zip_code"):
                if r["zip_code"].replace("-", "") == zip_code.replace("-", ""):
                    score += 5
                elif r["zip_code"][:2] == zip_code[:2]:
                    score += 2

            if city and r.get("city"):
                r_city = r["city"].lower().strip()
                city_lower = city.lower().strip()
                if r_city == city_lower:
                    score += 5
                elif city_lower in r_city or r_city in city_lower:
                    score += 3

            if r.get("is_active"):
                score += 1

            scored.append((score, r))

        scored.sort(key=lambda x: x[0], reverse=True)
        return scored[0][1] if scored[0][0] > 0 else results[0]
