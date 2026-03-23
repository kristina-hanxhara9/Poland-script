"""
Poland Government API client for looking up businesses and retrieving full details.

Supports:
- CEIDG API v2 (dane.biznes.gov.pl) - for sole proprietors
- KRS API (api-krs.ms.gov.pl) - for registered companies (free, no auth)
- GUS REGON/BIR1 API - for all entities by name (sandbox available)
"""

import os
import time
import logging
import requests
from xml.etree import ElementTree

logger = logging.getLogger(__name__)

# API endpoints
CEIDG_API_BASE = "https://dane.biznes.gov.pl/api/ceidg/v2"
KRS_API_BASE = "https://api-krs.ms.gov.pl/api/krs"
GUS_SANDBOX_URL = "https://wyszukiwarkaregon.stat.gov.pl/wsBIR/UslugaBIRzworku.svc"
GUS_PRODUCTION_URL = "https://wyszukiwarkaregontest.stat.gov.pl/wsBIR/UslugaBIRzworku.svc"
GUS_SANDBOX_KEY = "abcde12345abcde12345"

# Rate limiting: CEIDG allows 50 requests per 3 minutes
CEIDG_RATE_LIMIT_DELAY = 3.6  # seconds between requests (conservative)
# KRS: ~100 requests per 15 minutes
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
        self.ceidg_api_key = ceidg_api_key or os.environ.get("CEIDG_API_KEY")
        self.gus_api_key = gus_api_key or os.environ.get("GUS_API_KEY")
        self.use_sandbox = use_sandbox
        self._last_ceidg_request = 0
        self._last_krs_request = 0
        self._gus_session_id = None

        if use_sandbox and not self.gus_api_key:
            self.gus_api_key = GUS_SANDBOX_KEY

    def _rate_limit_ceidg(self):
        """Enforce rate limiting for CEIDG API calls."""
        elapsed = time.time() - self._last_ceidg_request
        if elapsed < CEIDG_RATE_LIMIT_DELAY:
            time.sleep(CEIDG_RATE_LIMIT_DELAY - elapsed)
        self._last_ceidg_request = time.time()

    def _rate_limit_krs(self):
        """Enforce rate limiting for KRS API calls."""
        elapsed = time.time() - self._last_krs_request
        if elapsed < KRS_RATE_LIMIT_DELAY:
            time.sleep(KRS_RATE_LIMIT_DELAY - elapsed)
        self._last_krs_request = time.time()

    def has_api_access(self):
        """Check if any API credentials are configured. KRS is always available."""
        return True  # KRS is open, always available

    # ---- CEIDG API v2 ----

    def search_ceidg(self, name, zip_code=None, city=None):
        """
        Search CEIDG for sole proprietors by name and optional location.

        Returns list of dicts with full business info.
        """
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
                    # Full address
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
                    # Mailing address
                    "mailing_address_str": _format_address(
                        _build_address_dict(firm.get("adresKorespondencyjny", {}))
                    ),
                    # Status
                    "status": firm.get("status", ""),
                    "is_active": firm.get("status", "") == "AKTYWNY",
                    # Dates
                    "date_started": firm.get("dataRozpoczecia", ""),
                    "date_ended": firm.get("dataZakonczenia", ""),
                    "date_suspended": firm.get("dataZawieszenia", ""),
                    "date_resumed": firm.get("dataWznowienia", ""),
                    "date_deleted": firm.get("dataWykreslenia", ""),
                    # PKD
                    "pkd_codes": [],
                    "pkd_main": firm.get("pkdGlowny", ""),
                    # Owner
                    "owner_first_name": owner.get("imie", ""),
                    "owner_last_name": owner.get("nazwisko", ""),
                    # Contact
                    "email": firm.get("email", ""),
                    "phone": firm.get("telefon", ""),
                    "website": firm.get("www", ""),
                    "ceidg_link": firm.get("link", ""),
                    # Source
                    "source": "CEIDG",
                }

                # Extract PKD codes
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
        """
        Look up a company in KRS by its KRS number.

        The official KRS API only supports lookup by KRS number (not by name).
        Returns a dict with full company info, or None.
        """
        if not krs_number:
            return None

        # Pad to 10 digits
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

        # Section 1: basic data
        dane_podmiotu = dzial1.get("danePodmiotu", {})
        siedzibaiadres = dzial1.get("siedzibaIAdres", {})
        siedziba = siedzibaiadres.get("siedziba", {})
        adres = siedzibaiadres.get("adres", {})

        name = dane_podmiotu.get("nazwa", "")
        nip = dane_podmiotu.get("identyfikatory", {}).get("nip", "")
        regon = dane_podmiotu.get("identyfikatory", {}).get("regon", "")

        # Address from KRS
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

        # Section 3: PKD codes
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

        # Status
        info_o_zakonczeniu = dane_podmiotu.get("czyOstatecznieWykreslony", False)
        is_active = not info_o_zakonczeniu

        result = {
            "name": name,
            "nip": nip,
            "regon": regon,
            "krs": krs_number,
            # Full address
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
            # Status
            "status": "AKTYWNY" if is_active else "WYKRESLONY",
            "is_active": is_active,
            # Dates
            "date_started": dane_podmiotu.get("dataOrzeczeniaOWpisaniu", ""),
            "date_ended": dane_podmiotu.get("dataOrzeczeniaOWykresleniu", ""),
            "date_suspended": "",
            "date_resumed": "",
            "date_deleted": dane_podmiotu.get("dataOrzeczeniaOWykresleniu", ""),
            # PKD
            "pkd_codes": pkd_codes,
            "pkd_main": pkd_main,
            # Owner (KRS has board members, not single owner)
            "owner_first_name": "",
            "owner_last_name": "",
            # Contact (KRS doesn't expose email/phone)
            "email": "",
            "phone": "",
            "website": "",
            "ceidg_link": "",
            # Source
            "source": "KRS",
        }

        return result

    # ---- GUS REGON/BIR1 API ----

    def _gus_login(self):
        """Authenticate with GUS BIR1 API and get session ID."""
        if not self.gus_api_key:
            logger.warning("No GUS API key configured. Set GUS_API_KEY env var.")
            return False

        url = GUS_SANDBOX_URL if self.use_sandbox else GUS_PRODUCTION_URL

        envelope = f"""<?xml version="1.0" encoding="utf-8"?>
        <soap:Envelope xmlns:soap="http://www.w3.org/2003/05/soap-envelope"
                       xmlns:ns="http://CIS/BIR/PUBL/2014/07">
            <soap:Header xmlns:wsa="http://www.w3.org/2005/08/addressing">
                <wsa:To>{url}</wsa:To>
                <wsa:Action>http://CIS/BIR/PUBL/2014/07/IUslugaBIRzworku/Zaloguj</wsa:Action>
            </soap:Header>
            <soap:Body>
                <ns:Zaloguj>
                    <ns:pKluczUzytkownika>{self.gus_api_key}</ns:pKluczUzytkownika>
                </ns:Zaloguj>
            </soap:Body>
        </soap:Envelope>"""

        headers = {"Content-Type": "application/soap+xml; charset=utf-8"}

        try:
            response = requests.post(url, data=envelope, headers=headers, timeout=30)
            response.raise_for_status()

            root = ElementTree.fromstring(response.text)
            for elem in root.iter():
                if "ZalogujResult" in elem.tag:
                    self._gus_session_id = elem.text
                    return bool(self._gus_session_id)
            return False

        except Exception as e:
            logger.error(f"GUS login error: {e}")
            return False

    def _gus_full_report(self, regon, entity_type):
        """
        Fetch a full report from GUS BIR1 for a given REGON.

        entity_type should be 'F' (osoba fizyczna / sole proprietor),
        'P' (osoba prawna / legal entity), or 'LP' (local unit of legal entity).
        """
        if not self._gus_session_id:
            return {}

        url = GUS_SANDBOX_URL if self.use_sandbox else GUS_PRODUCTION_URL

        # Choose report name based on entity type
        if entity_type == "F":
            report_name = "BIR11OsFizycznaDzworkczosc"
        elif entity_type == "P":
            report_name = "BIR11OsPrawna"
        else:
            return {}

        envelope = f"""<?xml version="1.0" encoding="utf-8"?>
        <soap:Envelope xmlns:soap="http://www.w3.org/2003/05/soap-envelope"
                       xmlns:ns="http://CIS/BIR/PUBL/2014/07">
            <soap:Header xmlns:wsa="http://www.w3.org/2005/08/addressing">
                <wsa:To>{url}</wsa:To>
                <wsa:Action>http://CIS/BIR/PUBL/2014/07/IUslugaBIRzworku/DanePobierzPelnyRaport</wsa:Action>
            </soap:Header>
            <soap:Body>
                <ns:DanePobierzPelnyRaport>
                    <ns:pRegon>{regon}</ns:pRegon>
                    <ns:pNazwaRaportu>{report_name}</ns:pNazwaRaportu>
                </ns:DanePobierzPelnyRaport>
            </soap:Body>
        </soap:Envelope>"""

        headers = {
            "Content-Type": "application/soap+xml; charset=utf-8",
            "sid": self._gus_session_id,
        }

        try:
            response = requests.post(url, data=envelope, headers=headers, timeout=30)
            response.raise_for_status()

            root = ElementTree.fromstring(response.text)
            for elem in root.iter():
                if "DanePobierzPelnyRaportResult" in elem.tag and elem.text:
                    inner = ElementTree.fromstring(elem.text)
                    for dane in inner.iter("dane"):
                        return {child.tag: (child.text or "").strip() for child in dane}
            return {}

        except Exception as e:
            logger.error(f"GUS full report error for REGON {regon}: {e}")
            return {}

    def search_gus_by_name(self, name):
        """
        Search GUS REGON by company name.

        Returns list of dicts with full business info.
        """
        if not self._gus_session_id:
            if not self._gus_login():
                return []

        url = GUS_SANDBOX_URL if self.use_sandbox else GUS_PRODUCTION_URL

        envelope = f"""<?xml version="1.0" encoding="utf-8"?>
        <soap:Envelope xmlns:soap="http://www.w3.org/2003/05/soap-envelope"
                       xmlns:ns="http://CIS/BIR/PUBL/2014/07"
                       xmlns:dat="http://CIS/BIR/PUBL/2014/07/DataContract">
            <soap:Header xmlns:wsa="http://www.w3.org/2005/08/addressing">
                <wsa:To>{url}</wsa:To>
                <wsa:Action>http://CIS/BIR/PUBL/2014/07/IUslugaBIRzworku/DaneSzukajPodmioty</wsa:Action>
            </soap:Header>
            <soap:Body>
                <ns:DaneSzukajPodmioty>
                    <ns:pParametryWyszukiwania>
                        <dat:Nazwa>{name}</dat:Nazwa>
                    </ns:pParametryWyszukiwania>
                </ns:DaneSzukajPodmioty>
            </soap:Body>
        </soap:Envelope>"""

        headers = {
            "Content-Type": "application/soap+xml; charset=utf-8",
            "sid": self._gus_session_id,
        }

        try:
            response = requests.post(url, data=envelope, headers=headers, timeout=30)
            response.raise_for_status()

            results = []
            root = ElementTree.fromstring(response.text)

            for elem in root.iter():
                if "DaneSzukajPodmiotyResult" in elem.tag and elem.text:
                    inner = ElementTree.fromstring(elem.text)
                    for dane in inner.iter("dane"):
                        regon = self._get_xml_text(dane, "Regon")
                        entity_type = self._get_xml_text(dane, "Typ")

                        address = {
                            "street": self._get_xml_text(dane, "Ulica"),
                            "building": self._get_xml_text(dane, "NrNieruchomosci"),
                            "unit": self._get_xml_text(dane, "NrLokalu"),
                            "zip_code": self._get_xml_text(dane, "KodPocztowy"),
                            "city": self._get_xml_text(dane, "Miejscowosc"),
                            "municipality": self._get_xml_text(dane, "Gmina"),
                            "county": self._get_xml_text(dane, "Powiat"),
                            "province": self._get_xml_text(dane, "Wojewodztwo"),
                            "country": "PL",
                        }

                        status_zakonczenia = self._get_xml_text(dane, "DataZakonczeniaDzialalnosci")

                        result = {
                            "name": self._get_xml_text(dane, "Nazwa"),
                            "nip": self._get_xml_text(dane, "Nip"),
                            "regon": regon,
                            "krs": "",
                            # Full address
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
                            # Status
                            "status": "WYKRESLONY" if status_zakonczenia else "AKTYWNY",
                            "is_active": not bool(status_zakonczenia),
                            # Dates
                            "date_started": self._get_xml_text(dane, "DataRozpoczeciaDzialalnosci"),
                            "date_ended": status_zakonczenia,
                            "date_suspended": "",
                            "date_resumed": "",
                            "date_deleted": "",
                            # PKD
                            "pkd_codes": [],
                            "pkd_main": self._get_xml_text(dane, "PKD"),
                            # Owner
                            "owner_first_name": "",
                            "owner_last_name": "",
                            # Contact
                            "email": "",
                            "phone": "",
                            "website": "",
                            "ceidg_link": "",
                            # Metadata
                            "entity_type": entity_type,
                            "source": "GUS",
                        }
                        if result["pkd_main"]:
                            result["pkd_codes"].append(result["pkd_main"])

                        # Try to get full report for more details
                        if regon and entity_type in ("F", "P"):
                            full = self._gus_full_report(regon, entity_type)
                            if full:
                                self._enrich_from_gus_report(result, full, entity_type)

                        results.append(result)

            return results

        except Exception as e:
            logger.error(f"GUS search error for '{name}': {e}")
            return []

    def _enrich_from_gus_report(self, result, report, entity_type):
        """Enrich a result dict with data from a GUS full report."""
        if entity_type == "P":
            # Legal entity report fields
            result["email"] = report.get("praw_adresEmail", "") or result["email"]
            result["phone"] = report.get("praw_numerTelefonu", "") or result["phone"]
            result["website"] = report.get("praw_adresStronyinternetowej", "") or result["website"]
            krs = report.get("praw_numerWRejestrzeEwidencji", "")
            if krs:
                result["krs"] = krs

            # Additional PKD from report
            for i in range(1, 10):
                pkd_key = f"praw_pkdKod{i}"
                code = report.get(pkd_key, "")
                if code and code not in result["pkd_codes"]:
                    result["pkd_codes"].append(code)

        elif entity_type == "F":
            # Sole proprietor report fields
            result["email"] = report.get("fiz_adresEmail", "") or result["email"]
            result["phone"] = report.get("fiz_numerTelefonu", "") or result["phone"]
            result["website"] = report.get("fiz_adresStronyinternetowej", "") or result["website"]

            for i in range(1, 10):
                pkd_key = f"fiz_pkdKod{i}"
                code = report.get(pkd_key, "")
                if code and code not in result["pkd_codes"]:
                    result["pkd_codes"].append(code)

    @staticmethod
    def _get_xml_text(element, tag):
        """Helper to get text from an XML element by tag name."""
        found = element.find(tag)
        if found is not None and found.text:
            return found.text.strip()
        return ""

    # ---- Unified lookup ----

    def lookup_business(self, name, zip_code=None, city=None):
        """
        Look up a business by name, zip code, and/or city using all available APIs.

        Cascade: CEIDG → GUS → KRS (if KRS number found via GUS).

        Returns dict with full business info, or None if not found.
        """
        best_result = None

        # Try CEIDG first (sole proprietors, most detailed)
        if self.ceidg_api_key:
            results = self.search_ceidg(name, zip_code, city)
            if results:
                best_result = self._best_match(results, name, zip_code, city)

        # Try GUS (all entity types)
        if not best_result and self.gus_api_key:
            results = self.search_gus_by_name(name)
            if results:
                best_result = self._best_match(results, name, zip_code, city)

        # If GUS found a KRS number, enrich with KRS data
        if best_result and best_result.get("krs"):
            krs_data = self.search_krs(best_result["krs"])
            if krs_data:
                best_result = self._merge_results(best_result, krs_data)

        # If nothing found via CEIDG/GUS, KRS can't help (no name search)
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
                # Merge lists
                existing = set(merged.get(key, []))
                for item in val:
                    if item not in existing:
                        merged[key].append(item)
                        existing.add(item)
                continue
            # Fill empty fields from secondary
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

            # Exact name match
            if r_name == name_lower:
                score += 10
            # Name contains search term
            elif name_lower in r_name or r_name in name_lower:
                score += 5
            # Partial word overlap
            else:
                name_words = set(name_lower.split())
                result_words = set(r_name.split())
                overlap = name_words & result_words
                score += len(overlap) * 2

            # Zip code match
            if zip_code and r.get("zip_code"):
                if r["zip_code"].replace("-", "") == zip_code.replace("-", ""):
                    score += 5
                elif r["zip_code"][:2] == zip_code[:2]:
                    score += 2

            # City match
            if city and r.get("city"):
                r_city = r["city"].lower().strip()
                city_lower = city.lower().strip()
                if r_city == city_lower:
                    score += 5
                elif city_lower in r_city or r_city in city_lower:
                    score += 3

            # Bonus for active businesses
            if r.get("is_active"):
                score += 1

            scored.append((score, r))

        scored.sort(key=lambda x: x[0], reverse=True)
        return scored[0][1] if scored[0][0] > 0 else results[0]
