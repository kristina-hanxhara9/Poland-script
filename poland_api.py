"""
Poland Government API client for looking up businesses and retrieving full details.

Supports:
- CEIDG API v2 (dane.biznes.gov.pl) - for sole proprietors
- KRS API (api-krs.ms.gov.pl) - for registered companies (free, no auth)
- GUS REGON/BIR1 API - for all entities (sandbox available)
"""

import os
import re
import time
import logging
import warnings
import requests
import urllib3
from xml.etree import ElementTree

# Suppress SSL warnings for GUS API (known certificate chain issues)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

logger = logging.getLogger(__name__)

# API endpoints
CEIDG_API_BASE = "https://dane.biznes.gov.pl/api/ceidg/v2"
KRS_API_BASE = "https://api-krs.ms.gov.pl/api/krs"
GUS_SANDBOX_URL = "https://wyszukiwarkaregontest.stat.gov.pl/wsBIR/UslugaBIRzewnPubl.svc"
GUS_PRODUCTION_URL = "https://wyszukiwarkaregon.stat.gov.pl/wsBIR/UslugaBIRzewnPubl.svc"
GUS_SANDBOX_KEY = "abcde12345abcde12345"

# Rate limiting
CEIDG_RATE_LIMIT_DELAY = 3.6
KRS_RATE_LIMIT_DELAY = 9.0


def _extract_xml_from_mtom(raw_bytes):
    """
    Extract SOAP XML from an MTOM/XOP multipart response.

    GUS returns responses wrapped in MIME multipart. The XML envelope
    is between the MIME boundary markers. We extract it from raw bytes
    to avoid encoding issues.
    """
    # Decode to string, trying utf-8 first
    try:
        text = raw_bytes.decode("utf-8")
    except (UnicodeDecodeError, AttributeError):
        text = str(raw_bytes) if not isinstance(raw_bytes, str) else raw_bytes

    # Look for the SOAP envelope (GUS uses s: prefix)
    match = re.search(r'(<s:Envelope[^>]*>.*?</s:Envelope>)', text, re.DOTALL)
    if match:
        return match.group(1)

    # Also try soap: prefix
    match = re.search(r'(<soap:Envelope[^>]*>.*?</soap:Envelope>)', text, re.DOTALL)
    if match:
        return match.group(1)

    # Maybe it's already plain XML
    return text


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


def _gus_soap_request(url, envelope, session_id=None):
    """
    Send a SOAP request to GUS and extract the XML response from MTOM.

    Returns an ElementTree root element, or None on failure.
    """
    headers = {"Content-Type": "application/soap+xml; charset=utf-8"}
    if session_id:
        headers["sid"] = session_id

    response = requests.post(
        url,
        data=envelope.encode("utf-8"),
        headers=headers,
        timeout=30,
        verify=False,
    )
    response.raise_for_status()

    # Extract XML from MTOM multipart response (use raw bytes)
    xml_str = _extract_xml_from_mtom(response.content)
    logger.debug(f"GUS response XML: {xml_str[:300]}")
    return ElementTree.fromstring(xml_str)


class PolandAPIClient:
    """Client for Poland government business APIs."""

    def __init__(self, ceidg_api_key=None, gus_api_key=None, use_sandbox=False):
        self.ceidg_api_key = (ceidg_api_key or os.environ.get("CEIDG_API_KEY", "")).strip().strip('"').strip("'")
        self.gus_api_key = (gus_api_key or os.environ.get("GUS_API_KEY", "")).strip().strip('"').strip("'")
        self.use_sandbox = use_sandbox
        self._last_ceidg_request = 0
        self._last_krs_request = 0
        self._gus_session_id = None

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
        return True

    # ---- CEIDG API v2 ----

    def search_ceidg(self, name, zip_code=None, city=None):
        """Search CEIDG for sole proprietors by name and optional location."""
        if not self.ceidg_api_key:
            logger.warning("No CEIDG API key configured.")
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
                f"{CEIDG_API_BASE}/firmy", headers=headers, params=params, timeout=30,
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
                headers={"Accept": "application/json"}, timeout=30, verify=False,
            )
            if response.status_code == 404:
                return None
            response.raise_for_status()
            return self._parse_krs_response(response.json(), krs_number)
        except requests.exceptions.RequestException as e:
            logger.error(f"KRS API error for '{krs_number}': {e}")
            return None

    def _parse_krs_response(self, data, krs_number):
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

        is_active = not dane_podmiotu.get("czyOstatecznieWykreslony", False)

        return {
            "name": name, "nip": nip, "regon": regon, "krs": krs_number,
            "address": address, "address_str": _format_address(address),
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
            "date_suspended": "", "date_resumed": "",
            "date_deleted": dane_podmiotu.get("dataOrzeczeniaOWykresleniu", ""),
            "pkd_codes": pkd_codes, "pkd_main": pkd_main,
            "owner_first_name": "", "owner_last_name": "",
            "email": "", "phone": "", "website": "", "ceidg_link": "",
            "source": "KRS",
        }

    # ---- GUS REGON/BIR1 API ----

    @property
    def _gus_url(self):
        return GUS_SANDBOX_URL if self.use_sandbox else GUS_PRODUCTION_URL

    def _gus_login(self):
        """Authenticate with GUS BIR1 API and get session ID."""
        if not self.gus_api_key:
            logger.warning("No GUS API key configured. Set GUS_API_KEY env var.")
            return False

        envelope = (
            '<?xml version="1.0" encoding="utf-8"?>'
            '<soap:Envelope xmlns:soap="http://www.w3.org/2003/05/soap-envelope"'
            ' xmlns:ns="http://CIS/BIR/PUBL/2014/07">'
            '<soap:Header xmlns:wsa="http://www.w3.org/2005/08/addressing">'
            f'<wsa:To>{self._gus_url}</wsa:To>'
            '<wsa:Action>http://CIS/BIR/PUBL/2014/07/IUslugaBIRzewnPubl/Zaloguj</wsa:Action>'
            '</soap:Header>'
            '<soap:Body>'
            '<ns:Zaloguj>'
            f'<ns:pKluczUzytkownika>{self.gus_api_key}</ns:pKluczUzytkownika>'
            '</ns:Zaloguj>'
            '</soap:Body>'
            '</soap:Envelope>'
        )

        try:
            logger.info(f"GUS login: {self._gus_url}")
            root = _gus_soap_request(self._gus_url, envelope)

            for elem in root.iter():
                if "ZalogujResult" in elem.tag:
                    self._gus_session_id = elem.text
                    if self._gus_session_id:
                        logger.info(f"GUS login OK. Session: {self._gus_session_id[:8]}...")
                        return True
                    else:
                        logger.error("GUS login returned empty session ID — check your API key.")
                        return False

            logger.error("GUS login: no ZalogujResult in response")
            return False

        except Exception as e:
            logger.error(f"GUS login error: {e}")
            return False

    def _gus_ensure_session(self):
        """Ensure we have a valid GUS session."""
        if self._gus_session_id:
            return True
        return self._gus_login()

    def _gus_search(self, search_params_xml):
        """
        Run a DaneSzukajPodmioty query with the given search parameter XML fragment.
        Returns the parsed XML root, or None.
        """
        if not self._gus_ensure_session():
            return None

        envelope = (
            '<?xml version="1.0" encoding="utf-8"?>'
            '<soap:Envelope xmlns:soap="http://www.w3.org/2003/05/soap-envelope"'
            ' xmlns:ns="http://CIS/BIR/PUBL/2014/07"'
            ' xmlns:dat="http://CIS/BIR/PUBL/2014/07/DataContract">'
            '<soap:Header xmlns:wsa="http://www.w3.org/2005/08/addressing">'
            f'<wsa:To>{self._gus_url}</wsa:To>'
            '<wsa:Action>http://CIS/BIR/PUBL/2014/07/IUslugaBIRzewnPubl/DaneSzukajPodmioty</wsa:Action>'
            '</soap:Header>'
            '<soap:Body>'
            '<ns:DaneSzukajPodmioty>'
            '<ns:pParametryWyszukiwania>'
            f'{search_params_xml}'
            '</ns:pParametryWyszukiwania>'
            '</ns:DaneSzukajPodmioty>'
            '</soap:Body>'
            '</soap:Envelope>'
        )

        try:
            return _gus_soap_request(self._gus_url, envelope, self._gus_session_id)
        except Exception as e:
            logger.error(f"GUS search error: {e}")
            return None

    def _gus_fetch_report(self, regon, report_name):
        """Fetch a single report from GUS for a given REGON. Returns list of dane dicts."""
        if not self._gus_session_id or not regon:
            return []

        envelope = (
            '<?xml version="1.0" encoding="utf-8"?>'
            '<soap:Envelope xmlns:soap="http://www.w3.org/2003/05/soap-envelope"'
            ' xmlns:ns="http://CIS/BIR/PUBL/2014/07">'
            '<soap:Header xmlns:wsa="http://www.w3.org/2005/08/addressing">'
            f'<wsa:To>{self._gus_url}</wsa:To>'
            '<wsa:Action>http://CIS/BIR/PUBL/2014/07/IUslugaBIRzewnPubl/DanePobierzPelnyRaport</wsa:Action>'
            '</soap:Header>'
            '<soap:Body>'
            '<ns:DanePobierzPelnyRaport>'
            f'<ns:pRegon>{regon}</ns:pRegon>'
            f'<ns:pNazwaRaportu>{report_name}</ns:pNazwaRaportu>'
            '</ns:DanePobierzPelnyRaport>'
            '</soap:Body>'
            '</soap:Envelope>'
        )

        try:
            root = _gus_soap_request(self._gus_url, envelope, self._gus_session_id)
            results = []
            for elem in root.iter():
                if "DanePobierzPelnyRaportResult" in elem.tag and elem.text:
                    inner = ElementTree.fromstring(elem.text)
                    for dane in inner.iter("dane"):
                        results.append({child.tag: (child.text or "").strip() for child in dane})
            return results
        except Exception as e:
            logger.info(f"GUS report '{report_name}' error for REGON {regon}: {e}")
            return []

    def _gus_full_report(self, regon, entity_type):
        """Fetch full entity report from GUS."""
        # Correct report names for BIR 1.1
        if entity_type == "F":
            report_name = "BIR11OsFizycznaDaneOgolne"
        elif entity_type == "P":
            report_name = "BIR11OsPrawna"
        else:
            return {}

        results = self._gus_fetch_report(regon, report_name)
        return results[0] if results else {}

    def _gus_pkd_report(self, regon, entity_type):
        """Fetch PKD codes from GUS dedicated PKD report."""
        if entity_type == "F":
            report_name = "BIR11OsFizycznaPkd"
        elif entity_type == "P":
            report_name = "BIR11OsPrawnaPkd"
        else:
            return []

        results = self._gus_fetch_report(regon, report_name)
        logger.info(f"  PKD report '{report_name}' returned {len(results)} entries")
        if results:
            logger.info(f"  PKD report keys: {list(results[0].keys())}")
        pkd_codes = []
        for item in results:
            # PKD report returns one row per code — try all known field name variants
            code = (item.get("fiz_pkd_Kod", "") or item.get("praw_pkdKod", "")
                    or item.get("Kod", "") or item.get("kod", ""))
            desc = (item.get("fiz_pkd_Nazwa", "") or item.get("praw_pkdNazwa", "")
                    or item.get("Nazwa", "") or item.get("nazwa", ""))
            is_main = (item.get("fiz_pkd_Przewazajace", "") == "1"
                       or item.get("praw_pkdPrzewazajace", "") == "1"
                       or item.get("Przewazajace", "") == "1")
            if code:
                pkd_codes.append({"code": code, "description": desc, "is_main": is_main})
        return pkd_codes

    def _parse_gus_dane(self, dane):
        """Parse a single <dane> element from GUS search results."""
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

        status_end = self._get_xml_text(dane, "DataZakonczeniaDzialalnosci")

        result = {
            "name": self._get_xml_text(dane, "Nazwa"),
            "nip": self._get_xml_text(dane, "Nip"),
            "regon": regon,
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
            "date_started": self._get_xml_text(dane, "DataRozpoczeciaDzialalnosci"),
            "date_ended": status_end,
            "date_suspended": "", "date_resumed": "", "date_deleted": "",
            "pkd_codes": [],
            "pkd_main": self._get_xml_text(dane, "PKD"),
            "owner_first_name": "", "owner_last_name": "",
            "email": "", "phone": "", "website": "", "ceidg_link": "",
            "entity_type": entity_type,
            "source": "GUS",
        }

        if result["pkd_main"]:
            result["pkd_codes"].append(result["pkd_main"])

        # Enrich with full report (email, phone, website)
        if regon and entity_type in ("F", "P"):
            full = self._gus_full_report(regon, entity_type)
            if full:
                self._enrich_from_gus_report(result, full, entity_type)

            # Fetch PKD codes from dedicated PKD report
            pkd_list = self._gus_pkd_report(regon, entity_type)
            if pkd_list:
                result["pkd_descriptions"] = {}
                for pkd in pkd_list:
                    code = pkd["code"]
                    if code not in result["pkd_codes"]:
                        result["pkd_codes"].append(code)
                    if pkd["description"]:
                        result["pkd_descriptions"][code] = pkd["description"]
                    if pkd["is_main"]:
                        result["pkd_main"] = code
                # If we still have no main, use first code
                if not result["pkd_main"] and result["pkd_codes"]:
                    result["pkd_main"] = result["pkd_codes"][0]
                logger.info(f"  PKD codes: {result['pkd_codes']} (main: {result['pkd_main']})")

        return result

    def search_gus_by_nip(self, nip):
        """Search GUS by NIP. Returns a result dict or None."""
        root = self._gus_search(f'<dat:Nip>{nip}</dat:Nip>')
        if root is None:
            return None

        try:
            for elem in root.iter():
                if "DaneSzukajPodmiotyResult" in elem.tag and elem.text:
                    inner = ElementTree.fromstring(elem.text)
                    for dane in inner.iter("dane"):
                        return self._parse_gus_dane(dane)
            return None
        except Exception as e:
            logger.error(f"GUS NIP parse error for '{nip}': {e}")
            return None

    def search_gus_by_name(self, name):
        """Search GUS by company name. Returns list of result dicts."""
        root = self._gus_search(f'<dat:Nazwa>{name}</dat:Nazwa>')
        if root is None:
            return []

        try:
            results = []
            for elem in root.iter():
                if "DaneSzukajPodmiotyResult" in elem.tag and elem.text:
                    inner = ElementTree.fromstring(elem.text)
                    for dane in inner.iter("dane"):
                        results.append(self._parse_gus_dane(dane))
            return results
        except Exception as e:
            logger.error(f"GUS name search error for '{name}': {e}")
            return []

    def search_gus_by_zip(self, zip_code, max_results=10):
        """Search GUS by postal code. Returns list of result dicts.

        Uses the Miejscowosc/KodPocztowy parameters combined with a wildcard-like
        name search. GUS doesn't support zip-only search, so we search with
        common Polish business prefixes to find entities at that postal code.
        """
        # Normalize zip: remove dash if present (GUS expects XX-XXX format)
        z = zip_code.strip().replace("-", "").replace(" ", "")
        if len(z) == 5:
            formatted_zip = f"{z[:2]}-{z[2:]}"
        else:
            formatted_zip = zip_code.strip()

        # Try searching with zip + common prefixes to find businesses
        all_results = []
        seen_regons = set()

        # Strategy: search with short common business name prefixes
        # GUS requires a name param, so we use very common Polish words
        search_terms = ["*"]  # Try wildcard first
        if not all_results:
            search_terms = ["P", "S", "A", "M", "K", "B", "F", "Z", "W", "T"]

        for term in search_terms:
            params_xml = (
                f'<dat:Nazwa>{term}</dat:Nazwa>'
                f'<dat:KodPocztowy>{formatted_zip}</dat:KodPocztowy>'
            )
            root = self._gus_search(params_xml)
            if root is None:
                continue

            try:
                for elem in root.iter():
                    if "DaneSzukajPodmiotyResult" in elem.tag and elem.text:
                        inner = ElementTree.fromstring(elem.text)
                        for dane in inner.iter("dane"):
                            regon = self._get_xml_text(dane, "Regon")
                            if regon and regon not in seen_regons:
                                seen_regons.add(regon)
                                result = self._parse_gus_dane(dane)
                                all_results.append(result)
            except Exception as e:
                logger.error(f"GUS zip search error for '{formatted_zip}' term '{term}': {e}")

            if len(all_results) >= max_results:
                break

        logger.info(f"  GUS zip search '{formatted_zip}': found {len(all_results)} unique businesses")
        return all_results[:max_results]

    def _enrich_from_gus_report(self, result, report, entity_type):
        """Enrich a result dict with data from a GUS full report."""
        if entity_type == "P":
            result["email"] = report.get("praw_adresEmail", "") or result["email"]
            result["phone"] = report.get("praw_numerTelefonu", "") or result["phone"]
            result["website"] = report.get("praw_adresStronyinternetowej", "") or result["website"]
            krs = report.get("praw_numerWRejestrzeEwidencji", "")
            if krs:
                result["krs"] = krs
            for i in range(1, 10):
                code = report.get(f"praw_pkdKod{i}", "")
                if code and code not in result["pkd_codes"]:
                    result["pkd_codes"].append(code)

        elif entity_type == "F":
            result["email"] = report.get("fiz_adresEmail", "") or result["email"]
            result["phone"] = report.get("fiz_numerTelefonu", "") or result["phone"]
            result["website"] = report.get("fiz_adresStronyinternetowej", "") or result["website"]
            for i in range(1, 10):
                code = report.get(f"fiz_pkdKod{i}", "")
                if code and code not in result["pkd_codes"]:
                    result["pkd_codes"].append(code)

    @staticmethod
    def _get_xml_text(element, tag):
        found = element.find(tag)
        if found is not None and found.text:
            return found.text.strip()
        return ""

    # ---- Unified lookup ----

    def lookup_business(self, name, zip_code=None, city=None):
        """Look up a business by name. Cascade: CEIDG -> GUS -> KRS."""
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
        merged = dict(primary)
        for key, val in secondary.items():
            if key == "source":
                merged["source"] = f"{primary.get('source', '')},{secondary.get('source', '')}"
                continue
            if key == "pkd_codes":
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
                overlap = set(name_lower.split()) & set(r_name.split())
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
