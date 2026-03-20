"""
Poland Government API client for looking up businesses and retrieving PKD codes.

Supports:
- CEIDG API v2 (dane.biznes.gov.pl) - for sole proprietors
- GUS REGON/BIR1 API - for PKD code lookups (sandbox available)
- Fallback keyword-based matching when no API key is available
"""

import os
import time
import logging
import requests
from xml.etree import ElementTree

logger = logging.getLogger(__name__)

# API endpoints
CEIDG_API_BASE = "https://dane.biznes.gov.pl/api/ceidg/v2"
GUS_SANDBOX_URL = "https://wyszukiwarkaregon.stat.gov.pl/wsBIR/UslugaBIRzworku.svc"
GUS_PRODUCTION_URL = "https://wyszukiwarkaregontest.stat.gov.pl/wsBIR/UslugaBIRzworku.svc"
GUS_SANDBOX_KEY = "abcde12345abcde12345"

# Rate limiting: CEIDG allows 50 requests per 3 minutes
CEIDG_RATE_LIMIT_DELAY = 3.6  # seconds between requests (conservative)


class PolandAPIClient:
    """Client for Poland government business APIs."""

    def __init__(self, ceidg_api_key=None, gus_api_key=None, use_sandbox=False):
        self.ceidg_api_key = ceidg_api_key or os.environ.get("CEIDG_API_KEY")
        self.gus_api_key = gus_api_key or os.environ.get("GUS_API_KEY")
        self.use_sandbox = use_sandbox
        self._last_request_time = 0
        self._gus_session_id = None

        if use_sandbox and not self.gus_api_key:
            self.gus_api_key = GUS_SANDBOX_KEY

    def _rate_limit(self):
        """Enforce rate limiting between API calls."""
        elapsed = time.time() - self._last_request_time
        if elapsed < CEIDG_RATE_LIMIT_DELAY:
            time.sleep(CEIDG_RATE_LIMIT_DELAY - elapsed)
        self._last_request_time = time.time()

    def has_api_access(self):
        """Check if any API credentials are configured."""
        return bool(self.ceidg_api_key or self.gus_api_key)

    # ---- CEIDG API ----

    def search_ceidg(self, name, zip_code=None):
        """
        Search CEIDG for businesses by name and optional zip code.

        Args:
            name: Business name or entrepreneur name
            zip_code: Polish postal code (format: XX-XXX)

        Returns:
            List of dicts with business info including PKD codes, or empty list.
        """
        if not self.ceidg_api_key:
            logger.warning("No CEIDG API key configured. Set CEIDG_API_KEY env var.")
            return []

        self._rate_limit()

        headers = {
            "Authorization": f"Bearer {self.ceidg_api_key}",
            "Accept": "application/json",
        }

        params = {"firma": name}
        if zip_code:
            params["kod"] = zip_code

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
                result = {
                    "name": firm.get("nazwa", ""),
                    "nip": firm.get("nip", ""),
                    "regon": firm.get("regon", ""),
                    "zip_code": firm.get("adresDzialalnosci", {}).get("kod", ""),
                    "city": firm.get("adresDzialalnosci", {}).get("miasto", ""),
                    "pkd_codes": [],
                    "pkd_main": firm.get("pkdGlowny", ""),
                    "status": firm.get("status", ""),
                    "source": "CEIDG",
                }

                # Extract PKD codes
                if firm.get("pkdGlowny"):
                    result["pkd_codes"].append(firm["pkdGlowny"])
                for pkd in firm.get("pkdPozostale", []):
                    if isinstance(pkd, str):
                        result["pkd_codes"].append(pkd)
                    elif isinstance(pkd, dict):
                        result["pkd_codes"].append(pkd.get("kod", ""))

                results.append(result)

            return results

        except requests.exceptions.RequestException as e:
            logger.error(f"CEIDG API error for '{name}': {e}")
            return []

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
            # Extract session ID from SOAP response
            for elem in root.iter():
                if "ZalogujResult" in elem.tag:
                    self._gus_session_id = elem.text
                    return bool(self._gus_session_id)
            return False

        except Exception as e:
            logger.error(f"GUS login error: {e}")
            return False

    def search_gus_by_name(self, name):
        """
        Search GUS REGON by company name.

        Args:
            name: Business name to search for

        Returns:
            List of dicts with business info including PKD codes.
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
            self._rate_limit()
            response = requests.post(url, data=envelope, headers=headers, timeout=30)
            response.raise_for_status()

            results = []
            root = ElementTree.fromstring(response.text)

            for elem in root.iter():
                if "DaneSzukajPodmiotyResult" in elem.tag and elem.text:
                    # Parse the inner XML result
                    inner = ElementTree.fromstring(elem.text)
                    for dane in inner.iter("dane"):
                        result = {
                            "name": self._get_xml_text(dane, "Nazwa"),
                            "nip": self._get_xml_text(dane, "Nip"),
                            "regon": self._get_xml_text(dane, "Regon"),
                            "zip_code": self._get_xml_text(dane, "KodPocztowy"),
                            "city": self._get_xml_text(dane, "Miejscowosc"),
                            "pkd_codes": [],
                            "pkd_main": self._get_xml_text(dane, "PKD"),
                            "status": "",
                            "source": "GUS",
                        }
                        if result["pkd_main"]:
                            result["pkd_codes"].append(result["pkd_main"])
                        results.append(result)

            return results

        except Exception as e:
            logger.error(f"GUS search error for '{name}': {e}")
            return []

    @staticmethod
    def _get_xml_text(element, tag):
        """Helper to get text from an XML element by tag name."""
        found = element.find(tag)
        if found is not None and found.text:
            return found.text.strip()
        return ""

    # ---- Unified lookup ----

    def lookup_business(self, name, zip_code=None):
        """
        Look up a business by name and zip code using available APIs.

        Tries CEIDG first, then GUS, returns best match.

        Args:
            name: Business name
            zip_code: Polish postal code (XX-XXX format)

        Returns:
            Dict with business info and PKD codes, or None if not found.
        """
        # Try CEIDG first
        if self.ceidg_api_key:
            results = self.search_ceidg(name, zip_code)
            if results:
                best = self._best_match(results, name, zip_code)
                if best:
                    return best

        # Try GUS
        if self.gus_api_key:
            results = self.search_gus_by_name(name)
            if results:
                best = self._best_match(results, name, zip_code)
                if best:
                    return best

        logger.info(f"No API results for '{name}' ({zip_code})")
        return None

    @staticmethod
    def _best_match(results, name, zip_code):
        """Find the best matching result based on name and zip code similarity."""
        if not results:
            return None

        name_lower = name.lower().strip()

        # Score each result
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

            scored.append((score, r))

        scored.sort(key=lambda x: x[0], reverse=True)
        return scored[0][1] if scored[0][0] > 0 else results[0]
