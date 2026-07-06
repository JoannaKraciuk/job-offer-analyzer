from __future__ import annotations

from collections.abc import Iterable
import json
import re
from urllib.parse import urlparse

from bs4 import BeautifulSoup
import requests

from job_offer_analyzer.models import OfferDraft, UNKNOWN_VALUE
from job_offer_analyzer.salary_parser import (
    extract_salary_info,
    extract_salary_info_from_structured_data,
)


REQUEST_TIMEOUT_SECONDS = 20
MAX_SUMMARY_LENGTH = 500
MAX_SOURCE_TEXT_LENGTH = 12000

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/120.0 Safari/537.36"
)

WORK_MODE_KEYWORDS = {
    "Remote": ["remote", "zdalnie", "zdalna", "zdalny", "work from home"],
    "Hybrid": ["hybrid", "hybrydowo", "hybrydowa", "hybrydowy"],
    "Office": ["office", "on-site", "onsite", "stacjonarnie", "biuro", "w biurze"],
}

SENIORITY_KEYWORDS = {
    "Intern": ["intern", "internship", "staż", "praktyk"],
    "Junior": ["junior", "młodszy", "entry level"],
    "Mid": ["mid", "regular", "mid-level"],
    "Senior": ["senior", "starszy"],
    "Experienced": ["experienced", "doświadczony", "experience"],
}

CONTRACT_KEYWORDS = [
    ("B2B", ["b2b", "b2b contracts"]),
    ("Umowa o pracę", ["umowa o pracę", "employment contract", "full time employment"]),
    ("Kontrakt", ["contract", "contracts", "kontrakt"]),
    ("Umowa zlecenie", ["umowa zlecenie"]),
    ("Pełny etat", ["full-time", "full time", "pełny etat"]),
    ("Część etatu", ["part-time", "part time", "część etatu"]),
]

MUST_HAVE_HEADINGS = [
    "requirements",
    "required skills",
    "qualifications",
    "what you bring",
    "what we're looking for",
    "must have",
    "wymagania",
    "czego oczekujemy",
    "kogo szukamy",
]

NICE_TO_HAVE_HEADINGS = [
    "nice to have",
    "preferred",
    "bonus",
    "mile widziane",
    "dodatkowym atutem",
]

RISK_KEYWORDS = [
    "timezone",
    "time zone",
    "overlap",
    "relocation",
    "travel",
    "night shift",
    "weekend",
    "strefa czasowa",
    "relokacja",
    "podróże",
    "weekendy",
]
RATE_CONTEXT_KEYWORDS = [
    "salary",
    "rate",
    "compensation",
    "pay",
    "wage",
    "wynagrodzenie",
    "stawka",
    "widełki",
    "oczekiwania finansowe",
    "expected",
]


class OfferFetchError(RuntimeError):
    pass


def fetch_offer_from_url(url: str) -> OfferDraft:
    normalized_url = _normalize_url(url)
    response = _fetch_html(normalized_url)
    soup = BeautifulSoup(response.text, "html.parser")

    job_posting = _find_job_posting(soup)
    page_text = _clean_text_from_soup(soup)
    description_text = _clean_html_text(_job_value(job_posting, "description"))
    source_text = description_text or page_text

    title = (
        _text_or_empty(_job_value(job_posting, "title"))
        or _meta_content(soup, "og:title", "twitter:title")
        or _page_title(soup)
    )
    company = (
        _company_from_job_posting(job_posting)
        or _meta_content(soup, "og:site_name", "application-name")
        or _domain_name(normalized_url)
    )

    location = _location_from_job_posting(job_posting) or _infer_location(source_text)
    work_mode = _work_mode_from_job_posting(job_posting) or _infer_from_keywords(
        source_text, WORK_MODE_KEYWORDS
    )
    salary = extract_salary_info_from_structured_data(job_posting, source_text)
    if salary.amount_min is None:
        salary = extract_salary_info(source_text)

    return OfferDraft(
        company=_trim(company, 120),
        title=_trim(_clean_title(title, company), 180),
        link=normalized_url,
        location=_trim(location, 180) or UNKNOWN_VALUE,
        work_mode=work_mode or UNKNOWN_VALUE,
        contract_type=_infer_contract_type(source_text),
        rate_expectations=salary.display_value,
        salary=salary,
        seniority=_infer_from_keywords(source_text, SENIORITY_KEYWORDS) or UNKNOWN_VALUE,
        must_have_summary=_section_summary(source_text, MUST_HAVE_HEADINGS),
        nice_to_have_summary=_section_summary(source_text, NICE_TO_HAVE_HEADINGS),
        risks_notes=_risk_summary(source_text),
        source_text=_trim(source_text, MAX_SOURCE_TEXT_LENGTH),
    )


def _normalize_url(url: str) -> str:
    stripped_url = url.strip()
    if not stripped_url:
        raise OfferFetchError("Wklej link do oferty.")

    if not stripped_url.startswith(("http://", "https://")):
        stripped_url = f"https://{stripped_url}"

    parsed_url = urlparse(stripped_url)
    if parsed_url.scheme not in {"http", "https"} or not parsed_url.netloc:
        raise OfferFetchError("Link ma nieprawidłowy format.")

    return stripped_url


def _fetch_html(url: str) -> requests.Response:
    try:
        response = requests.get(
            url,
            headers={"User-Agent": USER_AGENT, "Accept-Language": "pl,en;q=0.8"},
            timeout=REQUEST_TIMEOUT_SECONDS,
        )
        response.raise_for_status()
    except requests.RequestException as exc:
        raise OfferFetchError(f"Nie udało się pobrać strony: {exc}") from exc

    content_type = response.headers.get("content-type", "")
    if "html" not in content_type.lower():
        raise OfferFetchError("Link nie zwrócił strony HTML.")

    return response


def _find_job_posting(soup: BeautifulSoup) -> dict:
    for script in soup.find_all("script", type=lambda value: value and "ld+json" in value):
        raw_json = script.string or script.get_text(" ", strip=True)
        if not raw_json:
            continue

        try:
            data = json.loads(raw_json)
        except json.JSONDecodeError:
            continue

        posting = _walk_json_for_job_posting(data)
        if posting:
            return posting

    return {}


def _walk_json_for_job_posting(data) -> dict:
    if isinstance(data, dict):
        data_type = data.get("@type")
        types = data_type if isinstance(data_type, list) else [data_type]
        if any(str(value).lower() == "jobposting" for value in types if value):
            return data

        for value in data.values():
            result = _walk_json_for_job_posting(value)
            if result:
                return result

    if isinstance(data, list):
        for item in data:
            result = _walk_json_for_job_posting(item)
            if result:
                return result

    return {}


def _job_value(job_posting: dict, key: str):
    return job_posting.get(key) if job_posting else None


def _company_from_job_posting(job_posting: dict) -> str:
    organization = _job_value(job_posting, "hiringOrganization")
    if isinstance(organization, dict):
        return _text_or_empty(organization.get("name"))
    return _text_or_empty(organization)


def _location_from_job_posting(job_posting: dict) -> str:
    location = _job_value(job_posting, "jobLocation")
    locations = location if isinstance(location, list) else [location]
    values: list[str] = []

    for item in locations:
        if not isinstance(item, dict):
            continue

        address = item.get("address")
        if isinstance(address, dict):
            values.extend(
                _text_or_empty(address.get(key))
                for key in ("addressLocality", "addressRegion", "addressCountry")
            )
        else:
            values.append(_text_or_empty(item.get("name")))

    cleaned_values = [value for value in values if value]
    return " / ".join(dict.fromkeys(cleaned_values))


def _work_mode_from_job_posting(job_posting: dict) -> str:
    location_type = _text_or_empty(_job_value(job_posting, "jobLocationType"))
    if location_type.upper() == "TELECOMMUTE":
        return "Remote"
    return ""


def _meta_content(soup: BeautifulSoup, *names: str) -> str:
    for name in names:
        tag = soup.find("meta", attrs={"property": name}) or soup.find(
            "meta", attrs={"name": name}
        )
        if tag and tag.get("content"):
            return _clean_whitespace(tag["content"])
    return ""


def _page_title(soup: BeautifulSoup) -> str:
    if soup.title and soup.title.string:
        return _clean_whitespace(soup.title.string)
    return ""


def _clean_text_from_soup(soup: BeautifulSoup) -> str:
    for tag in soup(["script", "style", "noscript", "svg"]):
        tag.decompose()

    for tag in soup(["header", "footer", "nav", "aside"]):
        tag.decompose()

    lines = [
        _clean_whitespace(line)
        for line in soup.get_text("\n").splitlines()
        if _clean_whitespace(line)
    ]
    return "\n".join(_dedupe_lines(lines))


def _clean_html_text(value) -> str:
    if not value:
        return ""
    soup = BeautifulSoup(str(value), "html.parser")
    return _clean_text_from_soup(soup)


def _clean_title(title: str, company: str) -> str:
    cleaned_title = _clean_whitespace(title)
    if not cleaned_title:
        return ""

    for separator in (" | ", " – ", " — "):
        if separator in cleaned_title:
            cleaned_title = cleaned_title.split(separator)[0].strip()
            break

    if company:
        cleaned_title = re.sub(re.escape(company), "", cleaned_title, flags=re.I)

    cleaned_title = cleaned_title.strip(" -–—|")
    return cleaned_title or title


def _infer_location(text: str) -> str:
    lines = _candidate_lines(text)
    for line in lines:
        lower_line = line.lower()
        if any(word in lower_line for word in ["location", "lokalizacja", "poland", "polska"]):
            return _trim(line, 180)
    return ""


def _infer_from_keywords(text: str, keyword_map: dict[str, list[str]]) -> str:
    lower_text = text.lower()
    for label, keywords in keyword_map.items():
        if any(keyword.lower() in lower_text for keyword in keywords):
            return label
    return ""


def _infer_contract_type(text: str) -> str:
    lower_text = text.lower()
    matches = []
    for label, keywords in CONTRACT_KEYWORDS:
        if any(keyword in lower_text for keyword in keywords):
            matches.append(label)

    if not matches:
        return UNKNOWN_VALUE

    unique_matches = list(dict.fromkeys(matches))
    if "B2B" in unique_matches and "Kontrakt" in unique_matches:
        unique_matches.remove("Kontrakt")

    return "; ".join(unique_matches[:4])


def _infer_rate(text: str) -> str:
    lines = [
        line
        for line in _candidate_lines(text)
        if any(keyword in line.lower() for keyword in RATE_CONTEXT_KEYWORDS)
    ]
    if not lines:
        return UNKNOWN_VALUE

    patterns = [
        r"(?:USD|EUR|PLN)\s*[0-9][0-9\s,.]*(?:[-–]\s*[0-9][0-9\s,.]*)?",
        r"[0-9][0-9\s,.]*(?:[-–]\s*[0-9][0-9\s,.]*)?\s*(?:USD|EUR|PLN|zł|€|\$)",
    ]
    matches: list[str] = []
    for line in lines:
        for pattern in patterns:
            matches.extend(re.findall(pattern, line, flags=re.I))

    cleaned_matches = [_clean_whitespace(match) for match in matches if match]
    if not cleaned_matches:
        return UNKNOWN_VALUE

    unique_matches = list(dict.fromkeys(cleaned_matches))
    pln_matches = [
        match for match in unique_matches if re.search(r"\bPLN\b|zł", match, flags=re.I)
    ]
    preferred_matches = pln_matches or unique_matches
    return "; ".join(preferred_matches[:3])


def _section_summary(text: str, headings: Iterable[str]) -> str:
    lines = _candidate_lines(text)
    lower_headings = [heading.lower() for heading in headings]

    for index, line in enumerate(lines):
        lower_line = line.lower()
        if not any(heading in lower_line for heading in lower_headings):
            continue

        section_lines = lines[index + 1 : index + 9]
        if section_lines:
            return _trim("; ".join(section_lines), MAX_SUMMARY_LENGTH)

        return _trim(line, MAX_SUMMARY_LENGTH)

    return UNKNOWN_VALUE


def _risk_summary(text: str) -> str:
    lines = _candidate_lines(text)
    risk_lines = [
        line
        for line in lines
        if any(keyword in line.lower() for keyword in RISK_KEYWORDS)
    ]
    if not risk_lines:
        return UNKNOWN_VALUE
    return _trim("; ".join(risk_lines[:4]), MAX_SUMMARY_LENGTH)


def _candidate_lines(text: str) -> list[str]:
    return [
        line
        for line in (_clean_whitespace(line) for line in text.splitlines())
        if 4 <= len(line) <= 260
    ]


def _dedupe_lines(lines: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for line in lines:
        key = line.lower()
        if key in seen:
            continue
        seen.add(key)
        result.append(line)
    return result


def _domain_name(url: str) -> str:
    domain = urlparse(url).netloc.removeprefix("www.")
    return domain.split(":")[0]


def _text_or_empty(value) -> str:
    if isinstance(value, str):
        return _clean_whitespace(value)
    return ""


def _clean_whitespace(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def _trim(value: str, max_length: int) -> str:
    cleaned_value = _clean_whitespace(value)
    if len(cleaned_value) <= max_length:
        return cleaned_value
    return f"{cleaned_value[: max_length - 3].rstrip()}..."
