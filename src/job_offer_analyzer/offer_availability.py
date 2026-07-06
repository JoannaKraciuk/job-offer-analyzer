from __future__ import annotations

from dataclasses import dataclass
import re

from bs4 import BeautifulSoup
import requests

from job_offer_analyzer.fetch_offer import USER_AGENT


REQUEST_TIMEOUT_SECONDS = 20

CLOSED_HTTP_STATUSES = {404, 410}
CLOSED_PHRASES = [
    "job no longer available",
    "job is no longer available",
    "position is no longer available",
    "position has been filled",
    "job has expired",
    "this job is closed",
    "application is closed",
    "not accepting applications",
    "recruitment has ended",
    "no longer accepting applications",
    "404 not found",
    "page not found",
    "oferta nieaktualna",
    "oferta wygasła",
    "ogłoszenie wygasło",
    "rekrutacja zakończona",
    "nie znaleziono oferty",
    "strona nie istnieje",
    "oferta została zamknięta",
    "aplikowanie zakończone",
]


@dataclass(frozen=True)
class AvailabilityResult:
    availability: str
    note: str
    http_status: int | None = None
    final_url: str = ""

    @property
    def is_available(self) -> bool:
        return self.availability == "Dostępna"

    @property
    def is_closed(self) -> bool:
        return self.availability == "Zamknięta"

    @property
    def is_uncertain(self) -> bool:
        return self.availability == "Niepewna"


def check_offer_availability(url: str) -> AvailabilityResult:
    cleaned_url = _normalize_url(url)
    try:
        response = requests.get(
            cleaned_url,
            headers={"User-Agent": USER_AGENT, "Accept-Language": "pl,en;q=0.8"},
            timeout=REQUEST_TIMEOUT_SECONDS,
            allow_redirects=True,
        )
    except requests.Timeout:
        return AvailabilityResult(
            availability="Niepewna",
            note="Timeout podczas pobierania strony.",
            final_url=cleaned_url,
        )
    except requests.RequestException as exc:
        return AvailabilityResult(
            availability="Niepewna",
            note=f"Nie udało się pobrać strony: {exc}",
            final_url=cleaned_url,
        )

    status_code = response.status_code
    if status_code in CLOSED_HTTP_STATUSES:
        return AvailabilityResult(
            availability="Zamknięta",
            note=f"Strona zwróciła HTTP {status_code}.",
            http_status=status_code,
            final_url=response.url,
        )

    if status_code >= 400:
        return AvailabilityResult(
            availability="Niepewna",
            note=f"Strona zwróciła HTTP {status_code}.",
            http_status=status_code,
            final_url=response.url,
        )

    content_type = response.headers.get("content-type", "")
    if "html" not in content_type.lower():
        return AvailabilityResult(
            availability="Niepewna",
            note=f"Link zwrócił nietypowy content-type: {content_type or 'brak'}.",
            http_status=status_code,
            final_url=response.url,
        )

    page_text = _page_text(response.text)
    closed_phrase = _find_closed_phrase(page_text)
    if closed_phrase:
        return AvailabilityResult(
            availability="Zamknięta",
            note=f"Na stronie znaleziono komunikat zamknięcia: {closed_phrase}.",
            http_status=status_code,
            final_url=response.url,
        )

    return AvailabilityResult(
        availability="Dostępna",
        note="Strona jest dostępna i nie znaleziono komunikatu zamknięcia oferty.",
        http_status=status_code,
        final_url=response.url,
    )


def _normalize_url(url: str) -> str:
    cleaned_url = url.strip()
    if not cleaned_url:
        raise ValueError("Brak linku do oferty.")

    if not cleaned_url.startswith(("http://", "https://")):
        cleaned_url = f"https://{cleaned_url}"

    return cleaned_url


def _page_text(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "noscript", "svg"]):
        tag.decompose()

    text = soup.get_text(" ", strip=True)
    return re.sub(r"\s+", " ", text).lower()


def _find_closed_phrase(page_text: str) -> str:
    for phrase in CLOSED_PHRASES:
        if phrase in page_text:
            return phrase
    return ""
