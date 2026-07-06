from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date


UNKNOWN_VALUE = "Do uzupełnienia"


@dataclass(frozen=True)
class OfferRecord:
    company: str
    title: str
    link: str
    added_at: date = field(default_factory=date.today)
    last_checked_at: date = field(default_factory=date.today)
    availability: str = "Dostępna"
    status: str = "Nowa"
    cv_match: str = UNKNOWN_VALUE
    priority: str = UNKNOWN_VALUE
    work_mode: str = UNKNOWN_VALUE
    location: str = UNKNOWN_VALUE
    contract_type: str = UNKNOWN_VALUE
    rate_expectations: str = UNKNOWN_VALUE
    seniority: str = UNKNOWN_VALUE
    days_since_check: int = 0
    must_have_summary: str = UNKNOWN_VALUE
    nice_to_have_summary: str = UNKNOWN_VALUE
    risks_notes: str = UNKNOWN_VALUE
    next_step: str = "Przeanalizować ofertę i porównać z CV"
    source: str | None = None


@dataclass(frozen=True)
class HistoryRecord:
    checked_at: date
    offer_id: str
    company: str
    title: str
    link: str
    result: str
    checked_scope: str
    note: str
    source: str


@dataclass(frozen=True)
class OfferDraft:
    company: str = ""
    title: str = ""
    link: str = ""
    location: str = UNKNOWN_VALUE
    work_mode: str = UNKNOWN_VALUE
    contract_type: str = UNKNOWN_VALUE
    rate_expectations: str = UNKNOWN_VALUE
    seniority: str = UNKNOWN_VALUE
    must_have_summary: str = UNKNOWN_VALUE
    nice_to_have_summary: str = UNKNOWN_VALUE
    risks_notes: str = UNKNOWN_VALUE
    source_text: str = ""


@dataclass(frozen=True)
class AvailabilityRefreshRow:
    offer_id: str
    company: str
    title: str
    link: str
    previous_availability: str
    availability: str
    note: str
    changed: bool


@dataclass(frozen=True)
class AvailabilityRefreshSummary:
    checked_count: int
    available_count: int
    closed_count: int
    uncertain_count: int
    changed_count: int
    results: list[AvailabilityRefreshRow]
