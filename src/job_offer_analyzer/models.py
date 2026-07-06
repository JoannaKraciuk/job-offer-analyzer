from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date


UNKNOWN_VALUE = "Do uzupełnienia"
HOURS_PER_MONTH = 160


@dataclass(frozen=True)
class SalaryInfo:
    original_text: str = UNKNOWN_VALUE
    currency: str = UNKNOWN_VALUE
    amount_min: float | None = None
    amount_max: float | None = None
    period: str = UNKNOWN_VALUE
    tax_type: str = UNKNOWN_VALUE
    exchange_rate_to_pln: float | None = None
    exchange_rate_date: str = ""
    pln_min_monthly: float | None = None
    pln_max_monthly: float | None = None
    pln_min_hourly: float | None = None
    pln_max_hourly: float | None = None
    conversion_assumptions: str = ""

    @property
    def display_value(self) -> str:
        if self.currency == UNKNOWN_VALUE or self.amount_min is None:
            return UNKNOWN_VALUE

        amount = _format_number(self.amount_min)
        if self.amount_max is not None and self.amount_max != self.amount_min:
            amount = f"{amount}-{_format_number(self.amount_max)}"

        details = [amount, self.currency]
        if self.period != UNKNOWN_VALUE:
            details.append(self.period)
        if self.tax_type != UNKNOWN_VALUE:
            details.append(self.tax_type)

        return " ".join(details)


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
    salary: SalaryInfo = field(default_factory=SalaryInfo)
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
    salary: SalaryInfo = field(default_factory=SalaryInfo)
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


@dataclass(frozen=True)
class SalaryRefreshRow:
    offer_id: str
    company: str
    title: str
    link: str
    salary_display: str
    updated: bool
    note: str


@dataclass(frozen=True)
class SalaryRefreshSummary:
    checked_count: int
    updated_count: int
    missing_count: int
    failed_count: int
    results: list[SalaryRefreshRow]


def _format_number(value: float) -> str:
    if value.is_integer():
        return str(int(value))
    return f"{value:.2f}".rstrip("0").rstrip(".")
