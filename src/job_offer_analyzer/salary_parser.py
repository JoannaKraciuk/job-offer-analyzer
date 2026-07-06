from __future__ import annotations

from dataclasses import replace
from datetime import date
import re

import requests

from job_offer_analyzer.models import HOURS_PER_MONTH, SalaryInfo, UNKNOWN_VALUE


SUPPORTED_CURRENCIES = {"PLN", "EUR", "USD"}
NBP_RATE_URL = "https://api.nbp.pl/api/exchangerates/rates/A/{currency}/?format=json"
RATE_TIMEOUT_SECONDS = 10

CURRENCY_PATTERN = r"PLN|EUR|USD|zł|zloty|złoty|€|\$"
NUMBER_PATTERN = r"\d[\d\s\u00a0.,]*"
RANGE_SEPARATOR_PATTERN = r"(?:-|–|—|to|do)"

SALARY_CONTEXT_KEYWORDS = [
    "salary",
    "rate",
    "compensation",
    "pay",
    "wage",
    "wynagrodzenie",
    "stawka",
    "widełki",
    "wynagrodzenia",
    "oczekiwania finansowe",
    "expected",
]

PERIOD_KEYWORDS = {
    "godzinowo": [
        "/h",
        " per hour",
        "hourly",
        "godz",
        "godzin",
        "za godzinę",
        "na godzinę",
    ],
    "miesięcznie": [
        "/month",
        "monthly",
        "per month",
        "mies",
        "miesiąc",
        "miesięcznie",
        "msc",
    ],
    "rocznie": [
        "/year",
        "yearly",
        "annual",
        "annually",
        "per year",
        "rocznie",
        "rok",
    ],
}

TAX_KEYWORDS = {
    "brutto": ["brutto", "gross"],
    "netto": ["netto", "net", "+ vat", "+vat", "vat"],
}
EXCLUDED_CONTEXT_KEYWORDS = [
    "additional up to",
    "opportunity to earn",
    "participating in the company's activities",
    "referral bonus",
    "benefit",
    "allowance",
    "dodatkowo",
    "program poleceń",
]


def extract_salary_info(text: str) -> SalaryInfo:
    candidates = _salary_candidates(text)
    if not candidates:
        return SalaryInfo()

    candidate = max(candidates, key=_candidate_score)
    currency = _normalize_currency(candidate["currency"])
    amount_min = _parse_amount(candidate["amount_min"])
    amount_max = _parse_amount(candidate.get("amount_max")) or amount_min
    if currency not in SUPPORTED_CURRENCIES or amount_min is None:
        return SalaryInfo(original_text=candidate["context"])

    period = _infer_period(candidate["context"], amount_min)
    tax_type = _infer_tax_type(candidate["context"])
    salary = SalaryInfo(
        original_text=candidate["context"],
        currency=currency,
        amount_min=amount_min,
        amount_max=amount_max,
        period=period,
        tax_type=tax_type,
    )
    return normalize_salary_to_pln(salary)


def extract_salary_info_from_structured_data(
    job_posting: dict, context: str = ""
) -> SalaryInfo:
    if not isinstance(job_posting, dict):
        return SalaryInfo()

    for base_salary in _as_list(job_posting.get("baseSalary")):
        salary = _salary_from_base_salary(base_salary, context)
        if salary.amount_min is not None:
            return salary

    return SalaryInfo()


def normalize_salary_to_pln(salary: SalaryInfo) -> SalaryInfo:
    if salary.currency == UNKNOWN_VALUE or salary.amount_min is None:
        return salary

    rate, rate_date, rate_note = _exchange_rate_to_pln(salary.currency)
    if rate is None:
        return replace(
            salary,
            conversion_assumptions=(
                f"Nie przeliczono waluty. {rate_note}".strip()
            ),
        )

    amount_min_pln = salary.amount_min * rate
    amount_max_pln = (salary.amount_max or salary.amount_min) * rate
    monthly_min, hourly_min = _convert_period(amount_min_pln, salary.period)
    monthly_max, hourly_max = _convert_period(amount_max_pln, salary.period)

    assumptions = f"{HOURS_PER_MONTH} h/mies."
    if salary.currency == "PLN":
        assumptions = f"{assumptions}; PLN bez przeliczenia waluty"
    else:
        assumptions = (
            f"{assumptions}; kurs NBP tabela A {salary.currency}/PLN "
            f"{rate} z dnia {rate_date}"
        )

    return replace(
        salary,
        exchange_rate_to_pln=rate,
        exchange_rate_date=rate_date,
        pln_min_monthly=_round_money(monthly_min),
        pln_max_monthly=_round_money(monthly_max),
        pln_min_hourly=_round_money(hourly_min),
        pln_max_hourly=_round_money(hourly_max),
        conversion_assumptions=assumptions,
    )


def _salary_candidates(text: str) -> list[dict[str, str]]:
    lines = _candidate_lines(text)
    candidates: list[dict[str, str]] = []
    for line in lines:
        candidates.extend(_line_candidates(line))
    return candidates


def _line_candidates(line: str) -> list[dict[str, str]]:
    if _is_excluded_salary_context(line):
        return []

    patterns = [
        re.compile(
            rf"(?P<currency>{CURRENCY_PATTERN})\s*"
            rf"(?P<amount_min>{NUMBER_PATTERN})"
            rf"(?:\s*{RANGE_SEPARATOR_PATTERN}\s*"
            rf"(?:(?:{CURRENCY_PATTERN})\s*)?"
            rf"(?P<amount_max>{NUMBER_PATTERN}))?",
            flags=re.I,
        ),
        re.compile(
            rf"(?P<amount_min>{NUMBER_PATTERN})"
            rf"(?:\s*{RANGE_SEPARATOR_PATTERN}\s*"
            rf"(?P<amount_max>{NUMBER_PATTERN}))?"
            rf"\s*(?P<currency>{CURRENCY_PATTERN})",
            flags=re.I,
        ),
    ]

    candidates: list[dict[str, str]] = []
    for pattern in patterns:
        for match in pattern.finditer(line):
            currency = match.group("currency")
            amount_min = match.group("amount_min")
            amount_max = match.groupdict().get("amount_max")
            if not currency or not amount_min:
                continue
            candidates.append(
                {
                    "currency": currency,
                    "amount_min": amount_min,
                    "amount_max": amount_max or "",
                    "context": line,
                }
            )
    return candidates


def _salary_from_base_salary(base_salary: object, context: str) -> SalaryInfo:
    if not isinstance(base_salary, dict):
        return SalaryInfo()

    currency = _normalize_currency(str(base_salary.get("currency", "")))
    if currency not in SUPPORTED_CURRENCIES:
        return SalaryInfo()

    value_objects = _as_list(base_salary.get("value"))
    if not value_objects:
        value_objects = [base_salary]

    for value_object in value_objects:
        salary = _salary_from_quantitative_value(value_object, base_salary, currency, context)
        if salary.amount_min is not None:
            return normalize_salary_to_pln(salary)

    return SalaryInfo()


def _salary_from_quantitative_value(
    value_object: object, base_salary: dict, currency: str, context: str
) -> SalaryInfo:
    if isinstance(value_object, dict):
        amount_min = _parse_amount(
            value_object.get("minValue")
            or value_object.get("value")
            or base_salary.get("minValue")
            or base_salary.get("value")
        )
        amount_max = _parse_amount(
            value_object.get("maxValue")
            or value_object.get("value")
            or base_salary.get("maxValue")
            or amount_min
        )
        unit_text = str(value_object.get("unitText") or base_salary.get("unitText") or "")
    else:
        amount_min = _parse_amount(value_object)
        amount_max = amount_min
        unit_text = str(base_salary.get("unitText") or "")

    if amount_min is None:
        return SalaryInfo()

    period = _period_from_unit_text(unit_text) or _infer_period(context, amount_min)
    amount_max = amount_max or amount_min
    return SalaryInfo(
        original_text=_structured_salary_context(amount_min, amount_max, currency, period),
        currency=currency,
        amount_min=amount_min,
        amount_max=amount_max,
        period=period,
        tax_type=_infer_tax_type(context),
    )


def _structured_salary_context(
    amount_min: float, amount_max: float, currency: str, period: str
) -> str:
    amount = _format_amount(amount_min)
    if amount_max != amount_min:
        amount = f"{amount}-{_format_amount(amount_max)}"
    return f"Dane strukturalne baseSalary: {amount} {currency} {period}"


def _period_from_unit_text(unit_text: str) -> str:
    normalized = unit_text.strip().lower()
    if normalized in {"hour", "hours", "h"} or "godz" in normalized:
        return "godzinowo"
    if normalized in {"month", "months", "mth"} or "mies" in normalized:
        return "miesięcznie"
    if normalized in {"year", "years"} or "annual" in normalized or "rok" in normalized:
        return "rocznie"
    return ""


def _candidate_score(candidate: dict[str, str]) -> int:
    context = candidate["context"].lower()
    score = 0
    if any(keyword in context for keyword in SALARY_CONTEXT_KEYWORDS):
        score += 5
    if any(keyword in context for keywords in PERIOD_KEYWORDS.values() for keyword in keywords):
        score += 3
    if candidate.get("amount_max"):
        score += 2
    if _normalize_currency(candidate["currency"]) == "PLN":
        score += 1
    return score


def _candidate_lines(text: str) -> list[str]:
    normalized = re.sub(r"\s+", " ", text)
    chunks = re.split(r"(?<=[.;])\s+|\n+", normalized)
    return [
        chunk.strip()
        for chunk in chunks
        if 4 <= len(chunk.strip()) <= 500
        and re.search(CURRENCY_PATTERN, chunk, flags=re.I)
        and (
            any(keyword in chunk.lower() for keyword in SALARY_CONTEXT_KEYWORDS)
            or any(
                keyword in chunk.lower()
                for keywords in PERIOD_KEYWORDS.values()
                for keyword in keywords
            )
        )
    ]


def _normalize_currency(currency: str) -> str:
    cleaned = currency.strip().upper()
    if cleaned in {"ZŁ", "ZLOTY", "ZŁOTY"}:
        return "PLN"
    if cleaned == "€":
        return "EUR"
    if cleaned == "$":
        return "USD"
    return cleaned


def _parse_amount(value: object | None) -> float | None:
    if not value:
        return None

    if isinstance(value, int | float):
        return float(value)

    cleaned = str(value).replace("\u00a0", " ").replace(" ", "").strip()
    if not cleaned:
        return None

    if "," in cleaned and "." in cleaned:
        if cleaned.rfind(",") > cleaned.rfind("."):
            cleaned = cleaned.replace(".", "").replace(",", ".")
        else:
            cleaned = cleaned.replace(",", "")
    elif "," in cleaned:
        decimal_part = cleaned.rsplit(",", 1)[-1]
        cleaned = cleaned.replace(",", ".") if len(decimal_part) <= 2 else cleaned.replace(",", "")
    elif "." in cleaned:
        decimal_part = cleaned.rsplit(".", 1)[-1]
        if len(decimal_part) > 2:
            cleaned = cleaned.replace(".", "")

    try:
        return float(cleaned)
    except ValueError:
        return None


def _as_list(value: object) -> list[object]:
    if value is None:
        return []
    if isinstance(value, list):
        return value
    return [value]


def _is_excluded_salary_context(context: str) -> bool:
    lowered = context.lower()
    return any(keyword in lowered for keyword in EXCLUDED_CONTEXT_KEYWORDS)


def _infer_period(context: str, amount_min: float) -> str:
    lowered = context.lower()
    for period, keywords in PERIOD_KEYWORDS.items():
        if any(keyword in lowered for keyword in keywords):
            return period

    if amount_min < 1000:
        return "godzinowo"
    if amount_min >= 1000:
        return "miesięcznie"
    return UNKNOWN_VALUE


def _infer_tax_type(context: str) -> str:
    lowered = context.lower()
    for tax_type, keywords in TAX_KEYWORDS.items():
        if any(keyword in lowered for keyword in keywords):
            return tax_type
    return UNKNOWN_VALUE


def _exchange_rate_to_pln(currency: str) -> tuple[float | None, str, str]:
    if currency == "PLN":
        return 1.0, date.today().isoformat(), ""

    try:
        response = requests.get(
            NBP_RATE_URL.format(currency=currency),
            timeout=RATE_TIMEOUT_SECONDS,
        )
        response.raise_for_status()
        payload = response.json()
        rate = payload["rates"][0]
        return float(rate["mid"]), str(rate["effectiveDate"]), ""
    except Exception as exc:
        return None, "", f"Nie udało się pobrać kursu NBP dla {currency}: {exc}"


def _convert_period(amount_pln: float, period: str) -> tuple[float, float]:
    if period == "godzinowo":
        return amount_pln * HOURS_PER_MONTH, amount_pln
    if period == "rocznie":
        monthly = amount_pln / 12
        return monthly, monthly / HOURS_PER_MONTH

    monthly = amount_pln
    return monthly, monthly / HOURS_PER_MONTH


def _round_money(value: float) -> float:
    return round(value, 2)


def _format_amount(value: float) -> str:
    if value.is_integer():
        return str(int(value))
    return f"{value:.2f}".rstrip("0").rstrip(".")
