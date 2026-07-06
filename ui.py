from __future__ import annotations

from html import escape
from pathlib import Path
import sys

import streamlit as st

PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from job_offer_analyzer.excel_writer import (
    append_offer_to_workbook,
    refresh_cv_matches_in_workbook,
    refresh_offer_availability_in_workbook,
    refresh_offer_salaries_in_workbook,
)
from job_offer_analyzer.fetch_offer import OfferFetchError, fetch_offer_from_url
from job_offer_analyzer.models import OfferDraft, SalaryInfo, UNKNOWN_VALUE
from job_offer_analyzer.models import OfferRecord
from services.cv_matcher import CvProfileError, CvProfileNotFoundError


WORKBOOK_PATH = PROJECT_ROOT / "data" / "baza_ofert_pracy_QA.xlsx"
CV_PROFILE_PATH = PROJECT_ROOT / "data" / "private" / "cv_profile.yml"

FIELD_DEFAULTS = {
    "link": "",
    "company": "",
    "title": "",
    "location": UNKNOWN_VALUE,
    "rate_expectations": UNKNOWN_VALUE,
    "availability": "Dostępna",
    "status": "Nowa",
    "priority": UNKNOWN_VALUE,
    "technologies": UNKNOWN_VALUE,
    "work_mode": UNKNOWN_VALUE,
    "seniority": UNKNOWN_VALUE,
    "contract_type": UNKNOWN_VALUE,
    "salary_original_text": UNKNOWN_VALUE,
    "salary_currency": UNKNOWN_VALUE,
    "salary_amount_min": "",
    "salary_amount_max": "",
    "salary_period": UNKNOWN_VALUE,
    "salary_tax_type": UNKNOWN_VALUE,
    "salary_exchange_rate_to_pln": "",
    "salary_exchange_rate_date": "",
    "salary_pln_min_monthly": "",
    "salary_pln_max_monthly": "",
    "salary_pln_min_hourly": "",
    "salary_pln_max_hourly": "",
    "salary_conversion_assumptions": "",
    "must_have_summary": UNKNOWN_VALUE,
    "nice_to_have_summary": UNKNOWN_VALUE,
    "risks_notes": UNKNOWN_VALUE,
    "next_step": "Pobrać treść oferty i wykonać analizę pod CV",
    "source_preview": "",
    "fetch_message": "",
}

AVAILABILITY_OPTIONS = ["Dostępna", "Do sprawdzenia", "Niepewna", "Zamknięta"]
STATUS_OPTIONS = ["Nowa", "Do analizy", "Do poprawy CV", "Aplikować", "Odrzucona"]
PRIORITY_OPTIONS = [UNKNOWN_VALUE, "HIGH", "MEDIUM", "LOW"]
WORK_MODE_OPTIONS = [UNKNOWN_VALUE, "Remote", "Hybrid", "Office"]
SENIORITY_OPTIONS = [UNKNOWN_VALUE, "Intern", "Junior", "Mid", "Senior", "Experienced"]
SALARY_CURRENCY_OPTIONS = [UNKNOWN_VALUE, "PLN", "EUR", "USD"]
SALARY_PERIOD_OPTIONS = [UNKNOWN_VALUE, "godzinowo", "miesięcznie", "rocznie"]
SALARY_TAX_OPTIONS = [UNKNOWN_VALUE, "netto", "brutto"]


def main() -> None:
    st.set_page_config(
        page_title="Job Offer Analyzer",
        page_icon=None,
        layout="wide",
    )

    st.title("Job Offer Analyzer")

    _ensure_session_defaults()

    link_col, fetch_col, refresh_col, salary_refresh_col, cv_match_col = st.columns(
        [3.4, 1, 1.25, 1.35, 1.35]
    )
    with link_col:
        st.text_input("Link do oferty", key="link")
    with fetch_col:
        st.markdown("<div style='height: 1.75rem'></div>", unsafe_allow_html=True)
        fetch_clicked = st.button("Pobierz dane", use_container_width=True)
    with refresh_col:
        st.markdown("<div style='height: 1.75rem'></div>", unsafe_allow_html=True)
        refresh_clicked = st.button("Sprawdź oferty", use_container_width=True)
    with salary_refresh_col:
        st.markdown("<div style='height: 1.75rem'></div>", unsafe_allow_html=True)
        salary_refresh_clicked = st.button("Uzupełnij stawki", use_container_width=True)
    with cv_match_col:
        st.markdown("<div style='height: 1.75rem'></div>", unsafe_allow_html=True)
        cv_match_clicked = st.button("Analizuj pod CV", use_container_width=True)

    if fetch_clicked:
        _fetch_link_into_form()

    if refresh_clicked:
        _refresh_offer_availability()

    if salary_refresh_clicked:
        _refresh_offer_salaries()

    if cv_match_clicked:
        _refresh_cv_matches()

    if st.session_state["fetch_message"]:
        st.success(st.session_state["fetch_message"])

    if st.session_state["source_preview"]:
        with st.expander("Podgląd pobranej treści"):
            st.text_area(
                "Treść oferty",
                value=st.session_state["source_preview"],
                height=260,
                disabled=True,
            )

    with st.form("offer_form", clear_on_submit=False):
        left, right = st.columns([1.1, 1])

        with left:
            company = st.text_input("Firma", key="company")
            title = st.text_input("Stanowisko", key="title")
            location = st.text_input("Lokalizacja", key="location")
            rate_expectations = st.text_input(
                "Stawka / oczekiwania (PLN)", key="rate_expectations"
            )

        with right:
            availability = st.selectbox(
                "Dostępność",
                AVAILABILITY_OPTIONS,
                key="availability",
            )
            status = st.selectbox(
                "Status",
                STATUS_OPTIONS,
                key="status",
            )
            priority = st.selectbox(
                "Priorytet",
                PRIORITY_OPTIONS,
                key="priority",
            )
            work_mode = st.selectbox(
                "Tryb",
                WORK_MODE_OPTIONS,
                key="work_mode",
            )
            seniority = st.selectbox(
                "Poziom",
                SENIORITY_OPTIONS,
                key="seniority",
            )

        contract_type = st.text_input("Forma umowy", key="contract_type")
        technologies = st.text_input("Technologie", key="technologies")

        with st.expander("Wynagrodzenie", expanded=True):
            st.text_area(
                "Stawka źródłowa",
                height=70,
                key="salary_original_text",
            )
            salary_col_1, salary_col_2, salary_col_3, salary_col_4 = st.columns(4)
            with salary_col_1:
                st.selectbox(
                    "Waluta",
                    SALARY_CURRENCY_OPTIONS,
                    key="salary_currency",
                )
            with salary_col_2:
                st.text_input("Stawka min", key="salary_amount_min")
            with salary_col_3:
                st.text_input("Stawka max", key="salary_amount_max")
            with salary_col_4:
                st.selectbox(
                    "Okres stawki",
                    SALARY_PERIOD_OPTIONS,
                    key="salary_period",
                )

            salary_col_5, salary_col_6, salary_col_7 = st.columns(3)
            with salary_col_5:
                st.selectbox(
                    "Brutto/netto",
                    SALARY_TAX_OPTIONS,
                    key="salary_tax_type",
                )
            with salary_col_6:
                st.text_input("Kurs waluty", key="salary_exchange_rate_to_pln")
            with salary_col_7:
                st.text_input("Data kursu", key="salary_exchange_rate_date")

            salary_col_8, salary_col_9, salary_col_10, salary_col_11 = st.columns(4)
            with salary_col_8:
                st.text_input("PLN min miesięcznie", key="salary_pln_min_monthly")
            with salary_col_9:
                st.text_input("PLN max miesięcznie", key="salary_pln_max_monthly")
            with salary_col_10:
                st.text_input("PLN min godzinowo", key="salary_pln_min_hourly")
            with salary_col_11:
                st.text_input("PLN max godzinowo", key="salary_pln_max_hourly")

            st.text_input(
                "Założenia przeliczenia",
                key="salary_conversion_assumptions",
            )

        must_have_summary = st.text_area(
            "Must-have skrót", height=90, key="must_have_summary"
        )
        nice_to_have_summary = st.text_area(
            "Nice-to-have skrót", height=90, key="nice_to_have_summary"
        )
        risks_notes = st.text_area("Ryzyka / uwagi", height=90, key="risks_notes")
        next_step = st.text_input(
            "Następny krok",
            key="next_step",
        )

        submitted = st.form_submit_button("Zapisz ofertę")

    if not submitted:
        return

    link = st.session_state["link"]
    errors = _validate_required_fields(company=company, title=title, link=link)
    if errors:
        for error in errors:
            st.error(error)
        return

    offer = OfferRecord(
        company=company.strip(),
        title=title.strip(),
        link=link.strip(),
        availability=availability,
        status=status,
        priority=priority,
        technologies=technologies.strip() or UNKNOWN_VALUE,
        work_mode=work_mode,
        location=location.strip() or UNKNOWN_VALUE,
        contract_type=contract_type.strip() or UNKNOWN_VALUE,
        rate_expectations=rate_expectations.strip() or UNKNOWN_VALUE,
        salary=_salary_from_session(),
        seniority=seniority,
        must_have_summary=must_have_summary.strip() or UNKNOWN_VALUE,
        nice_to_have_summary=nice_to_have_summary.strip() or UNKNOWN_VALUE,
        risks_notes=risks_notes.strip() or UNKNOWN_VALUE,
        next_step=next_step.strip() or "Pobrać treść oferty i wykonać analizę pod CV",
        source=link.strip(),
    )

    try:
        offer_id = append_offer_to_workbook(WORKBOOK_PATH, offer)
    except Exception as exc:
        st.error(f"Nie udało się zapisać oferty: {exc}")
        return

    st.success(f"Zapisano ofertę {offer_id}")
    st.caption(str(WORKBOOK_PATH))


def _ensure_session_defaults() -> None:
    for key, value in FIELD_DEFAULTS.items():
        if key not in st.session_state:
            st.session_state[key] = value
        elif st.session_state[key] == "TBD":
            st.session_state[key] = UNKNOWN_VALUE

    _ensure_allowed_value("priority", PRIORITY_OPTIONS)
    _ensure_allowed_value("work_mode", WORK_MODE_OPTIONS)
    _ensure_allowed_value("seniority", SENIORITY_OPTIONS)
    _ensure_allowed_value("salary_currency", SALARY_CURRENCY_OPTIONS)
    _ensure_allowed_value("salary_period", SALARY_PERIOD_OPTIONS)
    _ensure_allowed_value("salary_tax_type", SALARY_TAX_OPTIONS)


def _ensure_allowed_value(key: str, allowed_values: list[str]) -> None:
    if st.session_state[key] not in allowed_values:
        st.session_state[key] = allowed_values[0]


def _fetch_link_into_form() -> None:
    try:
        with st.spinner("Pobieram ofertę..."):
            draft = fetch_offer_from_url(st.session_state["link"])
    except OfferFetchError as exc:
        st.session_state["fetch_message"] = ""
        st.error(str(exc))
        return
    except Exception as exc:
        st.session_state["fetch_message"] = ""
        st.error(f"Nie udało się odczytać oferty: {exc}")
        return

    _apply_draft(draft)
    st.session_state["fetch_message"] = "Pobrano dane z linku."


def _refresh_offer_availability() -> None:
    try:
        with st.spinner("Sprawdzam dostępność ofert..."):
            summary = refresh_offer_availability_in_workbook(WORKBOOK_PATH)
    except PermissionError as exc:
        st.error(str(exc))
        return
    except Exception as exc:
        st.error(f"Nie udało się sprawdzić dostępności ofert: {exc}")
        return

    st.success(
        "Sprawdzono "
        f"{summary.checked_count} ofert: "
        f"{summary.available_count} dostępnych, "
        f"{summary.closed_count} zamkniętych, "
        f"{summary.uncertain_count} niepewnych. "
        f"Zmieniono status dla {summary.changed_count}."
    )

    if summary.results:
        _show_wrapped_results_table(
            [
                {
                    "ID": result.offer_id,
                    "Firma": result.company,
                    "Stanowisko": result.title,
                    "Poprzednio": result.previous_availability,
                    "Teraz": result.availability,
                    "Zmiana": "Tak" if result.changed else "Nie",
                    "Notatka": result.note,
                }
                for result in summary.results
            ]
        )


def _refresh_offer_salaries() -> None:
    try:
        with st.spinner("Uzupełniam stawki ofert..."):
            summary = refresh_offer_salaries_in_workbook(WORKBOOK_PATH)
    except PermissionError as exc:
        st.error(str(exc))
        return
    except Exception as exc:
        st.error(f"Nie udało się uzupełnić stawek ofert: {exc}")
        return

    st.success(
        "Sprawdzono "
        f"{summary.checked_count} ofert. "
        f"Uzupełniono stawki dla {summary.updated_count}, "
        f"nie znaleziono stawek dla {summary.missing_count}, "
        f"błędów pobrania: {summary.failed_count}."
    )

    if summary.results:
        _show_wrapped_results_table(
            [
                {
                    "ID": result.offer_id,
                    "Firma": result.company,
                    "Stanowisko": result.title,
                    "Stawka": result.salary_display,
                    "Aktualizacja": "Tak" if result.updated else "Nie",
                    "Notatka": result.note,
                }
                for result in summary.results
            ]
        )


def _refresh_cv_matches() -> None:
    selected_link = st.session_state["link"].strip() or None
    try:
        with st.spinner("Analizuję dopasowanie ofert do profilu..."):
            summary = refresh_cv_matches_in_workbook(
                WORKBOOK_PATH,
                CV_PROFILE_PATH,
                selected_link=selected_link,
            )
    except CvProfileNotFoundError as exc:
        st.error(str(exc))
        return
    except CvProfileError as exc:
        st.error(f"Profil CV ma nieprawidłowy format: {exc}")
        return
    except PermissionError as exc:
        st.error(str(exc))
        return
    except Exception as exc:
        st.error(f"Nie udało się przeanalizować dopasowania do CV: {exc}")
        return

    st.success(
        "Przeanalizowano "
        f"{summary.checked_count} ofert. "
        f"Zaktualizowano {summary.updated_count}, "
        f"błędów: {summary.failed_count}."
    )

    if summary.results:
        _show_wrapped_results_table(
            [
                {
                    "ID": result.offer_id,
                    "Firma": result.company,
                    "Stanowisko": result.title,
                    "Wynik": f"{result.match_score}%",
                    "Priorytet": result.priority,
                    "Dopasowane": "; ".join(result.matched_skills),
                    "Braki": "; ".join(result.missing_skills),
                    "Notatka": result.note,
                }
                for result in summary.results
            ]
        )


def _show_wrapped_results_table(rows: list[dict[str, object]]) -> None:
    if not rows:
        return

    columns = list(rows[0].keys())
    header_html = "".join(f"<th>{escape(column)}</th>" for column in columns)
    rows_html = []
    for row in rows:
        cells = []
        for column in columns:
            css_class = " class='note-cell'" if column == "Notatka" else ""
            value = "" if row.get(column) is None else str(row.get(column))
            cells.append(f"<td{css_class}>{escape(value)}</td>")
        rows_html.append(f"<tr>{''.join(cells)}</tr>")

    st.markdown(
        f"""
        <style>
            .wrapped-results {{
                overflow-x: auto;
                margin-top: 0.75rem;
            }}
            .wrapped-results table {{
                width: 100%;
                min-width: 900px;
                border-collapse: collapse;
                font-size: 0.875rem;
            }}
            .wrapped-results th {{
                background: #1f4e78;
                color: #ffffff;
                border: 1px solid #b7c9e2;
                padding: 0.45rem 0.6rem;
                text-align: left;
                vertical-align: top;
                white-space: nowrap;
            }}
            .wrapped-results td {{
                border: 1px solid #d9e2f3;
                padding: 0.45rem 0.6rem;
                vertical-align: top;
                white-space: nowrap;
            }}
            .wrapped-results .note-cell {{
                min-width: 360px;
                max-width: 640px;
                white-space: normal;
                overflow-wrap: anywhere;
                line-height: 1.35;
            }}
        </style>
        <div class="wrapped-results">
            <table>
                <thead><tr>{header_html}</tr></thead>
                <tbody>{''.join(rows_html)}</tbody>
            </table>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _apply_draft(draft: OfferDraft) -> None:
    _set_if_useful("company", draft.company)
    _set_if_useful("title", draft.title)
    _set_if_useful("location", draft.location)
    _set_if_useful("work_mode", draft.work_mode, allowed_values=WORK_MODE_OPTIONS)
    _set_if_useful("contract_type", draft.contract_type)
    _set_if_useful("technologies", draft.technologies)
    _set_if_useful("rate_expectations", draft.rate_expectations)
    _set_if_useful("seniority", draft.seniority, allowed_values=SENIORITY_OPTIONS)
    _apply_salary(draft.salary)
    _set_if_useful("must_have_summary", draft.must_have_summary)
    _set_if_useful("nice_to_have_summary", draft.nice_to_have_summary)
    _set_if_useful("risks_notes", draft.risks_notes)

    st.session_state["availability"] = "Dostępna"
    st.session_state["status"] = "Do analizy"
    st.session_state["next_step"] = "Porównać ofertę z CV i przygotować aplikację"
    st.session_state["source_preview"] = draft.source_text


def _apply_salary(salary: SalaryInfo) -> None:
    _set_if_useful("salary_original_text", salary.original_text)
    _set_if_useful(
        "salary_currency",
        salary.currency,
        allowed_values=SALARY_CURRENCY_OPTIONS,
    )
    _set_if_useful(
        "salary_period",
        salary.period,
        allowed_values=SALARY_PERIOD_OPTIONS,
    )
    _set_if_useful(
        "salary_tax_type",
        salary.tax_type,
        allowed_values=SALARY_TAX_OPTIONS,
    )
    _set_text_value("salary_amount_min", salary.amount_min)
    _set_text_value("salary_amount_max", salary.amount_max)
    _set_text_value("salary_exchange_rate_to_pln", salary.exchange_rate_to_pln)
    _set_text_value("salary_exchange_rate_date", salary.exchange_rate_date)
    _set_text_value("salary_pln_min_monthly", salary.pln_min_monthly)
    _set_text_value("salary_pln_max_monthly", salary.pln_max_monthly)
    _set_text_value("salary_pln_min_hourly", salary.pln_min_hourly)
    _set_text_value("salary_pln_max_hourly", salary.pln_max_hourly)
    _set_text_value("salary_conversion_assumptions", salary.conversion_assumptions)


def _salary_from_session() -> SalaryInfo:
    return SalaryInfo(
        original_text=_text_or_unknown("salary_original_text"),
        currency=st.session_state["salary_currency"],
        amount_min=_parse_float_field("salary_amount_min"),
        amount_max=_parse_float_field("salary_amount_max"),
        period=st.session_state["salary_period"],
        tax_type=st.session_state["salary_tax_type"],
        exchange_rate_to_pln=_parse_float_field("salary_exchange_rate_to_pln"),
        exchange_rate_date=st.session_state["salary_exchange_rate_date"].strip(),
        pln_min_monthly=_parse_float_field("salary_pln_min_monthly"),
        pln_max_monthly=_parse_float_field("salary_pln_max_monthly"),
        pln_min_hourly=_parse_float_field("salary_pln_min_hourly"),
        pln_max_hourly=_parse_float_field("salary_pln_max_hourly"),
        conversion_assumptions=st.session_state[
            "salary_conversion_assumptions"
        ].strip(),
    )


def _text_or_unknown(key: str) -> str:
    value = str(st.session_state[key]).strip()
    return value or UNKNOWN_VALUE


def _parse_float_field(key: str) -> float | None:
    value = str(st.session_state[key]).strip().replace(",", ".")
    if not value:
        return None
    try:
        return float(value)
    except ValueError:
        return None


def _set_text_value(key: str, value: object) -> None:
    if value is None:
        return
    text = str(value).strip()
    if text:
        st.session_state[key] = text


def _set_if_useful(
    key: str, value: str, allowed_values: list[str] | None = None
) -> None:
    cleaned_value = value.strip()
    if not cleaned_value or cleaned_value in {"TBD", UNKNOWN_VALUE}:
        return

    if allowed_values is not None and cleaned_value not in allowed_values:
        return

    st.session_state[key] = cleaned_value


def _validate_required_fields(company: str, title: str, link: str) -> list[str]:
    errors: list[str] = []
    if not company.strip():
        errors.append("Uzupełnij pole: Firma")
    if not title.strip():
        errors.append("Uzupełnij pole: Stanowisko")
    if not link.strip():
        errors.append("Uzupełnij pole: Link do oferty")
    return errors


if __name__ == "__main__":
    main()
