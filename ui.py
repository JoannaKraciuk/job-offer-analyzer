from __future__ import annotations

from pathlib import Path
import sys

import streamlit as st

PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from job_offer_analyzer.excel_writer import (
    append_offer_to_workbook,
    refresh_offer_availability_in_workbook,
)
from job_offer_analyzer.fetch_offer import OfferFetchError, fetch_offer_from_url
from job_offer_analyzer.models import OfferDraft, UNKNOWN_VALUE
from job_offer_analyzer.models import OfferRecord


WORKBOOK_PATH = PROJECT_ROOT / "data" / "baza_ofert_pracy_QA.xlsx"

FIELD_DEFAULTS = {
    "link": "",
    "company": "",
    "title": "",
    "location": UNKNOWN_VALUE,
    "rate_expectations": UNKNOWN_VALUE,
    "availability": "Dostępna",
    "status": "Nowa",
    "priority": UNKNOWN_VALUE,
    "work_mode": UNKNOWN_VALUE,
    "seniority": UNKNOWN_VALUE,
    "contract_type": UNKNOWN_VALUE,
    "must_have_summary": UNKNOWN_VALUE,
    "nice_to_have_summary": UNKNOWN_VALUE,
    "risks_notes": UNKNOWN_VALUE,
    "next_step": "Pobrać treść oferty i wykonać analizę pod CV",
    "source_preview": "",
    "fetch_message": "",
}

AVAILABILITY_OPTIONS = ["Dostępna", "Do sprawdzenia", "Niepewna", "Zamknięta"]
STATUS_OPTIONS = ["Nowa", "Do analizy", "Do poprawy CV", "Aplikować", "Odrzucona"]
PRIORITY_OPTIONS = [UNKNOWN_VALUE, "Wysoki", "Średni", "Niski"]
WORK_MODE_OPTIONS = [UNKNOWN_VALUE, "Remote", "Hybrid", "Office"]
SENIORITY_OPTIONS = [UNKNOWN_VALUE, "Intern", "Junior", "Mid", "Senior", "Experienced"]


def main() -> None:
    st.set_page_config(
        page_title="Job Offer Analyzer",
        page_icon=None,
        layout="wide",
    )

    st.title("Job Offer Analyzer")

    _ensure_session_defaults()

    link_col, fetch_col, refresh_col = st.columns([4, 1, 1.4])
    with link_col:
        st.text_input("Link do oferty", key="link")
    with fetch_col:
        st.markdown("<div style='height: 1.75rem'></div>", unsafe_allow_html=True)
        fetch_clicked = st.button("Pobierz dane", use_container_width=True)
    with refresh_col:
        st.markdown("<div style='height: 1.75rem'></div>", unsafe_allow_html=True)
        refresh_clicked = st.button("Sprawdź oferty", use_container_width=True)

    if fetch_clicked:
        _fetch_link_into_form()

    if refresh_clicked:
        _refresh_offer_availability()

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
        work_mode=work_mode,
        location=location.strip() or UNKNOWN_VALUE,
        contract_type=contract_type.strip() or UNKNOWN_VALUE,
        rate_expectations=rate_expectations.strip() or UNKNOWN_VALUE,
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
        st.dataframe(
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
            ],
            use_container_width=True,
        )


def _apply_draft(draft: OfferDraft) -> None:
    _set_if_useful("company", draft.company)
    _set_if_useful("title", draft.title)
    _set_if_useful("location", draft.location)
    _set_if_useful("work_mode", draft.work_mode, allowed_values=WORK_MODE_OPTIONS)
    _set_if_useful("contract_type", draft.contract_type)
    _set_if_useful("rate_expectations", draft.rate_expectations)
    _set_if_useful("seniority", draft.seniority, allowed_values=SENIORITY_OPTIONS)
    _set_if_useful("must_have_summary", draft.must_have_summary)
    _set_if_useful("nice_to_have_summary", draft.nice_to_have_summary)
    _set_if_useful("risks_notes", draft.risks_notes)

    st.session_state["availability"] = "Dostępna"
    st.session_state["status"] = "Do analizy"
    st.session_state["next_step"] = "Porównać ofertę z CV i przygotować aplikację"
    st.session_state["source_preview"] = draft.source_text


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
