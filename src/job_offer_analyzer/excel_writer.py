from __future__ import annotations

from copy import copy
from datetime import date
from pathlib import Path
import re
import unicodedata

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from job_offer_analyzer.models import (
    AvailabilityRefreshRow,
    AvailabilityRefreshSummary,
    HistoryRecord,
    OfferRecord,
    SalaryInfo,
    SalaryRefreshRow,
    SalaryRefreshSummary,
    UNKNOWN_VALUE,
)
from job_offer_analyzer.offer_availability import check_offer_availability


OFFERS_SHEET = "Oferty"
HISTORY_SHEET = "Historia_Sprawdzeń"
QUESTIONS_SHEET = "Pytania_Formularzy"
DASHBOARD_SHEET = "Dashboard"

OFFER_ID_PATTERN = re.compile(r"^JOB-(\d+)$")
POLISH_TRANSLATION = str.maketrans(
    "ąćęłńóśźżĄĆĘŁŃÓŚŹŻ",
    "acelnoszzACELNOSZZ",
)
HEADER_ALIASES = {
    "Forma umowy": ["Forma"],
    "Stawka / oczekiwania (PLN)": [
        "Stawka / oczekiwania",
        "Stawka / oczekiwania (USD)",
    ],
}
OFFER_HEADER_RENAMES = {
    "Dostepnosc": "Dostępność",
    "Forma": "Forma umowy",
    "Stawka / oczekiwania": "Stawka / oczekiwania (PLN)",
    "Must-have skrot": "Must-have skrót",
    "Nice-to-have skrot": "Nice-to-have skrót",
    "Nastepny krok": "Następny krok",
    "Zrodlo": "Źródło",
}
HISTORY_HEADER_RENAMES = {
    "Zrodlo": "Źródło",
}
DASHBOARD_TEXT_RENAMES = {
    "Baza ofert pracy - dashboard": "Baza ofert pracy - dashboard",
    "Wartosc": "Wartość",
    "Najblizsze dzialania": "Najbliższe działania",
    "Dostepne": "Dostępne",
    "Nastepny krok": "Następny krok",
    "Wymagaja sprawdzenia >14 dni": "Wymagają sprawdzenia >14 dni",
}
VALUE_RENAMES = {
    "TBD": UNKNOWN_VALUE,
    "Dostepna": "Dostępna",
    "Zamknieta": "Zamknięta",
    "Sredni": "Średni",
    "Pobrac tresc oferty i wykonac analize pod CV": (
        "Pobrać treść oferty i wykonać analizę pod CV"
    ),
    "Porownac z CV i przygotowac odpowiedzi do formularza": (
        "Porównać z CV i przygotować odpowiedzi do formularza"
    ),
}
ACTIVE_DASHBOARD_EXCLUDED_STATUSES = {"Aplikowano", "Odrzucona"}
ACTIVE_DASHBOARD_EXCLUDED_AVAILABILITY = {"Zamknięta", "Zamknieta"}
SALARY_HEADERS = [
    "Stawka źródłowa",
    "Waluta",
    "Stawka min",
    "Stawka max",
    "Okres stawki",
    "Brutto/netto",
    "Kurs waluty",
    "Data kursu",
    "PLN min miesięcznie",
    "PLN max miesięcznie",
    "PLN min godzinowo",
    "PLN max godzinowo",
    "Założenia przeliczenia",
]


def append_offer_to_workbook(workbook_path: Path, offer: OfferRecord) -> str:
    workbook_path = Path(workbook_path)
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook does not exist: {workbook_path}")
    _ensure_workbook_writable(workbook_path)

    workbook = load_workbook(workbook_path)
    offers_sheet = _get_sheet(workbook, OFFERS_SHEET)
    history_sheet = _get_sheet(workbook, HISTORY_SHEET)
    questions_sheet = _get_sheet(workbook, QUESTIONS_SHEET)
    dashboard_sheet = _get_sheet(workbook, DASHBOARD_SHEET)
    _normalize_sheet_headers(offers_sheet, OFFER_HEADER_RENAMES)
    _normalize_sheet_headers(history_sheet, HISTORY_HEADER_RENAMES)
    _ensure_sheet_headers(offers_sheet, SALARY_HEADERS)
    _normalize_sheet_values(offers_sheet)
    _normalize_sheet_values(history_sheet)
    _normalize_dashboard_sheet(dashboard_sheet)

    offer_id = _next_offer_id(offers_sheet)
    _append_offer_row(offers_sheet, offer_id, offer)
    _append_history_row(history_sheet, _history_from_offer(offer_id, offer))
    _append_question_placeholder_row(questions_sheet, offer_id, offer)
    _refresh_dashboard_actions(dashboard_sheet, offers_sheet)

    workbook.save(workbook_path)
    return offer_id


def refresh_offer_availability_in_workbook(
    workbook_path: Path,
) -> AvailabilityRefreshSummary:
    workbook_path = Path(workbook_path)
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook does not exist: {workbook_path}")
    _ensure_workbook_writable(workbook_path)

    workbook = load_workbook(workbook_path)
    offers_sheet = _get_sheet(workbook, OFFERS_SHEET)
    history_sheet = _get_sheet(workbook, HISTORY_SHEET)
    dashboard_sheet = _get_sheet(workbook, DASHBOARD_SHEET)
    _normalize_sheet_headers(offers_sheet, OFFER_HEADER_RENAMES)
    _normalize_sheet_headers(history_sheet, HISTORY_HEADER_RENAMES)
    _ensure_sheet_headers(offers_sheet, SALARY_HEADERS)
    _normalize_sheet_values(offers_sheet)
    _normalize_sheet_values(history_sheet)
    _normalize_dashboard_sheet(dashboard_sheet)

    headers = _header_positions(offers_sheet)
    results: list[AvailabilityRefreshRow] = []
    today = date.today()

    for row in range(2, offers_sheet.max_row + 1):
        offer_id = _cell_by_header(offers_sheet, row, headers, "ID")
        link = _cell_by_header(offers_sheet, row, headers, "Link")
        if not offer_id or not link:
            continue

        company = _cell_by_header(offers_sheet, row, headers, "Firma") or ""
        title = _cell_by_header(offers_sheet, row, headers, "Stanowisko") or ""
        previous_availability = (
            _cell_by_header(offers_sheet, row, headers, "Dostępność") or UNKNOWN_VALUE
        )
        availability_result = check_offer_availability(str(link))

        _set_cell_by_header(
            offers_sheet, row, headers, "Dostępność", availability_result.availability
        )
        _set_cell_by_header(offers_sheet, row, headers, "Ostatnio sprawdzono", today)
        _set_cell_by_header(offers_sheet, row, headers, "Dni od sprawdzenia", 0)

        changed = previous_availability != availability_result.availability
        results.append(
            AvailabilityRefreshRow(
                offer_id=str(offer_id),
                company=str(company),
                title=str(title),
                link=str(link),
                previous_availability=str(previous_availability),
                availability=availability_result.availability,
                note=availability_result.note,
                changed=changed,
            )
        )
        _append_history_row(
            history_sheet,
            HistoryRecord(
                checked_at=today,
                offer_id=str(offer_id),
                company=str(company),
                title=str(title),
                link=str(link),
                result=availability_result.availability,
                checked_scope="Automatyczne sprawdzenie dostępności",
                note=(
                    f"Poprzednia dostępność: {previous_availability}. "
                    f"{availability_result.note}"
                ),
                source=availability_result.final_url or str(link),
            ),
        )

    _refresh_dashboard_actions(dashboard_sheet, offers_sheet)
    workbook.save(workbook_path)

    return AvailabilityRefreshSummary(
        checked_count=len(results),
        available_count=sum(result.availability == "Dostępna" for result in results),
        closed_count=sum(result.availability == "Zamknięta" for result in results),
        uncertain_count=sum(result.availability == "Niepewna" for result in results),
        changed_count=sum(result.changed for result in results),
        results=results,
    )


def refresh_offer_salaries_in_workbook(workbook_path: Path) -> SalaryRefreshSummary:
    from job_offer_analyzer.fetch_offer import fetch_offer_from_url

    workbook_path = Path(workbook_path)
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook does not exist: {workbook_path}")
    _ensure_workbook_writable(workbook_path)

    workbook = load_workbook(workbook_path)
    offers_sheet = _get_sheet(workbook, OFFERS_SHEET)
    dashboard_sheet = _get_sheet(workbook, DASHBOARD_SHEET)
    _normalize_sheet_headers(offers_sheet, OFFER_HEADER_RENAMES)
    _ensure_sheet_headers(offers_sheet, SALARY_HEADERS)
    _normalize_sheet_values(offers_sheet)
    _normalize_dashboard_sheet(dashboard_sheet)

    headers = _header_positions(offers_sheet)
    results: list[SalaryRefreshRow] = []

    for row in range(2, offers_sheet.max_row + 1):
        offer_id = _cell_by_header(offers_sheet, row, headers, "ID")
        link = _cell_by_header(offers_sheet, row, headers, "Link")
        if not offer_id or not link:
            continue

        company = _cell_by_header(offers_sheet, row, headers, "Firma") or ""
        title = _cell_by_header(offers_sheet, row, headers, "Stanowisko") or ""

        try:
            draft = fetch_offer_from_url(str(link))
        except Exception as exc:
            results.append(
                SalaryRefreshRow(
                    offer_id=str(offer_id),
                    company=str(company),
                    title=str(title),
                    link=str(link),
                    salary_display=UNKNOWN_VALUE,
                    updated=False,
                    note=f"Nie udało się pobrać oferty: {exc}",
                )
            )
            continue

        salary = draft.salary
        if not _has_salary_values(salary):
            results.append(
                SalaryRefreshRow(
                    offer_id=str(offer_id),
                    company=str(company),
                    title=str(title),
                    link=str(link),
                    salary_display=UNKNOWN_VALUE,
                    updated=False,
                    note="Nie znaleziono widełek płacowych w danych oferty.",
                )
            )
            continue

        values = {"Stawka / oczekiwania (PLN)": salary.display_value}
        values.update(_salary_values(salary))
        _write_values(offers_sheet, row, headers, values)

        results.append(
            SalaryRefreshRow(
                offer_id=str(offer_id),
                company=str(company),
                title=str(title),
                link=str(link),
                salary_display=salary.display_value,
                updated=True,
                note="Uzupełniono stawki z linku oferty.",
            )
        )

    _refresh_dashboard_actions(dashboard_sheet, offers_sheet)
    workbook.save(workbook_path)

    return SalaryRefreshSummary(
        checked_count=len(results),
        updated_count=sum(result.updated for result in results),
        missing_count=sum(
            not result.updated and "Nie znaleziono" in result.note for result in results
        ),
        failed_count=sum(
            not result.updated and "Nie udało się" in result.note for result in results
        ),
        results=results,
    )


def _append_offer_row(sheet: Worksheet, offer_id: str, offer: OfferRecord) -> int:
    headers = _header_positions(sheet)
    row = _first_empty_row(sheet)
    _copy_row_style(sheet, source_row=2, target_row=row)

    source = offer.source or offer.link
    values = {
        "ID": offer_id,
        "Data dodania": offer.added_at,
        "Firma": offer.company,
        "Stanowisko": offer.title,
        "Link": offer.link,
        "Ostatnio sprawdzono": offer.last_checked_at,
        "Dostępność": offer.availability,
        "Status": offer.status,
        "Dopasowanie do CV": offer.cv_match,
        "Priorytet": offer.priority,
        "Tryb": offer.work_mode,
        "Lokalizacja": offer.location,
        "Forma umowy": offer.contract_type,
        "Stawka / oczekiwania (PLN)": offer.rate_expectations,
        "Poziom": offer.seniority,
        "Dni od sprawdzenia": offer.days_since_check,
        "Must-have skrót": offer.must_have_summary,
        "Nice-to-have skrót": offer.nice_to_have_summary,
        "Ryzyka / uwagi": offer.risks_notes,
        "Następny krok": offer.next_step,
        "Źródło": source,
    }
    values.update(_salary_values(offer.salary))
    _write_values(sheet, row, headers, values)
    return row


def _append_history_row(sheet: Worksheet, history: HistoryRecord) -> int:
    headers = _header_positions(sheet)
    row = _first_empty_row(sheet)
    _copy_row_style(sheet, source_row=2, target_row=row)

    values = {
        "Data sprawdzenia": history.checked_at,
        "ID oferty": history.offer_id,
        "Firma": history.company,
        "Stanowisko": history.title,
        "Link": history.link,
        "Wynik": history.result,
        "Co sprawdzono": history.checked_scope,
        "Notatka": history.note,
        "Źródło": history.source,
    }
    _write_values(sheet, row, headers, values)
    return row


def _append_question_placeholder_row(
    sheet: Worksheet, offer_id: str, offer: OfferRecord
) -> int:
    headers = _header_positions(sheet)
    if _question_row_exists(sheet, headers, offer_id):
        return 0

    row = _first_empty_row(sheet)
    _copy_row_style(sheet, source_row=2, target_row=row)

    values = {
        "ID oferty": offer_id,
        "Firma": offer.company,
        "Stanowisko": offer.title,
        "Pytanie z formularza": "Nie odczytano pytań z formularza aplikacyjnego",
        "Wymagane?": "Do sprawdzenia",
        "Szkic odpowiedzi": None,
        "Status odpowiedzi": "Do sprawdzenia",
        "Uwagi": (
            "Uzupełnić po wejściu w formularz aplikacyjny. "
            "Automatyczne pobranie opisu oferty nie zawiera pytań formularza."
        ),
    }
    _write_values(sheet, row, headers, values)
    return row


def _question_row_exists(
    sheet: Worksheet, headers: dict[str, int], offer_id: str
) -> bool:
    id_column = _column_for_header(headers, "ID oferty")
    if id_column is None:
        return False

    for row in range(2, sheet.max_row + 1):
        if sheet.cell(row=row, column=id_column).value == offer_id:
            return True
    return False


def _history_from_offer(offer_id: str, offer: OfferRecord) -> HistoryRecord:
    return HistoryRecord(
        checked_at=offer.last_checked_at,
        offer_id=offer_id,
        company=offer.company,
        title=offer.title,
        link=offer.link,
        result=offer.availability,
        checked_scope="Ręczny wpis testowy",
        note="Dodano testową ofertę przez pierwszy moduł zapisu do Excela.",
        source=offer.source or offer.link,
    )


def _get_sheet(workbook, sheet_name: str) -> Worksheet:
    wanted_name = _text_key(sheet_name)
    for existing_name in workbook.sheetnames:
        if _text_key(existing_name) == wanted_name:
            return workbook[existing_name]

    raise KeyError(f"Nie znaleziono arkusza: {sheet_name}")


def _header_positions(sheet: Worksheet) -> dict[str, int]:
    headers: dict[str, int] = {}
    for column in range(1, sheet.max_column + 1):
        value = sheet.cell(row=1, column=column).value
        if value:
            header = str(value).strip()
            headers[header] = column
            headers[_text_key(header)] = column
    return headers


def _normalize_sheet_headers(sheet: Worksheet, renames: dict[str, str]) -> None:
    for column in range(1, sheet.max_column + 1):
        value = sheet.cell(row=1, column=column).value
        if isinstance(value, str) and value in renames:
            sheet.cell(row=1, column=column, value=renames[value])


def _ensure_sheet_headers(sheet: Worksheet, headers: list[str]) -> None:
    existing_headers = _header_positions(sheet)
    for header in headers:
        if _column_for_header(existing_headers, header) is not None:
            continue

        target_column = sheet.max_column + 1
        source_cell = sheet.cell(row=1, column=target_column - 1)
        target_cell = sheet.cell(row=1, column=target_column, value=header)
        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.fill = copy(source_cell.fill)
            target_cell.border = copy(source_cell.border)
            target_cell.alignment = copy(source_cell.alignment)
            target_cell.number_format = source_cell.number_format
            target_cell.protection = copy(source_cell.protection)
        existing_headers[header] = target_column
        existing_headers[_text_key(header)] = target_column


def _normalize_sheet_values(sheet: Worksheet) -> None:
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, str) and cell.value in VALUE_RENAMES:
                cell.value = VALUE_RENAMES[cell.value]


def _normalize_dashboard_sheet(sheet: Worksheet) -> None:
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value in DASHBOARD_TEXT_RENAMES:
                cell.value = DASHBOARD_TEXT_RENAMES[cell.value]

    sheet["B4"] = '=COUNTA(Oferty!A2:A200)'
    sheet["B5"] = '=COUNTIF(Oferty!G2:G200,"Dostępna")'
    sheet["B6"] = '=COUNTIF(Oferty!H2:H200,"Do analizy")'
    sheet["B7"] = '=COUNTIF(Oferty!H2:H200,"Aplikowano")'
    sheet["B8"] = '=COUNTIF(Oferty!J2:J200,"Wysoki")'
    sheet["B9"] = '=COUNTIF(Oferty!P2:P200,">14")'


def _refresh_dashboard_actions(dashboard_sheet: Worksheet, offers_sheet: Worksheet) -> None:
    action_rows = _dashboard_action_rows(offers_sheet)
    last_action_row = max(dashboard_sheet.max_row, 4 + len(action_rows), 9)

    for row in range(5, last_action_row + 1):
        for column in range(4, 9):
            dashboard_sheet.cell(row=row, column=column).value = None

    for offset, action in enumerate(action_rows, start=5):
        _copy_dashboard_action_style(dashboard_sheet, target_row=offset)
        for column_offset, value in enumerate(action, start=4):
            dashboard_sheet.cell(row=offset, column=column_offset, value=value)


def _dashboard_action_rows(offers_sheet: Worksheet) -> list[tuple[object, ...]]:
    headers = _header_positions(offers_sheet)
    actions: list[tuple[int, tuple[object, ...]]] = []

    for row in range(2, offers_sheet.max_row + 1):
        offer_id = _cell_by_header(offers_sheet, row, headers, "ID")
        if not offer_id:
            continue

        status = _cell_by_header(offers_sheet, row, headers, "Status")
        availability = _cell_by_header(offers_sheet, row, headers, "Dostępność")
        next_step = _cell_by_header(offers_sheet, row, headers, "Następny krok")
        if status in ACTIVE_DASHBOARD_EXCLUDED_STATUSES:
            continue
        if availability in ACTIVE_DASHBOARD_EXCLUDED_AVAILABILITY:
            continue
        if not next_step or next_step == UNKNOWN_VALUE:
            continue

        actions.append(
            (
                _offer_id_sort_key(str(offer_id)),
                (
                    offer_id,
                    _cell_by_header(offers_sheet, row, headers, "Firma"),
                    _cell_by_header(offers_sheet, row, headers, "Stanowisko"),
                    status,
                    next_step,
                ),
            )
        )

    actions.sort(key=lambda item: item[0], reverse=True)
    return [action for _, action in actions]


def _cell_by_header(
    sheet: Worksheet, row: int, headers: dict[str, int], header: str
) -> object:
    column = _column_for_header(headers, header)
    if column is None:
        return None
    return sheet.cell(row=row, column=column).value


def _set_cell_by_header(
    sheet: Worksheet,
    row: int,
    headers: dict[str, int],
    header: str,
    value: object,
) -> None:
    column = _column_for_header(headers, header)
    if column is None:
        raise ValueError(f"Brak kolumny w arkuszu {sheet.title}: {header}")
    sheet.cell(row=row, column=column, value=value)


def _offer_id_sort_key(offer_id: str) -> int:
    match = OFFER_ID_PATTERN.match(offer_id.strip())
    if not match:
        return 0
    return int(match.group(1))


def _copy_dashboard_action_style(sheet: Worksheet, target_row: int) -> None:
    source_row = 5
    if target_row == source_row:
        return

    for column in range(4, 9):
        source_cell = sheet.cell(row=source_row, column=column)
        target_cell = sheet.cell(row=target_row, column=column)
        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.fill = copy(source_cell.fill)
            target_cell.border = copy(source_cell.border)
            target_cell.alignment = copy(source_cell.alignment)
            target_cell.number_format = source_cell.number_format
            target_cell.protection = copy(source_cell.protection)


def _write_values(
    sheet: Worksheet, row: int, headers: dict[str, int], values: dict[str, object]
) -> None:
    missing_headers = [
        header for header in values if _column_for_header(headers, header) is None
    ]
    if missing_headers:
        joined_headers = ", ".join(missing_headers)
        raise ValueError(f"Brak kolumn w arkuszu {sheet.title}: {joined_headers}")

    for header, value in values.items():
        column = _column_for_header(headers, header)
        sheet.cell(row=row, column=column, value=value)


def _column_for_header(headers: dict[str, int], header: str) -> int | None:
    direct_column = headers.get(header) or headers.get(_text_key(header))
    if direct_column is not None:
        return direct_column

    for alias in HEADER_ALIASES.get(header, []):
        alias_column = headers.get(alias) or headers.get(_text_key(alias))
        if alias_column is not None:
            return alias_column

    return None


def _salary_values(salary: SalaryInfo) -> dict[str, object]:
    return {
        "Stawka źródłowa": salary.original_text,
        "Waluta": salary.currency,
        "Stawka min": salary.amount_min,
        "Stawka max": salary.amount_max,
        "Okres stawki": salary.period,
        "Brutto/netto": salary.tax_type,
        "Kurs waluty": salary.exchange_rate_to_pln,
        "Data kursu": salary.exchange_rate_date,
        "PLN min miesięcznie": salary.pln_min_monthly,
        "PLN max miesięcznie": salary.pln_max_monthly,
        "PLN min godzinowo": salary.pln_min_hourly,
        "PLN max godzinowo": salary.pln_max_hourly,
        "Założenia przeliczenia": salary.conversion_assumptions,
    }


def _has_salary_values(salary: SalaryInfo) -> bool:
    return salary.currency != UNKNOWN_VALUE and salary.amount_min is not None


def _text_key(text: str) -> str:
    ascii_text = unicodedata.normalize("NFKD", text.translate(POLISH_TRANSLATION))
    ascii_text = "".join(char for char in ascii_text if not unicodedata.combining(char))
    return ascii_text.strip().lower()


def _first_empty_row(sheet: Worksheet) -> int:
    for row in range(2, sheet.max_row + 2):
        values = [
            sheet.cell(row=row, column=column).value
            for column in range(1, sheet.max_column + 1)
        ]
        if all(value is None for value in values):
            return row
    return sheet.max_row + 1


def _next_offer_id(sheet: Worksheet) -> str:
    max_number = 0
    for row in range(2, sheet.max_row + 1):
        value = sheet.cell(row=row, column=1).value
        if not isinstance(value, str):
            continue

        match = OFFER_ID_PATTERN.match(value.strip())
        if match:
            max_number = max(max_number, int(match.group(1)))

    return f"JOB-{max_number + 1:03d}"


def _copy_row_style(sheet: Worksheet, source_row: int, target_row: int) -> None:
    if source_row == target_row:
        return

    for column in range(1, sheet.max_column + 1):
        source_cell = sheet.cell(row=source_row, column=column)
        target_cell = sheet.cell(row=target_row, column=column)

        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.fill = copy(source_cell.fill)
            target_cell.border = copy(source_cell.border)
            target_cell.alignment = copy(source_cell.alignment)
            target_cell.number_format = source_cell.number_format
            target_cell.protection = copy(source_cell.protection)


def _ensure_workbook_writable(workbook_path: Path) -> None:
    try:
        with open(workbook_path, "a+b"):
            pass
    except PermissionError as exc:
        raise PermissionError(
            "Nie można zapisać pliku Excel. Zamknij skoroszyt w Excelu i spróbuj ponownie."
        ) from exc
