from __future__ import annotations

import argparse
from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from job_offer_analyzer.excel_writer import append_offer_to_workbook
from job_offer_analyzer.models import OfferRecord


DEFAULT_WORKBOOK = PROJECT_ROOT / "data" / "baza_ofert_pracy_QA.xlsx"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Dodaje testową ofertę pracy do pliku Excel."
    )
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument("--company", default="Test Company")
    parser.add_argument("--title", default="QA Engineer")
    parser.add_argument("--link", default="https://example.com/job/qa-engineer")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    offer = OfferRecord(
        company=args.company,
        title=args.title,
        link=args.link,
        status="Nowa",
        availability="Dostępna",
        next_step="Pobrać treść oferty i wykonać analizę pod CV",
        source=args.link,
    )

    offer_id = append_offer_to_workbook(args.workbook, offer)
    print(f"Dodano {offer_id} do {args.workbook}")


if __name__ == "__main__":
    main()
