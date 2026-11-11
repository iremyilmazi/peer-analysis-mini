from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from teb_bankasi import teb_yan_excel_guncelle as _base


PEERPDF_SCRIPT_PATH = Path(__file__).resolve().parent / "peerpdf_scripts" / "vakifbank_dipnot.py"
DEFAULT_PDF_DIR = Path(__file__).resolve().parent / "data" / "history"
DEFAULT_CACHE_DIR = Path(__file__).resolve().parent / "data" / "side_excel_cache"
DEFAULT_STAGE_OUTPUT = PROJECT_ROOT / "birinci_ikinci_asama_krediler.xlsx"
DEFAULT_DEPOSIT_OUTPUT = PROJECT_ROOT / "mevduat_faiz.xlsx"
DEFAULT_WORKERS = 1
CACHE_VERSION = 1
SHEET_NAME = "vakifbank"
_ORIGINAL_EXTRACT_REPORT_DATE = _base.extract_report_date


def extract_vakifbank_report_date(pdf_path: Path) -> str:
    quarter_dates = {
        "Q1": "31.03",
        "Q2": "30.06",
        "Q3": "30.09",
        "Q4": "31.12",
    }
    parent_name = pdf_path.parent.name.upper()
    if len(parent_name) == 6 and parent_name[:4].isdigit() and parent_name[4:] in quarter_dates:
        return f"{quarter_dates[parent_name[4:]]}.{parent_name[:4]}"
    return _ORIGINAL_EXTRACT_REPORT_DATE(pdf_path)


def is_vakifbank_primary_pdf(pdf_path: Path) -> bool:
    name_key = pdf_path.name.casefold()
    if "ingilizce" in name_key or "unconsolidated" in name_key:
        return False
    return "solo vakbn" in name_key or "vakbn solo" in name_key or name_key.startswith("vakbn_solo")


def list_vakifbank_pdf_files(pdf_dir: Path) -> list[Path]:
    if not pdf_dir.exists():
        return []
    candidates = [
        path
        for path in pdf_dir.rglob("*.pdf")
        if path.is_file()
        and not path.name.startswith("~$")
        and path.stat().st_size > 0
        and is_vakifbank_primary_pdf(path)
    ]
    return sorted(candidates, key=lambda path: _base.period_sort_key(extract_vakifbank_report_date(path)))


def configure_base() -> None:
    _base.PEERPDF_SCRIPT_PATH = PEERPDF_SCRIPT_PATH
    _base.DEFAULT_PDF_DIR = DEFAULT_PDF_DIR
    _base.DEFAULT_CACHE_DIR = DEFAULT_CACHE_DIR
    _base.DEFAULT_STAGE_OUTPUT = DEFAULT_STAGE_OUTPUT
    _base.DEFAULT_DEPOSIT_OUTPUT = DEFAULT_DEPOSIT_OUTPUT
    _base.CACHE_VERSION = CACHE_VERSION
    _base.SHEET_NAME = SHEET_NAME
    _base.extract_report_date = extract_vakifbank_report_date
    _base.list_pdf_files = list_vakifbank_pdf_files


def update_side_excels(
    pdf_dir: Path = DEFAULT_PDF_DIR,
    cache_dir: Path | None = DEFAULT_CACHE_DIR,
    stage_output: Path = DEFAULT_STAGE_OUTPUT,
    deposit_output: Path = DEFAULT_DEPOSIT_OUTPUT,
    workers: int = DEFAULT_WORKERS,
) -> tuple[int, int]:
    configure_base()
    return _base.update_side_excels(
        pdf_dir=pdf_dir,
        cache_dir=cache_dir,
        stage_output=stage_output,
        deposit_output=deposit_output,
        workers=workers,
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Vakifbank yan Excel dosyalarina history PDF'lerinden Vakifbank sheet'i ekler.")
    parser.add_argument("--pdf-dir", default=str(DEFAULT_PDF_DIR), help="Vakifbank PDF history klasoru.")
    parser.add_argument("--cache-dir", default=str(DEFAULT_CACHE_DIR), help="Yan Excel tablo cache klasoru.")
    parser.add_argument("--stage-output", default=str(DEFAULT_STAGE_OUTPUT), help="Birinci/ikinci asama krediler Excel yolu.")
    parser.add_argument("--deposit-output", default=str(DEFAULT_DEPOSIT_OUTPUT), help="Mevduat faiz Excel yolu.")
    parser.add_argument("--workers", type=int, default=DEFAULT_WORKERS, help="Paralel PDF worker sayisi.")
    args = parser.parse_args()

    update_side_excels(
        pdf_dir=Path(args.pdf_dir),
        cache_dir=Path(args.cache_dir) if args.cache_dir else None,
        stage_output=Path(args.stage_output),
        deposit_output=Path(args.deposit_output),
        workers=args.workers,
    )


if __name__ == "__main__":
    main()
