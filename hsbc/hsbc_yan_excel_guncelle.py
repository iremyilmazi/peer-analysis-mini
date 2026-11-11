from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from teb_bankasi import teb_yan_excel_guncelle as _base


_BASE_EXTRACT_REPORT_DATE = _base.extract_report_date
PEERPDF_SCRIPT_PATH = Path(__file__).resolve().parent / "peerpdf_scripts" / "hsbc_dipnot.py"
DEFAULT_PDF_DIR = Path(__file__).resolve().parent / "data" / "history"
DEFAULT_CACHE_DIR = Path(__file__).resolve().parent / "data" / "side_excel_cache"
DEFAULT_STAGE_OUTPUT = PROJECT_ROOT / "birinci_ikinci_asama_krediler.xlsx"
DEFAULT_DEPOSIT_OUTPUT = PROJECT_ROOT / "mevduat_faiz.xlsx"
DEFAULT_WORKERS = 1
CACHE_VERSION = 2
SHEET_NAME = "hsbc"


def extract_hsbc_report_date(pdf_path: Path) -> str:
    period = _BASE_EXTRACT_REPORT_DATE(pdf_path)
    if period:
        return period
    match = re.search(r"(?<!\d)(03|06|09|12)(20|21|22|23|24|25|26)(?:TR)?(?!\d)", pdf_path.stem, re.IGNORECASE)
    if not match:
        return ""
    month, year_suffix = match.groups()
    day_by_month = {"03": "31", "06": "30", "09": "30", "12": "31"}
    return f"{day_by_month[month]}.{month}.20{year_suffix}"


def configure_base() -> None:
    _base.PEERPDF_SCRIPT_PATH = PEERPDF_SCRIPT_PATH
    _base.DEFAULT_PDF_DIR = DEFAULT_PDF_DIR
    _base.DEFAULT_CACHE_DIR = DEFAULT_CACHE_DIR
    _base.DEFAULT_STAGE_OUTPUT = DEFAULT_STAGE_OUTPUT
    _base.DEFAULT_DEPOSIT_OUTPUT = DEFAULT_DEPOSIT_OUTPUT
    _base.CACHE_VERSION = CACHE_VERSION
    _base.SHEET_NAME = SHEET_NAME
    _base.extract_report_date = extract_hsbc_report_date


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
    parser = argparse.ArgumentParser(description="HSBC yan Excel dosyalarini gunceller.")
    parser.add_argument("--pdf-dir", default=str(DEFAULT_PDF_DIR))
    parser.add_argument("--cache-dir", default=str(DEFAULT_CACHE_DIR))
    parser.add_argument("--stage-output", default=str(DEFAULT_STAGE_OUTPUT))
    parser.add_argument("--deposit-output", default=str(DEFAULT_DEPOSIT_OUTPUT))
    parser.add_argument("--workers", type=int, default=DEFAULT_WORKERS)
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
