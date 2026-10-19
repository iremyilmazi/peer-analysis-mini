from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from teb_bankasi import teb_dipnot_guncelle as _base


COMPANY_NAME = "vakifbank"
DEFAULT_WORKBOOK_PATH = "butunlesik_banka_verisi.xlsx"
DEFAULT_PDF_DIR = "vakifbank/data/history"
DEFAULT_CACHE_DIR = "vakifbank/data/pdf_cache"
DEFAULT_WORKERS = 1
PEERPDF_SCRIPT_PATH = Path(__file__).resolve().parent / "peerpdf_scripts" / "vakifbank_dipnot.py"
CACHE_VERSION = 6
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

STAGE_PROVISION_TITLE = (
    "Standart nitelikli ve yakın izlemedeki krediler ile yeniden yapılandırılan "
    "yakın izlemedeki kredilere ilişkin bilgiler"
)


def configure_base() -> None:
    _base.COMPANY_NAME = COMPANY_NAME
    _base.DEFAULT_WORKBOOK_PATH = DEFAULT_WORKBOOK_PATH
    _base.DEFAULT_PDF_DIR = DEFAULT_PDF_DIR
    _base.DEFAULT_CACHE_DIR = DEFAULT_CACHE_DIR
    _base.PEERPDF_SCRIPT_PATH = PEERPDF_SCRIPT_PATH
    _base.CACHE_VERSION = CACHE_VERSION
    _base.STAGE_PROVISION_TITLE = STAGE_PROVISION_TITLE
    _base.extract_report_date = extract_vakifbank_report_date
    _base.list_pdf_files = list_vakifbank_pdf_files
    _base.STAGE_PROVISION_TARGET_LABELS = {
        "12 Aylik Beklenen Zarar Karsiligi",
        "Kredi Riskinde Onemli Artis",
    }


def append_vakifbank_pdf_notes(
    workbook_path: Path,
    pdf_dir: Path,
    pdf_path: Path | None = None,
    sheet_name: str = COMPANY_NAME,
    cache_dir: Path | None = Path(DEFAULT_CACHE_DIR),
    workers: int = DEFAULT_WORKERS,
) -> tuple[Path, int]:
    configure_base()
    return _base.append_teb_pdf_notes(
        workbook_path=workbook_path,
        pdf_dir=pdf_dir,
        pdf_path=pdf_path,
        sheet_name=sheet_name,
        cache_dir=cache_dir,
        workers=workers,
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Vakifbank PDF dipnotlarini butunlesik Excel'deki Vakifbank sheetinin altina yazar.")
    parser.add_argument("--workbook", default=DEFAULT_WORKBOOK_PATH, help="Guncellenecek butunlesik Excel dosyasi.")
    parser.add_argument("--pdf-dir", default=DEFAULT_PDF_DIR, help="Vakifbank PDF dosyalarinin bulundugu klasor.")
    parser.add_argument("--pdf", help="Belirli bir PDF dosyasi kullan.")
    parser.add_argument("--cache-dir", default=DEFAULT_CACHE_DIR, help="PDF dipnot cache klasoru.")
    parser.add_argument("--workers", type=int, default=DEFAULT_WORKERS, help="History PDF ayrıştırma icin paralel worker sayisi.")
    parser.add_argument("--sheet", default=COMPANY_NAME, help="Guncellenecek sheet adi.")
    args = parser.parse_args()

    source_pdf, rows_written = append_vakifbank_pdf_notes(
        workbook_path=Path(args.workbook),
        pdf_dir=Path(args.pdf_dir),
        pdf_path=Path(args.pdf) if args.pdf else None,
        sheet_name=args.sheet,
        cache_dir=Path(args.cache_dir) if args.cache_dir else None,
        workers=args.workers,
    )
    print(f"PDF dipnotlari guncellendi: {source_pdf} -> {args.workbook} ({rows_written} satir)")


if __name__ == "__main__":
    main()
