from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from teb_bankasi import teb_dipnot_guncelle as _base


COMPANY_NAME = "denizbank"
DEFAULT_WORKBOOK_PATH = "butunlesik_banka_verisi.xlsx"
DEFAULT_PDF_DIR = "denizbank/data/history"
DEFAULT_CACHE_DIR = "denizbank/data/pdf_cache"
DEFAULT_WORKERS = 1
PEERPDF_SCRIPT_PATH = Path(__file__).resolve().parent / "peerpdf_scripts" / "denizbank_dipnot.py"
CACHE_VERSION = 10
STAGE_PROVISION_TITLE = (
    "Standart nitelikli ve yakin izlemedeki krediler ile yeniden yapilandirilan "
    "yakin izlemedeki kredilere iliskin bilgiler"
)


def configure_base() -> None:
    _base.COMPANY_NAME = COMPANY_NAME
    _base.DEFAULT_WORKBOOK_PATH = DEFAULT_WORKBOOK_PATH
    _base.DEFAULT_PDF_DIR = DEFAULT_PDF_DIR
    _base.DEFAULT_CACHE_DIR = DEFAULT_CACHE_DIR
    _base.PEERPDF_SCRIPT_PATH = PEERPDF_SCRIPT_PATH
    _base.CACHE_VERSION = CACHE_VERSION
    _base.STAGE_PROVISION_TITLE = STAGE_PROVISION_TITLE
    _base.STAGE_PROVISION_TARGET_LABELS = {
        "12 Aylik Beklenen Zarar Karsiligi",
        "Kredi Riskinde Onemli Artis",
    }


def append_denizbank_pdf_notes(
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
    parser = argparse.ArgumentParser(description="Denizbank PDF dipnotlarini butunlesik Excel'e yazar.")
    parser.add_argument("--workbook", default=DEFAULT_WORKBOOK_PATH)
    parser.add_argument("--pdf-dir", default=DEFAULT_PDF_DIR)
    parser.add_argument("--pdf")
    parser.add_argument("--cache-dir", default=DEFAULT_CACHE_DIR)
    parser.add_argument("--workers", type=int, default=DEFAULT_WORKERS)
    parser.add_argument("--sheet", default=COMPANY_NAME)
    args = parser.parse_args()
    source_pdf, rows_written = append_denizbank_pdf_notes(
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
