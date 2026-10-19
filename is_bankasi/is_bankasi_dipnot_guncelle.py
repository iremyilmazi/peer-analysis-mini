from __future__ import annotations

import argparse
import importlib.util
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from teb_bankasi import teb_dipnot_guncelle as _base


COMPANY_NAME = "is-bankasi"
DEFAULT_WORKBOOK_PATH = "butunlesik_banka_verisi.xlsx"
DEFAULT_PDF_DIR = "is_bankasi/data/history"
DEFAULT_CACHE_DIR = "is_bankasi/data/pdf_cache"
DEFAULT_WORKERS = 1
PEERPDF_SCRIPT_PATH = Path(__file__).resolve().parent / "peerpdf_scripts" / "is_bankasi_dipnot.py"
CACHE_VERSION = 43
STAGE_PROVISION_TITLE = (
    "Standart nitelikli ve yakin izlemedeki krediler ile yeniden yapilandirilan "
    "yakin izlemedeki kredilere iliskin bilgiler"
)
STAGE_AGEING_TITLE = (
    "Birinci ve ikinci grup krediler, diger alacaklar ile yeniden yapilandirilan "
    "ya da yeni bir itfa planina baglanan krediler ve bunlar icin ayrilan "
    "karsiliklara iliskin bilgiler"
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
    _base.extract_stage_provision_rows = extract_is_bankasi_stage_rows


def extract_is_bankasi_stage_rows(pdf_path: Path) -> tuple[str, list[str], list[list[str]]]:
    """Use İş Bankası's 31-60 / 61-90 day table and retain only Total."""
    spec = importlib.util.spec_from_file_location(
        "peerpdf_is_bankasi_stage_totals",
        PEERPDF_SCRIPT_PATH,
    )
    if spec is None or spec.loader is None:
        raise RuntimeError(f"İş Bankası PDF ayrıştırıcısı yüklenemedi: {PEERPDF_SCRIPT_PATH}")
    parser = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(parser)
    with parser.pdfplumber.open(str(pdf_path)) as pdf:
        page_texts = [
            page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            for page in pdf.pages
        ]
    return (
        STAGE_PROVISION_TITLE,
        ["Kalem", "", "Toplam"],
        parser.extract_stage_provision_rows(page_texts),
    )


def extract_is_bankasi_stage_ageing_rows(
    pdf_path: Path,
) -> tuple[str, list[str], list[list[str]]]:
    spec = importlib.util.spec_from_file_location(
        "peerpdf_is_bankasi_stage_ageing",
        PEERPDF_SCRIPT_PATH,
    )
    if spec is None or spec.loader is None:
        raise RuntimeError(f"İş Bankası PDF ayrıştırıcısı yüklenemedi: {PEERPDF_SCRIPT_PATH}")
    parser = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(parser)
    with parser.pdfplumber.open(str(pdf_path)) as pdf:
        page_texts = [
            page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            for page in pdf.pages
        ]
    return (
        STAGE_AGEING_TITLE,
        ["Kalem", "", "Toplam"],
        parser.extract_stage_ageing_rows(page_texts),
    )


def append_is_bankasi_pdf_notes(
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
    parser = argparse.ArgumentParser(description="Is Bankasi PDF dipnotlarini butunlesik Excel'e yazar.")
    parser.add_argument("--workbook", default=DEFAULT_WORKBOOK_PATH)
    parser.add_argument("--pdf-dir", default=DEFAULT_PDF_DIR)
    parser.add_argument("--pdf")
    parser.add_argument("--cache-dir", default=DEFAULT_CACHE_DIR)
    parser.add_argument("--workers", type=int, default=DEFAULT_WORKERS)
    parser.add_argument("--sheet", default=COMPANY_NAME)
    args = parser.parse_args()
    source_pdf, rows_written = append_is_bankasi_pdf_notes(
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
