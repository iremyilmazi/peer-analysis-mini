from __future__ import annotations

import argparse
import importlib.util
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from teb_bankasi import teb_dipnot_guncelle as _base


COMPANY_NAME = "akbank"
DEFAULT_WORKBOOK_PATH = "butunlesik_banka_verisi.xlsx"
DEFAULT_PDF_DIR = "akbank/data/history"
DEFAULT_CACHE_DIR = "akbank/data/pdf_cache"
DEFAULT_WORKERS = 1
PEERPDF_SCRIPT_PATH = Path(__file__).resolve().parent / "peerpdf_scripts" / "akbank_dipnot.py"
CACHE_VERSION = 9
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
    _base.TARGET_SECTIONS = [
        (
            order,
            suborder,
            key,
            STAGE_PROVISION_TITLE if key == "stage12" else title,
            ["Kalem", "Standart Nitelikli Krediler", "Yakın İzlemedeki Krediler"] if key == "stage12" else headers,
        )
        for order, suborder, key, title, headers in _base.TARGET_SECTIONS
    ]
    _base.STAGE_PROVISION_TARGET_LABELS = {
        "12 Aylik Beklenen Zarar Karsiligi",
        "Kredi Riskinde Onemli Artis",
    }
    _base.extract_stage_provision_rows = extract_akbank_stage_rows


def extract_akbank_stage_rows(pdf_path: Path) -> tuple[str, list[str], list[list[str]]]:
    """Extract current standard/close-monitoring provisions from Akbank PDF."""
    spec = importlib.util.spec_from_file_location(
        "peerpdf_akbank_stage_totals",
        PEERPDF_SCRIPT_PATH,
    )
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Akbank PDF ayrıştırıcısı yüklenemedi: {PEERPDF_SCRIPT_PATH}")
    parser = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(parser)
    with parser.pdfplumber.open(str(pdf_path)) as pdf:
        page_texts = [
            page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            for page in pdf.pages
        ]
    rows: list[list[str]] = []
    targets = (
        ("12 aylik beklenen zarar karsiligi", "12 Aylık Beklenen Zarar Karşılığı"),
        ("kredi riskinde onemli artis", "Kredi Riskinde Önemli Artış"),
    )
    for text in page_texts:
        for raw_line in text.splitlines():
            parsed = parser._parse_amount_line(raw_line)
            if not parsed:
                continue
            key = parser._base.normalize_key(parsed[0])
            canonical = next((label for target, label in targets if target in key), None)
            if canonical is None or len(parsed) < 3:
                continue
            rows.append([canonical, parsed[1], parsed[2]])
        if len(rows) >= 2:
            break
    if len(rows) < 2:
        raise RuntimeError("Akbank standart/yakin izleme karsilik satirlari bulunamadi")
    return (
        STAGE_PROVISION_TITLE,
        ["Kalem", "Standart Nitelikli Krediler", "Yakın İzlemedeki Krediler"],
        rows[:2],
    )


def append_akbank_pdf_notes(
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
    parser = argparse.ArgumentParser(description="Akbank PDF dipnotlarini butunlesik Excel'e yazar.")
    parser.add_argument("--workbook", default=DEFAULT_WORKBOOK_PATH)
    parser.add_argument("--pdf-dir", default=DEFAULT_PDF_DIR)
    parser.add_argument("--pdf")
    parser.add_argument("--cache-dir", default=DEFAULT_CACHE_DIR)
    parser.add_argument("--workers", type=int, default=DEFAULT_WORKERS)
    parser.add_argument("--sheet", default=COMPANY_NAME)
    args = parser.parse_args()
    source_pdf, rows_written = append_akbank_pdf_notes(
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
