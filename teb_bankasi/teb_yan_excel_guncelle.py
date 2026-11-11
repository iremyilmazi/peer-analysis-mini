from __future__ import annotations

import argparse
import concurrent.futures
import importlib.util
import json
import re
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from yan_excel_donem_ekle import ensure_sheet, new_pdf_files, next_append_row, payloads_for_new_periods, workbook_periods


PROJECT_ROOT = Path(__file__).resolve().parents[1]
PEERPDF_SCRIPT_PATH = Path(__file__).resolve().parent / "peerpdf_scripts" / "teb_dipnot.py"
DEFAULT_PDF_DIR = Path(__file__).resolve().parent / "data" / "history"
DEFAULT_CACHE_DIR = Path(__file__).resolve().parent / "data" / "side_excel_cache"
DEFAULT_STAGE_OUTPUT = PROJECT_ROOT / "birinci_ikinci_asama_krediler.xlsx"
DEFAULT_DEPOSIT_OUTPUT = PROJECT_ROOT / "mevduat_faiz.xlsx"
DEFAULT_WORKERS = 4
CACHE_VERSION = 2
SHEET_NAME = "teb"
HEADER_FILL = "D9EAF7"
TITLE_FILL = "548235"

STAGE_HEADERS = [
    "Kalem",
    "Standart Nitelikli Krediler",
    "Yakın İzlemedeki Krediler - Sözleşme Koşullarında Değişiklik Yapılanlar",
    "Yakın İzlemedeki Krediler - Yeniden Yapılandırma Kapsamında Yer Almayanlar",
    "Yakın İzlemedeki Krediler - Toplam",
]

DEPOSIT_INTEREST_HEADERS = [
    "Kalem",
    "Vadesiz Mevduat",
    "1 Aya Kadar",
    "3 Aya Kadar",
    "6 Aya Kadar",
    "1 Yıla Kadar",
    "1 Yıldan Uzun",
    "Birikimli Mevduat",
    "Toplam",
]

DEPOSIT_MATURITY_HEADERS = [
    "Kalem",
    "Vadesiz",
    "7 Gün İhbarlı",
    "1 Aya Kadar",
    "1-3 Ay",
    "3-6 Ay",
    "6 Ay-1 Yıl",
    "1 Yıl ve Üstü",
    "Birikimli Mevduat",
    "Toplam",
]


def add_bundled_pdf_dependencies() -> None:
    bundled_site_packages = (
        Path.home()
        / ".cache"
        / "codex-runtimes"
        / "codex-primary-runtime"
        / "dependencies"
        / "python"
        / "Lib"
        / "site-packages"
    )
    if bundled_site_packages.exists() and str(bundled_site_packages) not in sys.path:
        sys.path.append(str(bundled_site_packages))


def load_parser():
    add_bundled_pdf_dependencies()
    if str(PEERPDF_SCRIPT_PATH.parent) not in sys.path:
        sys.path.insert(0, str(PEERPDF_SCRIPT_PATH.parent))
    spec = importlib.util.spec_from_file_location("peerpdf_teb_dipnot_side", PEERPDF_SCRIPT_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"TEB parser yuklenemedi: {PEERPDF_SCRIPT_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def list_pdf_files(pdf_dir: Path) -> list[Path]:
    if not pdf_dir.exists():
        return []
    return sorted(
        [
            path
            for path in pdf_dir.rglob("*.pdf")
            if path.is_file() and not path.name.startswith("~$") and path.stat().st_size > 0
        ],
        key=lambda path: period_sort_key(extract_report_date(path) or "31.12.9999"),
    )


def extract_report_date(pdf_path: Path) -> str:
    match = re.search(r"(\d{2})\s+(\d{2})\s+(\d{4})", pdf_path.stem)
    if not match:
        match = re.search(r"(\d{2})[._-](\d{2})[._-](\d{4})", pdf_path.stem)
    if not match:
        compact = re.search(r"(?<!\d)(\d{2})(\d{2})(20\d{2})(?!\d)", pdf_path.stem)
        if compact:
            day, month, year = compact.groups()
            return f"{day}.{month}.{year}"
    if not match:
        return ""
    day, month, year = match.groups()
    return f"{day}.{month}.{year}"


def period_sort_key(period: str) -> tuple[int, int, int]:
    day, month, year = [int(part) for part in period.split(".")]
    return year, month, day


def cache_path_for_pdf(pdf_path: Path, cache_dir: Path, kind: str) -> Path:
    period = extract_report_date(pdf_path) or "donem_yok"
    safe_stem = re.sub(r"[^A-Za-z0-9_.-]+", "_", pdf_path.stem).strip("_") or "pdf"
    return cache_dir / kind / f"{period.replace('.', '')}_{safe_stem}.json"


def load_cached_payload(pdf_path: Path, cache_dir: Path, kind: str) -> dict | None:
    cache_path = cache_path_for_pdf(pdf_path, cache_dir, kind)
    if not cache_path.exists():
        return None
    try:
        payload = json.loads(cache_path.read_text(encoding="utf-8"))
        pdf_stat = pdf_path.stat()
        if (
            payload.get("cache_version") == CACHE_VERSION
            and payload.get("source_size") == pdf_stat.st_size
            and payload.get("source_mtime") == pdf_stat.st_mtime
        ):
            print(f"{kind} cache kullanildi: {cache_path.name}", flush=True)
            return payload
    except Exception as exc:
        print(f"WARN: {kind} cache okunamadi, yeniden alinacak: {cache_path} -> {type(exc).__name__}: {exc}", flush=True)
    return None


def write_cached_payload(pdf_path: Path, cache_dir: Path, kind: str, payload: dict) -> None:
    cache_path = cache_path_for_pdf(pdf_path, cache_dir, kind)
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    pdf_stat = pdf_path.stat()
    payload.update(
        {
            "cache_version": CACHE_VERSION,
            "source_path": str(pdf_path),
            "source_size": pdf_stat.st_size,
            "source_mtime": pdf_stat.st_mtime,
        }
    )
    cache_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"{kind} cache yazildi: {cache_path.name}", flush=True)


def extract_stage_payload(pdf_path: Path, cache_dir: Path | None) -> dict:
    if cache_dir:
        cached = load_cached_payload(pdf_path, cache_dir, "stage")
        if cached:
            return cached

    parser = load_parser()
    period = extract_report_date(pdf_path)
    try:
        with parser.pdfplumber.open(str(pdf_path)) as pdf:
            page_texts = [page.extract_text(x_tolerance=1, y_tolerance=3) or "" for page in pdf.pages]
            page_tables = [page.extract_tables() for page in pdf.pages]
        rows = parser.extract_raw_rows(page_texts, page_tables, parser.STAGE_PROVISION_TITLE)
        table = [STAGE_HEADERS, *rows]
    except Exception as exc:
        print(f"WARN: {period} TEB stage tablosu bos birakildi: {type(exc).__name__}: {exc}", flush=True)
        table = [[period, *[""] * (len(STAGE_HEADERS) - 1)]]

    payload = {"period": period, "title": parser.STAGE_PROVISION_TITLE, "table": table}
    if cache_dir:
        write_cached_payload(pdf_path, cache_dir, "stage", payload)
    return payload


def extract_deposit_payload(pdf_path: Path, cache_dir: Path | None) -> dict:
    if cache_dir:
        cached = load_cached_payload(pdf_path, cache_dir, "deposit")
        if cached:
            return cached

    parser = load_parser()
    period = extract_report_date(pdf_path)
    try:
        with parser.pdfplumber.open(str(pdf_path)) as pdf:
            page_texts = [page.extract_text(x_tolerance=1, y_tolerance=3) or "" for page in pdf.pages]
            page_tables = [page.extract_tables() for page in pdf.pages]
        previous_pdf_path = getattr(parser, "_ACTIVE_PDF_PATH", None)
        if hasattr(parser, "_ACTIVE_PDF_PATH"):
            parser._ACTIVE_PDF_PATH = pdf_path
        try:
            deposit_rows = parser.extract_deposit_interest_rows(page_texts)
        finally:
            if hasattr(parser, "_ACTIVE_PDF_PATH"):
                parser._ACTIVE_PDF_PATH = previous_pdf_path
    except Exception as exc:
        print(f"WARN: {period} TEB mevduat faiz tablosu bos birakildi: {type(exc).__name__}: {exc}", flush=True)
        deposit_rows = []

    try:
        if "page_texts" not in locals():
            with parser.pdfplumber.open(str(pdf_path)) as pdf:
                page_texts = [page.extract_text(x_tolerance=1, y_tolerance=3) or "" for page in pdf.pages]
                page_tables = [page.extract_tables() for page in pdf.pages]
        maturity_rows = parser.extract_raw_rows(page_texts, page_tables, parser.DEPOSIT_MATURITY_TITLE)
    except Exception as exc:
        print(f"WARN: {period} TEB mevduat vade yapisi tablosu bos birakildi: {type(exc).__name__}: {exc}", flush=True)
        maturity_rows = []

    payload = {
        "period": period,
        "title": parser.DEPOSIT_INTEREST_TITLE,
        "headers": DEPOSIT_INTEREST_HEADERS,
        "rows": deposit_rows,
        "maturity_title": parser.DEPOSIT_MATURITY_TITLE,
        "maturity_headers": DEPOSIT_MATURITY_HEADERS,
        "maturity_rows": maturity_rows,
    }
    if cache_dir:
        write_cached_payload(pdf_path, cache_dir, "deposit", payload)
    return payload


def blank_payload(pdf_path: Path, kind: str) -> dict:
    period = extract_report_date(pdf_path)
    parser = load_parser()
    if kind == "stage":
        return {
            "period": period,
            "title": parser.STAGE_PROVISION_TITLE,
            "table": [[period, *[""] * (len(STAGE_HEADERS) - 1)]],
        }
    if kind == "deposit":
        return {
            "period": period,
            "title": parser.DEPOSIT_INTEREST_TITLE,
            "headers": DEPOSIT_INTEREST_HEADERS,
            "rows": [],
            "maturity_title": parser.DEPOSIT_MATURITY_TITLE,
            "maturity_headers": DEPOSIT_MATURITY_HEADERS,
            "maturity_rows": [],
        }
    raise RuntimeError(f"Bilinmeyen yan Excel tipi: {kind}")


def extract_side_payload_worker(pdf_path_text: str, cache_dir_text: str, kind: str) -> tuple[str, dict]:
    pdf_path = Path(pdf_path_text)
    cache_dir = Path(cache_dir_text) if cache_dir_text else None
    period = extract_report_date(pdf_path)
    print(f"TEB {kind} okunuyor: {period} - {pdf_path.name}", flush=True)
    try:
        if kind == "stage":
            payload = extract_stage_payload(pdf_path, cache_dir)
        elif kind == "deposit":
            payload = extract_deposit_payload(pdf_path, cache_dir)
        else:
            raise RuntimeError(f"Bilinmeyen yan Excel tipi: {kind}")
        return str(pdf_path.resolve()), payload
    except Exception as exc:
        print(f"WARN: {period} TEB {kind} tablosu alinamadi: {pdf_path.name} -> {type(exc).__name__}: {exc}", flush=True)
        return str(pdf_path.resolve()), blank_payload(pdf_path, kind)


def extract_history_payloads(pdf_files: list[Path], cache_dir: Path | None, kind: str, workers: int) -> list[dict]:
    if not pdf_files:
        return []
    worker_count = max(1, min(workers, len(pdf_files)))
    results: dict[str, dict] = {}
    if worker_count == 1:
        for pdf_path in pdf_files:
            key, payload = extract_side_payload_worker(str(pdf_path), str(cache_dir) if cache_dir else "", kind)
            results[key] = payload
    else:
        with concurrent.futures.ProcessPoolExecutor(max_workers=worker_count) as executor:
            futures = {
                executor.submit(extract_side_payload_worker, str(pdf_path), str(cache_dir) if cache_dir else "", kind): pdf_path
                for pdf_path in pdf_files
            }
            for future in concurrent.futures.as_completed(futures):
                key, payload = future.result()
                results[key] = payload
    return [results.get(str(pdf_path.resolve()), blank_payload(pdf_path, kind)) for pdf_path in pdf_files]


def to_excel_number(value):
    if not isinstance(value, str):
        return value
    cleaned = " ".join(value.split())
    if re.fullmatch(r"-?\d+(?:,\d{3})*", cleaned):
        return int(cleaned.replace(",", ""))
    return value


def style_title_range(sheet, row_index: int, start_col: int, max_col: int) -> None:
    for col_index in range(start_col, start_col + max_col):
        cell = sheet.cell(row_index, col_index)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=TITLE_FILL)
        cell.alignment = Alignment(vertical="center", wrap_text=True)


def style_header_range(sheet, row_index: int, start_col: int, max_col: int) -> None:
    for col_index in range(start_col, start_col + max_col):
        cell = sheet.cell(row_index, col_index)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor=HEADER_FILL)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def write_rows_at(sheet, start_row: int, start_col: int, rows: list[list], header_rows: int = 0) -> int:
    current_row = start_row
    max_col = max((len(row) for row in rows), default=1)
    for row_offset, row in enumerate(rows):
        for offset, value in enumerate(row):
            col_index = start_col + offset
            cell = sheet.cell(current_row, col_index, to_excel_number(value))
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if isinstance(cell.value, int):
                cell.number_format = "#,##0"
            if offset > 0:
                cell.alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)
        if row_offset < header_rows:
            style_header_range(sheet, current_row, start_col, max_col)
        current_row += 1
    return current_row


def load_or_create_workbook(path: Path):
    if path.exists():
        return load_workbook(path)
    return Workbook()


def replace_sheet(workbook, sheet_name: str):
    if sheet_name in workbook.sheetnames:
        old_sheet = workbook[sheet_name]
        workbook.remove(old_sheet)
    if len(workbook.sheetnames) == 1 and workbook[workbook.sheetnames[0]].max_row == 1 and workbook[workbook.sheetnames[0]].max_column == 1 and workbook[workbook.sheetnames[0]]["A1"].value is None:
        sheet = workbook[workbook.sheetnames[0]]
        sheet.title = sheet_name
        return sheet
    return workbook.create_sheet(sheet_name)


def write_stage_workbook(payloads: list[dict], output_path: Path) -> int:
    workbook = load_or_create_workbook(output_path)
    existing_sheet = SHEET_NAME in workbook.sheetnames
    sheet = ensure_sheet(workbook, SHEET_NAME)
    pending_payloads = payloads_for_new_periods(payloads, workbook_periods(output_path, SHEET_NAME), SHEET_NAME)
    if not pending_payloads:
        workbook.close()
        return 0
    current_row = next_append_row(sheet)
    written_blocks = 0
    for payload in pending_payloads:
        period = payload["period"]
        table = [list(row) for row in payload["table"]]
        max_col = max((len(row) for row in table), default=len(STAGE_HEADERS))
        sheet.cell(current_row, 1, f"{period} - {payload['title']}")
        style_title_range(sheet, current_row, 1, max_col)
        current_row += 1
        current_row = write_rows_at(sheet, current_row, 1, table, header_rows=1 if table else 0)
        current_row += 2
        written_blocks += 1

    if not existing_sheet:
        for col_index, width in enumerate([34, 24, 34, 34, 24], start=1):
            sheet.column_dimensions[get_column_letter(col_index)].width = width
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    workbook.close()
    return written_blocks


def write_deposit_workbook(payloads: list[dict], output_path: Path) -> int:
    workbook = load_or_create_workbook(output_path)
    existing_sheet = SHEET_NAME in workbook.sheetnames
    sheet = ensure_sheet(workbook, SHEET_NAME)
    pending_payloads = payloads_for_new_periods(payloads, workbook_periods(output_path, SHEET_NAME), SHEET_NAME)
    if not pending_payloads:
        workbook.close()
        return 0
    current_row = next_append_row(sheet)
    written_blocks = 0
    for payload in pending_payloads:
        period = payload["period"]
        left_rows = [payload["headers"], *payload["rows"]]
        right_rows = [payload["maturity_headers"], *payload["maturity_rows"]]
        left_start_col = 1
        right_start_col = len(payload["headers"]) + 2

        left_max_col = max((len(row) for row in left_rows), default=len(payload["headers"]))
        sheet.cell(current_row, left_start_col, f"{period} - {payload['title']}")
        style_title_range(sheet, current_row, left_start_col, left_max_col)
        left_end = write_rows_at(sheet, current_row + 1, left_start_col, left_rows, header_rows=1)

        right_max_col = max((len(row) for row in right_rows), default=len(payload["maturity_headers"]))
        sheet.cell(current_row, right_start_col, f"{period} - {payload['maturity_title']}")
        style_title_range(sheet, current_row, right_start_col, right_max_col)
        right_end = write_rows_at(sheet, current_row + 1, right_start_col, right_rows, header_rows=1)

        current_row = max(left_end, right_end) + 2
        written_blocks += 1

    widths = [
        36,
        18,
        18,
        18,
        18,
        18,
        18,
        20,
        18,
        4,
        34,
        18,
        18,
        18,
        18,
        18,
        18,
        20,
        20,
        18,
    ]
    if not existing_sheet:
        for col_index, width in enumerate(widths, start=1):
            sheet.column_dimensions[get_column_letter(col_index)].width = width
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    workbook.close()
    return written_blocks


def update_side_excels(
    pdf_dir: Path = DEFAULT_PDF_DIR,
    cache_dir: Path | None = DEFAULT_CACHE_DIR,
    stage_output: Path = DEFAULT_STAGE_OUTPUT,
    deposit_output: Path = DEFAULT_DEPOSIT_OUTPUT,
    workers: int = DEFAULT_WORKERS,
) -> tuple[int, int]:
    pdf_files = list_pdf_files(pdf_dir)
    if not pdf_files:
        raise FileNotFoundError(f"TEB PDF history bulunamadi: {pdf_dir}")

    stage_files = new_pdf_files(pdf_files, workbook_periods(stage_output, SHEET_NAME), extract_report_date)
    stage_payloads = extract_history_payloads(stage_files, cache_dir, "stage", workers)
    stage_count = write_stage_workbook(stage_payloads, stage_output)
    print(f"TEB yan Excel OK: {stage_output} -> {stage_count} yeni donem")

    deposit_files = new_pdf_files(pdf_files, workbook_periods(deposit_output, SHEET_NAME), extract_report_date)
    deposit_payloads = extract_history_payloads(deposit_files, cache_dir, "deposit", workers)
    deposit_count = write_deposit_workbook(deposit_payloads, deposit_output)
    print(f"TEB yan Excel OK: {deposit_output} -> {deposit_count} yeni donem")
    return stage_count, deposit_count


def main() -> None:
    parser = argparse.ArgumentParser(description="TEB yan Excel dosyalarina history PDF'lerinden TEB sheet'i ekler.")
    parser.add_argument("--pdf-dir", default=str(DEFAULT_PDF_DIR), help="TEB PDF history klasoru.")
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
