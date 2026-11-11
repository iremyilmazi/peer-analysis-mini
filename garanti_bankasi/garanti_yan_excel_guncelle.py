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
PEERPDF_DIR = Path(__file__).resolve().parent / "peerpdf_scripts"
DEFAULT_PDF_DIR = Path(__file__).resolve().parent / "data" / "history"
DEFAULT_CACHE_DIR = Path(__file__).resolve().parent / "data" / "side_excel_cache"
DEFAULT_STAGE_OUTPUT = PROJECT_ROOT / "birinci_ikinci_asama_krediler.xlsx"
DEFAULT_DEPOSIT_OUTPUT = PROJECT_ROOT / "mevduat_faiz.xlsx"
DEFAULT_WORKERS = 4
CACHE_VERSION = 1
HEADER_FILL = "D9EAF7"
TITLE_FILL = "548235"
STAGE_BLANK_COLUMNS = 5


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
        return ""
    day, month, year = match.groups()
    return f"{day}.{month}.{year}"


def period_sort_key(period: str) -> tuple[int, int, int]:
    day, month, year = [int(part) for part in period.split(".")]
    return year, month, day


def load_helper(module_name: str):
    add_bundled_pdf_dependencies()
    helper_path = PEERPDF_DIR / f"{module_name}.py"
    if not helper_path.exists():
        raise FileNotFoundError(f"PeerPDF yardimci dosyasi bulunamadi: {helper_path}")
    module_dir = str(PEERPDF_DIR)
    if module_dir not in sys.path:
        sys.path.insert(0, module_dir)
    spec = importlib.util.spec_from_file_location(f"peerpdf_{module_name}", helper_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"PeerPDF yardimci dosyasi yuklenemedi: {helper_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


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

    helper = load_helper("birinci_ikinci_asama_krediler")
    period = extract_report_date(pdf_path)
    try:
        table = helper.extract_current_period_table(pdf_path)
    except Exception as exc:
        print(f"WARN: {period} stage tablosu bos birakildi: {type(exc).__name__}: {exc}", flush=True)
        table = [[period, *[""] * (STAGE_BLANK_COLUMNS - 1)]]
    payload = {
        "period": period,
        "title": helper.TABLE_TITLE,
        "table": table,
    }
    if cache_dir:
        write_cached_payload(pdf_path, cache_dir, "stage", payload)
    return payload


def extract_deposit_payload(pdf_path: Path, cache_dir: Path | None) -> dict:
    if cache_dir:
        cached = load_cached_payload(pdf_path, cache_dir, "deposit")
        if cached:
            return cached

    helper = load_helper("mevduat_faiz")
    period = extract_report_date(pdf_path)
    try:
        tables = helper.extract_tables(pdf_path)
        table_payloads = [{"period_label": label, "rows": rows} for label, rows in tables]
    except Exception as exc:
        print(f"WARN: {period} mevduat faiz tablosu bos birakildi: {type(exc).__name__}: {exc}", flush=True)
        table_payloads = []

    try:
        maturity_rows = helper.extract_maturity_rows(pdf_path)
    except Exception as exc:
        print(f"WARN: {period} mevduat vade yapisi tablosu bos birakildi: {type(exc).__name__}: {exc}", flush=True)
        maturity_rows = []

    payload = {
        "period": period,
        "title": helper.TABLE_TITLE,
        "headers": helper.HEADERS,
        "tables": table_payloads,
        "maturity_title": helper.MATURITY_TABLE_TITLE,
        "maturity_headers": helper.MATURITY_HEADERS,
        "maturity_rows": maturity_rows,
    }
    if cache_dir:
        write_cached_payload(pdf_path, cache_dir, "deposit", payload)
    return payload


def extract_side_payload_worker(pdf_path_text: str, cache_dir_text: str, kind: str) -> tuple[str, dict | None]:
    pdf_path = Path(pdf_path_text)
    cache_dir = Path(cache_dir_text) if cache_dir_text else None
    period = extract_report_date(pdf_path)
    print(f"{kind} okunuyor: {period} - {pdf_path.name}", flush=True)
    try:
        if kind == "stage":
            payload = extract_stage_payload(pdf_path, cache_dir)
        elif kind == "deposit":
            payload = extract_deposit_payload(pdf_path, cache_dir)
        else:
            raise RuntimeError(f"Bilinmeyen yan Excel tipi: {kind}")
        return str(pdf_path.resolve()), payload
    except Exception as exc:
        print(f"WARN: {period} {kind} tablosu alinamadi: {pdf_path.name} -> {type(exc).__name__}: {exc}", flush=True)
        return str(pdf_path.resolve()), blank_payload(pdf_path, kind)


def blank_payload(pdf_path: Path, kind: str) -> dict:
    period = extract_report_date(pdf_path)
    if kind == "stage":
        helper = load_helper("birinci_ikinci_asama_krediler")
        return {
            "period": period,
            "title": helper.TABLE_TITLE,
            "table": [[period, *[""] * (STAGE_BLANK_COLUMNS - 1)]],
        }
    if kind == "deposit":
        helper = load_helper("mevduat_faiz")
        return {
            "period": period,
            "title": helper.TABLE_TITLE,
            "headers": helper.HEADERS,
            "tables": [],
            "maturity_title": helper.MATURITY_TABLE_TITLE,
            "maturity_headers": helper.MATURITY_HEADERS,
            "maturity_rows": [],
        }
    raise RuntimeError(f"Bilinmeyen yan Excel tipi: {kind}")


def extract_history_payloads(
    pdf_files: list[Path],
    cache_dir: Path | None,
    kind: str,
    workers: int,
) -> list[dict]:
    if not pdf_files:
        return []
    worker_count = max(1, min(workers, len(pdf_files)))
    results: dict[str, dict | None] = {}
    if worker_count == 1:
        for pdf_path in pdf_files:
            key, payload = extract_side_payload_worker(str(pdf_path), str(cache_dir) if cache_dir else "", kind)
            results[key] = payload
    else:
        with concurrent.futures.ProcessPoolExecutor(max_workers=worker_count) as executor:
            futures = {
                executor.submit(
                    extract_side_payload_worker,
                    str(pdf_path),
                    str(cache_dir) if cache_dir else "",
                    kind,
                ): pdf_path
                for pdf_path in pdf_files
            }
            for future in concurrent.futures.as_completed(futures):
                key, payload = future.result()
                results[key] = payload

    ordered_payloads: list[dict] = []
    for pdf_path in pdf_files:
        payload = results.get(str(pdf_path.resolve()))
        if payload is None:
            payload = blank_payload(pdf_path, kind)
        ordered_payloads.append(payload)
    return ordered_payloads


def to_excel_number(value):
    if not isinstance(value, str):
        return value
    cleaned = " ".join(value.split())
    if re.fullmatch(r"-?\d+(?:,\d{3})*", cleaned):
        return int(cleaned.replace(",", ""))
    return value


def style_title_row(sheet, row_index: int, max_col: int) -> None:
    for col_index in range(1, max_col + 1):
        cell = sheet.cell(row_index, col_index)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=TITLE_FILL)
        cell.alignment = Alignment(vertical="center", wrap_text=True)


def style_header_row(sheet, row_index: int, max_col: int) -> None:
    for col_index in range(1, max_col + 1):
        cell = sheet.cell(row_index, col_index)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor=HEADER_FILL)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def write_rows(sheet, start_row: int, rows: list[list], header_rows: int = 0) -> int:
    current_row = start_row
    max_col = max((len(row) for row in rows), default=1)
    for row_offset, row in enumerate(rows):
        for col_index, value in enumerate(row, start=1):
            cell = sheet.cell(current_row, col_index, to_excel_number(value))
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if isinstance(cell.value, int):
                cell.number_format = "#,##0"
            if col_index > 1:
                cell.alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)
        if row_offset < header_rows:
            style_header_row(sheet, current_row, max_col)
        current_row += 1
    return current_row


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


def write_rows_at(
    sheet,
    start_row: int,
    start_col: int,
    rows: list[list],
    header_rows: int = 0,
) -> int:
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


def write_stage_workbook(payloads: list[dict], output_path: Path) -> int:
    workbook = load_workbook(output_path) if output_path.exists() else Workbook()
    existing_sheet = "garanti" in workbook.sheetnames
    sheet = ensure_sheet(workbook, "garanti")
    pending_payloads = payloads_for_new_periods(payloads, workbook_periods(output_path, "garanti"), "garanti")
    if not pending_payloads:
        workbook.close()
        return 0
    current_row = next_append_row(sheet)
    written_blocks = 0
    for payload in pending_payloads:
        period = payload["period"]
        table = [list(row) for row in payload["table"]]
        if table:
            table[0][0] = period
        max_col = max((len(row) for row in table), default=5)
        sheet.cell(current_row, 1, f"{period} - {payload['title']}")
        style_title_row(sheet, current_row, max_col)
        current_row += 1
        current_row = write_rows(sheet, current_row, table, header_rows=min(3, len(table)))
        current_row += 2
        written_blocks += 1

    if not existing_sheet:
        for col_index, width in enumerate([30, 22, 26, 26, 20], start=1):
            sheet.column_dimensions[get_column_letter(col_index)].width = width
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    workbook.close()
    return written_blocks


def write_deposit_workbook(payloads: list[dict], output_path: Path) -> int:
    workbook = load_workbook(output_path) if output_path.exists() else Workbook()
    existing_sheet = "garanti" in workbook.sheetnames
    sheet = ensure_sheet(workbook, "garanti")
    pending_payloads = payloads_for_new_periods(payloads, workbook_periods(output_path, "garanti"), "garanti")
    if not pending_payloads:
        workbook.close()
        return 0
    current_row = next_append_row(sheet)
    written_blocks = 0
    for payload in pending_payloads:
        period = payload["period"]
        left_start_col = 1
        right_start_col = len(payload["headers"]) + 2
        block_end_row = current_row

        left_rows: list[list] = []
        for table in payload["tables"]:
            headers = [
                period if header == "{report_date}" else header
                for header in payload["headers"]
            ]
            left_rows.extend([[headers[0], *headers[1:]], *table["rows"]])
        if not left_rows:
            headers = [
                period if header == "{report_date}" else header
                for header in payload["headers"]
            ]
            left_rows = [[headers[0], *headers[1:]]]

        left_max_col = max((len(row) for row in left_rows), default=len(payload["headers"]))
        sheet.cell(current_row, left_start_col, f"{period} - {payload['title']}")
        style_title_range(sheet, current_row, left_start_col, left_max_col)
        left_end_row = write_rows_at(sheet, current_row + 1, left_start_col, left_rows, header_rows=1)
        block_end_row = max(block_end_row, left_end_row)

        maturity_rows = [payload["maturity_headers"], *payload["maturity_rows"]]
        right_max_col = max((len(row) for row in maturity_rows), default=len(payload["maturity_headers"]))
        sheet.cell(current_row, right_start_col, f"{period} - {payload['maturity_title']}")
        style_title_range(sheet, current_row, right_start_col, right_max_col)
        right_end_row = write_rows_at(sheet, current_row + 1, right_start_col, maturity_rows, header_rows=1)
        block_end_row = max(block_end_row, right_end_row)

        current_row = block_end_row + 2
        written_blocks += 1

    widths = [32, 18, 18, 18, 18, 18, 20, 20, 18, 4, 32, 18, 18, 18, 18, 18, 18, 20, 20, 18]
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
        raise FileNotFoundError(f"Garanti PDF history bulunamadi: {pdf_dir}")

    stage_files = new_pdf_files(pdf_files, workbook_periods(stage_output, "garanti"), extract_report_date)
    stage_payloads = extract_history_payloads(stage_files, cache_dir, "stage", workers)
    stage_count = write_stage_workbook(stage_payloads, stage_output)
    print(f"Yan Excel OK: {stage_output} -> {stage_count} yeni donem")

    deposit_files = new_pdf_files(pdf_files, workbook_periods(deposit_output, "garanti"), extract_report_date)
    deposit_payloads = extract_history_payloads(deposit_files, cache_dir, "deposit", workers)
    deposit_count = write_deposit_workbook(deposit_payloads, deposit_output)
    print(f"Yan Excel OK: {deposit_output} -> {deposit_count} yeni donem")
    return stage_count, deposit_count


def main() -> None:
    parser = argparse.ArgumentParser(description="Garanti yan Excel dosyalarini history PDF'lerinden blok blok uretir.")
    parser.add_argument("--pdf-dir", default=str(DEFAULT_PDF_DIR), help="Garanti PDF history klasoru.")
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
