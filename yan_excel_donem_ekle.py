from __future__ import annotations

"""Yan Excel sayfalarina tarih sirasi bozulmadan yeni donem bloklari ekler."""

import re
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook


PERIOD_TITLE_RE = re.compile(r"^(\d{2}\.\d{2}\.\d{4})\s*-")


def period_sort_key(period: str) -> datetime:
    return datetime.strptime(period, "%d.%m.%Y")


def sheet_periods(sheet) -> list[str]:
    periods: list[str] = []
    for (value,) in sheet.iter_rows(min_col=1, max_col=1, values_only=True):
        if not isinstance(value, str):
            continue
        match = PERIOD_TITLE_RE.match(value.strip())
        if match:
            periods.append(match.group(1))
    if len(periods) != len(set(periods)):
        raise RuntimeError(f"{sheet.title} sheet'inde yinelenen donem basligi var.")
    return periods


def workbook_periods(path: Path, sheet_name: str) -> list[str]:
    if not path.exists():
        return []
    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        if sheet_name not in workbook.sheetnames:
            return []
        return sheet_periods(workbook[sheet_name])
    finally:
        workbook.close()


def ensure_sheet(workbook, sheet_name: str):
    if sheet_name in workbook.sheetnames:
        return workbook[sheet_name]
    if (
        len(workbook.sheetnames) == 1
        and workbook.active.max_row == 1
        and workbook.active.max_column == 1
        and workbook.active["A1"].value is None
    ):
        workbook.active.title = sheet_name
        return workbook.active
    return workbook.create_sheet(sheet_name)


def payloads_for_new_periods(payloads: list[dict], existing_periods: list[str], sheet_name: str) -> list[dict]:
    seen: set[str] = set()
    ordered: list[dict] = []
    for payload in payloads:
        period = str(payload.get("period") or "").strip()
        try:
            period_sort_key(period)
        except ValueError as exc:
            raise RuntimeError(f"{sheet_name} icin gecersiz kaynak donem: {period!r}") from exc
        if period in seen:
            raise RuntimeError(f"{sheet_name} icin kaynakta yinelenen donem var: {period}")
        seen.add(period)
        ordered.append(payload)

    if existing_periods:
        latest_existing = max(period_sort_key(period) for period in existing_periods)
        late_missing = [
            str(payload["period"])
            for payload in ordered
            if str(payload["period"]) not in existing_periods
            and period_sort_key(str(payload["period"])) <= latest_existing
        ]
        if late_missing:
            raise RuntimeError(
                f"{sheet_name} sheet'inde tarihsel bosluk var; eski donem sona eklenemez: {late_missing}"
            )

    return [payload for payload in ordered if str(payload["period"]) not in existing_periods]


def next_append_row(sheet) -> int:
    last_value_row = 0
    for row in sheet.iter_rows():
        if any(cell.value is not None for cell in row):
            last_value_row = row[0].row
    return 1 if last_value_row == 0 else last_value_row + 3


def new_pdf_files(pdf_files: list[Path], existing_periods: list[str], extract_report_date) -> list[Path]:
    existing = set(existing_periods)
    selected: list[Path] = []
    seen: set[str] = set()
    for pdf_path in pdf_files:
        period = extract_report_date(pdf_path)
        if not period:
            raise RuntimeError(f"PDF rapor tarihi okunamadi: {pdf_path}")
        if period in seen:
            raise RuntimeError(f"Ayni donem icin birden fazla PDF bulundu: {period}")
        seen.add(period)
        if period not in existing:
            selected.append(pdf_path)
    return selected


def create_blank_workbook(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    workbook.save(path)
    workbook.close()
