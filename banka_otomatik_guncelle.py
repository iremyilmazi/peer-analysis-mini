from __future__ import annotations

import argparse
import copy
import difflib
import hashlib
import importlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from dataclasses import asdict, dataclass, field, replace
from datetime import date, datetime
from pathlib import Path
from typing import Callable

from openpyxl import load_workbook


PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_CONFIG = PROJECT_ROOT / "kap_config.json"
DEFAULT_WORKBOOK = PROJECT_ROOT / "butunlesik_banka_verisi.xlsx"
PERIOD_RE = re.compile(r"^(31\.03|30\.06|30\.09|31\.12)\.\d{4}$")
QUARTER_RE = re.compile(r"^(20\d{2})Q([1-4])$", re.IGNORECASE)
NUMBER_FORMAT = "#,##0;[Red](#,##0);-"

# Enpara's PDF extractor cannot distinguish an explicit dash from an empty,
# non-applicable cell.  In the fixed integrated layout these are the cells
# where the published Q2 dash is semantically zero.
ENPARA_Q2_ZERO_CELLS = {
    *(f"BZ{row}" for row in (270, 271, 272, 291, 292, 293, 297, 303, 306, 318, 319, 320, 324, 326)),
    *(f"BY{row}" for row in (360, 361, 363, 364, 366, 372, 394, 397, 398, 400, 401, 402, 411, 449)),
    *(f"BZ{row}" for row in (397, 398, 400, 401, 402, 409, 418, 448, 471, 473, 477, 482)),
    "CA395",
}
YAPIKREDI_Q2_ZERO_CELLS = {
    *(f"BY{row}" for row in (259, 260, 270, 337, 338, 339, 340, 346, 372, 376, 377, 378, 379, 395, 426, 443, 453)),
    *(f"BZ{row}" for row in (279, 280, 282, 299, 301, 302, 425, 449, 460)),
    "CA373",
}

# Published PDF layouts occasionally label the single-value exchange-rate
# column as "Toplam" although the integrated workbook's semantic header is
# "Kur".  Keep these audited bank/period coordinates explicit.
STRUCTURAL_HEADER_OVERRIDES = {
    "denizbank": {"CA488": "Kur"},
    "enpara": {"BZ447": "Yakın İzlemedeki Krediler"},
    "hsbc": {"CA487": "Kur"},
}

Q2_BLANK_OVERRIDES = {
    "vakifbank": {"CA215", "CA247", "CA248", "CA249"},
}

@dataclass(frozen=True)
class BankSpec:
    key: str
    display_name: str
    company_name: str
    folder: str
    downloader_module: str
    pdf_module: str
    pdf_function: str
    workers: int = 1
    postprocess: str | None = None

    @property
    def history_dir(self) -> Path:
        return PROJECT_ROOT / self.folder / "data" / "history"

    @property
    def cache_dir(self) -> Path:
        return PROJECT_ROOT / self.folder / "data" / "pdf_cache"


BANKS: dict[str, BankSpec] = {
    "akbank": BankSpec("akbank", "Akbank", "akbank", "akbank", "akbank.akbank_kap_history_indir", "akbank.akbank_dipnot_guncelle", "append_akbank_pdf_notes"),
    "denizbank": BankSpec("denizbank", "DenizBank", "denizbank", "denizbank", "denizbank.denizbank_kap_history_indir", "denizbank.denizbank_dipnot_guncelle", "append_denizbank_pdf_notes"),
    "enpara": BankSpec("enpara", "Enpara Bank", "enpara-bank", "enpara_bankasi", "enpara_bankasi.enpara_kap_history_indir", "enpara_bankasi.enpara_dipnot_guncelle", "append_enpara_pdf_notes"),
    "garanti": BankSpec("garanti", "Garanti BBVA", "garanti-bankasi", "garanti_bankasi", "garanti_bankasi.garanti_kap_history_indir", "garanti_bankasi.garanti_dipnot_guncelle", "append_garanti_pdf_notes", 4),
    "halkbank": BankSpec("halkbank", "Halkbank", "halkbank", "halkbank", "halkbank.halkbank_kap_history_indir", "halkbank.halkbank_dipnot_guncelle", "append_halkbank_pdf_notes"),
    "hsbc": BankSpec("hsbc", "HSBC", "hsbc", "hsbc", "hsbc.hsbc_kap_history_indir", "hsbc.hsbc_dipnot_guncelle", "append_hsbc_pdf_notes"),
    "ing": BankSpec("ing", "ING Bank", "ing-bank", "ing_bank", "ing_bank.ing_kap_history_indir", "ing_bank.ing_dipnot_guncelle", "append_ing_pdf_notes", 4, "ing"),
    "isbank": BankSpec("isbank", "Is Bankasi", "is-bankasi", "is_bankasi", "is_bankasi.is_bankasi_kap_history_indir", "is_bankasi.is_bankasi_dipnot_guncelle", "append_is_bankasi_pdf_notes", 4),
    "qnb": BankSpec("qnb", "QNB", "qnb-bank", "qnb_bank", "qnb_bank.qnb_kap_history_indir", "qnb_bank.qnb_dipnot_guncelle", "append_qnb_pdf_notes"),
    "teb": BankSpec("teb", "TEB", "turk-ekonomi-bankasi", "teb_bankasi", "teb_bankasi.teb_kap_history_indir", "teb_bankasi.teb_dipnot_guncelle", "append_teb_pdf_notes", 4),
    "vakifbank": BankSpec("vakifbank", "VakifBank", "vakifbank", "vakifbank", "vakifbank.vakifbank_kap_history_indir", "vakifbank.vakifbank_dipnot_guncelle", "append_vakifbank_pdf_notes"),
    "yapikredi": BankSpec("yapikredi", "Yapi Kredi", "yapi-kredi-bankasi", "yapi_kredi_bankasi", "yapi_kredi_bankasi.yapi_kredi_kap_history_indir", "yapi_kredi_bankasi.yapi_kredi_dipnot_guncelle", "append_yapi_kredi_pdf_notes"),
    "ziraat": BankSpec("ziraat", "Ziraat Bankasi", "ziraat-bankasi", "ziraat_bankasi", "ziraat_bankasi.ziraat_kap_history_indir", "ziraat_bankasi.ziraat_dipnot_guncelle", "append_ziraat_pdf_notes"),
}


@dataclass
class AuditReport:
    bank: str
    started_at: str
    workbook: str
    dry_run: bool
    status: str = "RUNNING"
    new_periods: list[str] = field(default_factory=list)
    downloaded_quarters: list[str] = field(default_factory=list)
    steps: list[dict] = field(default_factory=list)
    checked_cells: int = 0
    allowed_changes: int = 0
    output: str | None = None
    backup: str | None = None
    error: str | None = None


def log(step: str, message: str) -> None:
    print(f"[{datetime.now():%H:%M:%S}] {step:<12} {message}", flush=True)


def timed(report: AuditReport, name: str, action: Callable[[], object]) -> object:
    started = time.perf_counter()
    log(name, "Basladi.")
    try:
        result = action()
    except Exception as exc:
        report.steps.append({"name": name, "status": "ERROR", "seconds": round(time.perf_counter() - started, 3), "error": f"{type(exc).__name__}: {exc}"})
        raise
    report.steps.append({"name": name, "status": "OK", "seconds": round(time.perf_counter() - started, 3)})
    log(name, f"Tamamlandi ({time.perf_counter() - started:.1f} sn).")
    return result


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def quarter_to_period(value: str) -> str:
    match = QUARTER_RE.fullmatch(value)
    if not match:
        raise ValueError(f"Gecersiz donem klasoru: {value}")
    year, quarter = match.groups()
    ending = {"1": "31.03", "2": "30.06", "3": "30.09", "4": "31.12"}[quarter]
    return f"{ending}.{year}"


def period_to_quarter(value: str) -> str:
    day_month, year = value.rsplit(".", 1)
    quarter = {"31.03": "1", "30.06": "2", "30.09": "3", "31.12": "4"}.get(day_month)
    if not quarter:
        raise ValueError(f"Gecersiz rapor donemi: {value}")
    return f"{year}Q{quarter}"


def next_period(value: str) -> str:
    quarter = period_to_quarter(value)
    year, number = int(quarter[:4]), int(quarter[-1])
    return quarter_to_period(f'{year + 1}Q1' if number == 4 else f'{year}Q{number + 1}')


def workbook_periods(path: Path, sheet_name: str) -> list[str]:
    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        if sheet_name not in workbook.sheetnames:
            raise RuntimeError(f"Sheet bulunamadi: {sheet_name}")
        result: list[str] = []
        for cell in workbook[sheet_name][1]:
            value = cell.value
            if isinstance(value, str) and PERIOD_RE.fullmatch(value) and value not in result:
                result.append(value)
        return result
    finally:
        workbook.close()


def validate_workbook(path: Path, sheet_name: str) -> None:
    if not path.is_file():
        raise FileNotFoundError(f"Ana Excel bulunamadi: {path}")
    if path.suffix.lower() not in {".xlsx", ".xlsm"} or not zipfile.is_zipfile(path):
        raise ValueError(f"Gecerli bir Excel paketi degil: {path}")
    with zipfile.ZipFile(path) as archive:
        bad = archive.testzip()
        if bad:
            raise ValueError(f"Excel ZIP parcasi bozuk: {bad}")
    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        if sheet_name not in workbook.sheetnames:
            raise RuntimeError(f"Ana Excel icinde banka sheet'i yok: {sheet_name}")
    finally:
        workbook.close()


def _contains_pdf_header(path: Path) -> bool:
    # KAP'in bazi ekleri Java serialization zarfi ile gelir; gercek PDF
    # basligi ilk birkac yuz baytin icindedir.
    with path.open("rb") as stream:
        return b"%PDF-" in stream.read(4096)


def validate_pdf(path: Path) -> None:
    if not path.is_file() or path.stat().st_size < 10_000:
        raise ValueError(f"PDF bos veya cok kucuk: {path}")
    if not _contains_pdf_header(path):
        raise ValueError(f"Indirilen ek gercek PDF degil: {path}")


def validate_excel_source(path: Path) -> None:
    if not path.is_file() or path.stat().st_size < 10_000:
        raise ValueError(f"KAP Excel bos veya cok kucuk: {path}")
    head = path.read_bytes()[:4096]
    lower = head.lower()
    valid = head.startswith(b"PK") or head.startswith(bytes.fromhex("D0CF11E0")) or b"<html" in lower or b"<?xml" in lower
    if not valid:
        raise ValueError(f"Indirilen ek Excel degil: {path}")
    if b"<form" in lower and (b"login" in lower or b"giris" in lower):
        raise ValueError(f"Excel yerine giris/guvenlik sayfasi indi: {path}")


def validate_quarter_dir(history_dir: Path, quarter: str) -> tuple[Path, Path]:
    folder = history_dir / quarter
    excels = sorted((p for p in folder.glob("*.xls*") if p.is_file()), key=lambda p: (p.stat().st_size, p.name), reverse=True)
    pdfs = sorted((p for p in folder.glob("*.pdf") if p.is_file()), key=lambda p: (p.stat().st_size, p.name), reverse=True)
    if not excels or not pdfs:
        raise RuntimeError(f"{quarter} klasorunde hem Excel hem PDF bulunmali: {folder}")
    validate_excel_source(excels[0])
    for pdf in pdfs:
        validate_pdf(pdf)
    # Birden fazla dil/ek varsa finansal raporun en kapsamli kopyasi genelde
    # en buyuk PDF'tir; tum adaylar yukarida gercek PDF olarak dogrulandi.
    return excels[0], pdfs[0]


def complete_history_periods(history_dir: Path) -> set[str]:
    result: set[str] = set()
    if not history_dir.exists():
        return result
    for child in history_dir.iterdir():
        if child.is_dir() and QUARTER_RE.fullmatch(child.name):
            try:
                validate_quarter_dir(history_dir, child.name)
            except Exception:
                continue
            result.add(quarter_to_period(child.name))
    return result


def download_reports(
    spec: BankSpec,
    history_dir: Path,
    today: date,
    force: bool,
    after_period: str | None = None,
) -> list[str]:
    module = importlib.import_module(spec.downloader_module)
    # Bu hafif paket history tasimaz. Mevcut ana Excel'in son donemi,
    # ilk calismada KAP'tan eski history'nin tekrar indirilmesini engeller.
    # Banka adaptorlerinin tamami TEB tabanini kullandigindan override
    # degeri temel modulu; Garanti icin ise modulu hedefler.
    after_quarter = period_to_quarter(after_period) if after_period else None
    target_module = getattr(module, "_base", module)
    previous_after_quarter = getattr(target_module, "INCREMENTAL_AFTER_QUARTER", None)
    target_module.INCREMENTAL_AFTER_QUARTER = after_quarter
    try:
        quarters = list(module.run_incremental(history_dir, today=today, force=force))
    finally:
        target_module.INCREMENTAL_AFTER_QUARTER = previous_after_quarter
    for quarter in quarters:
        validate_quarter_dir(history_dir, quarter)
    return quarters


def update_excel(spec: BankSpec, config: Path, candidate: Path, history_dir: Path) -> None:
    from entegre_excel_olustur import build_integrated_workbook, load_config, select_companies

    companies, source_dir, _ = load_config(config)
    selected = select_companies(companies, spec.company_name)
    if not selected:
        raise RuntimeError(f"Config banka kaydi bulunamadi: {spec.company_name}")
    # Kullanici farkli bir history klasoru verdiyse sadece gecici config
    # nesnesindeki kaynak dizini degistiririz; kalici config'e dokunmayiz.
    selected = [replace(selected[0], source_dir=str(history_dir))]
    results = build_integrated_workbook(selected, source_dir, candidate, update_index=False)
    for item in results:
        log(
            "Excel",
            f"{item.company_name}: {item.rows_written} satir; {item.message or 'uyari yok'}; "
            f"kaynak={item.source_file}",
        )
    errors = [item for item in results if item.status in {"ERROR", "SKIP"}]
    if errors:
        raise RuntimeError("; ".join(f"{item.status}: {item.message}" for item in errors))

    if spec.postprocess == "ing":
        # Run ING's canonical row rebuild in a clean interpreter.  Several
        # bank-specific PDF modules monkey-patch the shared parser; process
        # isolation prevents those globals from affecting the four-column P&L
        # mapping and mirrors how the portable per-bank script is executed.
        subprocess.run(
            [
                sys.executable,
                str(PROJECT_ROOT / "ing_bank" / "ing_excel_guncelle.py"),
                "--workbook", str(candidate),
                "--history-dir", str(history_dir),
            ],
            cwd=PROJECT_ROOT,
            check=True,
        )


def update_pdf(spec: BankSpec, candidate: Path, history_dir: Path, cache_dir: Path, new_periods: list[str], workers: int) -> None:
    module = importlib.import_module(spec.pdf_module)
    function = getattr(module, spec.pdf_function)
    selected_pdfs: list[tuple[str, Path]] = []
    for period in sorted(new_periods, key=lambda p: tuple(reversed([int(x) for x in p.split(".")]))):
        quarter = period_to_quarter(period)
        _excel, pdf = validate_quarter_dir(history_dir, quarter)
        selected_pdfs.append((quarter, pdf))

    # Only rebuild a scratch PDF block for the incoming periods.  The final
    # merge copies just those period columns into the user's established row
    # layout, so historical PDF data is neither reparsed nor rewritten.
    with tempfile.TemporaryDirectory(prefix=f"{spec.key}_new_pdf_") as directory:
        latest_dir = Path(directory)
        for quarter, pdf in selected_pdfs:
            period_dir = latest_dir / quarter
            period_dir.mkdir(parents=True, exist_ok=True)
            destination = period_dir / pdf.name
            try:
                os.link(pdf, destination)
            except OSError:
                shutil.copy2(pdf, destination)
        source, rows = function(
            workbook_path=candidate,
            pdf_dir=latest_dir,
            pdf_path=None,
            sheet_name=spec.company_name,
            cache_dir=cache_dir,
            workers=min(workers, max(1, len(selected_pdfs))),
        )
    if rows <= 0:
        raise RuntimeError(f"PDF dipnot aktarimi satir uretmedi: {source}")
    log("PDF", f"Yalniz yeni donem {', '.join(new_periods)}: {Path(source).name} -> {rows} satir.")

    if spec.postprocess == "ing":
        ing_main = importlib.import_module("ing_bank.ing_guncelle")
        ing_main.finalize_ing_sheet(candidate, had_index=True)


def _label_key(value: object) -> str:
    text = " ".join(str(value or "").split()).casefold()
    text = text.translate(str.maketrans({"ı": "i", "İ": "i", "ş": "s", "Ş": "s", "ğ": "g", "Ğ": "g", "ü": "u", "Ü": "u", "ö": "o", "Ö": "o", "ç": "c", "Ç": "c"}))
    key = "".join(character for character in text if character.isalnum())
    # Several bank PDFs prepend one or two comparison-period dates to the first
    # row label (for example "30 Haziran 2026 30 Haziran 2025 Kasa/Efektif").
    # Those dates are column headings, not part of the semantic row identity.
    # The KAP HTML/PDF readers of a few legacy files contain the Unicode
    # replacement character.  Once punctuation is removed, e.g. "Aral�k"
    # becomes "aralk"; include those stable degraded spellings as well.
    months = "ocak|subat|ubat|mart|nisan|mayis|mays|haziran|temmuz|agustos|austos|eylul|eyll|ekim|kasim|kasm|aralik|aralk"
    key = re.sub(rf"^(?:\d{{1,2}}(?:{months})\d{{4}})+", "", key)
    return key


def _row_identity_map(sheet, aliases: dict[str, str] | None = None) -> dict[tuple[str, str, int], int]:
    aliases = aliases or {}
    labels = {
        row: aliases.get(_label_key(sheet.cell(row, 1).value), _label_key(sheet.cell(row, 1).value))
        for row in range(1, sheet.max_row + 1)
    }
    section = "__top__"
    occurrences: dict[tuple[str, str], int] = {}
    result: dict[tuple[str, str, int], int] = {}
    for row in range(1, sheet.max_row + 1):
        label = labels[row]
        if not label:
            continue
        next_label = labels.get(row + 1, "")
        if next_label in {"kalem", "vade"} and row > 2:
            section = f"pdf:{label}"
        base = (section, label)
        occurrence = occurrences.get(base, 0) + 1
        occurrences[base] = occurrence
        result[(section, label, occurrence)] = row
    return result


def _row_map_by_label_occurrence(source_sheet, target_sheet, aliases: dict[str, str] | None = None) -> dict[int, int]:
    target_rows = _row_identity_map(target_sheet, aliases)
    source_rows = _row_identity_map(source_sheet, aliases)
    mapping: dict[int, int] = {}
    missing: list[str] = []
    # A PDF parser can preserve a date-prefixed version of the same semantic
    # label once per historical period.  After date normalization those rows
    # have different occurrence numbers, while the reference layout correctly
    # has one row.  A unique (section, label) target is therefore a safe
    # fallback; later/non-empty current-period data wins during the merge.
    targets_by_base: dict[tuple[str, str], list[int]] = {}
    for (section, label, _occurrence), row in target_rows.items():
        targets_by_base.setdefault((section, label), []).append(row)
    for identity, row in source_rows.items():
        target_row = target_rows.get(identity)
        if target_row is None:
            candidates = targets_by_base.get(identity[:2], [])
            if len(candidates) == 1:
                target_row = candidates[0]
        if target_row is None:
            source_section, source_label, source_occurrence = identity
            ranked: list[tuple[float, int]] = []
            for (target_section, target_label, target_occurrence), candidate_row in target_rows.items():
                label_score = difflib.SequenceMatcher(None, source_label, target_label).ratio()
                if label_score < 0.78:
                    continue
                if source_section == target_section:
                    section_score = 1.0
                elif source_section == "__top__" or target_section == "__top__":
                    continue
                else:
                    section_score = difflib.SequenceMatcher(None, source_section, target_section).ratio()
                    if section_score < 0.68:
                        continue
                occurrence_score = 1.0 / (1.0 + abs(source_occurrence - target_occurrence))
                ranked.append((0.72 * label_score + 0.23 * section_score + 0.05 * occurrence_score, candidate_row))
            ranked.sort(reverse=True)
            if ranked and (len(ranked) == 1 or ranked[0][0] - ranked[1][0] >= 0.025):
                target_row = ranked[0][1]
        if target_row is None:
            missing.append(f"{source_sheet.cell(row, 1).value!s} ({identity[0]})")
        else:
            mapping[row] = target_row
    if missing:
        preview = "; ".join(missing[:5])
        log(
            "Satir esleme",
            f"Mevcut desende olmayan {len(missing)} kaynak satiri guvenle atlandi: {preview}",
        )
    if len(mapping) < 10:
        raise RuntimeError(f"Satir esleme kapsami yetersiz: {len(mapping)} satir")
    return mapping


def _financial_number(value: object, multiplier: int) -> object:
    if isinstance(value, bool) or value is None:
        return value
    if isinstance(value, (int, float)):
        return value * multiplier
    text = str(value).strip().replace(" ", "")
    if text in {"", "-", "—", "–"}:
        return 0
    negative = text.startswith("(") and text.endswith(")")
    if negative:
        text = text[1:-1]
    if not re.fullmatch(r"[+-]?\d[\d.,]*", text):
        return value
    number = int(re.sub(r"[.,]", "", text))
    if negative:
        number = -number
    return number * multiplier


def _scale_value_for_period(spec: BankSpec, period: str, section_label: str, value: object) -> object:
    if spec.postprocess == "ing":
        return value
    day, month, year = (int(item) for item in period.split("."))
    if (year, month, day) < (2026, 6, 30):
        return value
    normalized_section = _label_key(section_label)
    if "likiditekarsilama" in normalized_section or "kurriskine" in normalized_section:
        return _financial_number(value, 1)
    return _financial_number(value, 1000)


def merge_new_period_columns(original_path: Path, rebuilt_path: Path, candidate_path: Path, spec: BankSpec, periods: set[str]) -> None:
    """Copy only new-period columns out of a fully rebuilt scratch workbook."""
    shutil.copy2(original_path, candidate_path)
    rebuilt = load_workbook(rebuilt_path, data_only=False)
    candidate = load_workbook(candidate_path, data_only=False)
    try:
        source_sheet = rebuilt[spec.company_name]
        target_sheet = candidate[spec.company_name]
        aliases = {}
        if spec.key in {"akbank", "denizbank", "enpara", "garanti", "isbank", "teb", "yapikredi", "ziraat"}:
            aliases = {
                "kredimusterilerininfonlari": "mustakrizlerinfonlari",
                "kredimterilerininfonlar": "mustakrizlerinfonlari",
                "mstakrizlerinfonlar": "mustakrizlerinfonlari",
            }
        if spec.key == "vakifbank":
            aliases.update({"tcmerkezbankasi": "tcmb"})
        if spec.key == "ziraat":
            # Ziraat's Q2 parser preserves hierarchical labels as
            # "parent - child", while the established workbook stores the
            # parent once and the children on separate repeated rows.  Map
            # only those explicit composite source labels to their child;
            # occurrence order within the same section keeps them distinct.
            for row in range(253, source_sheet.max_row + 1):
                raw_label = str(source_sheet.cell(row, 1).value or "")
                parent_key = _label_key(raw_label.split(" - ", 1)[0])
                hierarchical_parent = any(
                    token in parent_key
                    for token in (
                        "tuketicikredileri", "tketicikredileri",
                        "bireyselkredikartlari", "bireyselkredikartlar",
                        "personelkredileri", "personelkredikartlari", "personelkredikartlar",
                    )
                )
                if " - " in raw_label and hierarchical_parent:
                    aliases[_label_key(raw_label)] = _label_key(raw_label.rsplit(" - ", 1)[1])
        row_map = _row_map_by_label_occurrence(source_sheet, target_sheet, aliases)
        source_columns = [cell.column for cell in source_sheet[1] if cell.value in periods]
        if not source_columns:
            raise RuntimeError(f"Yeniden olusturulan sheet'te yeni donem yok: {sorted(periods)}")

        for source_column in source_columns:
            # Standart donem diziliminde hedef sutun numarasi degismez. Hedefte
            # baska bir donem varsa yanlis yere yazmak yerine durdur.
            existing_header = target_sheet.cell(1, source_column).value
            source_header = source_sheet.cell(1, source_column).value
            if existing_header not in {None, "", source_header}:
                raise RuntimeError(
                    f"Yeni donem sutunu cakismasi: {target_sheet.cell(1, source_column).coordinate} "
                    f"{existing_header!r} / {source_header!r}"
                )
            # Bos yeni donem alanlari kullanici dosyasinda henuz
            # bicimlendirilmemis olabilir. Yalniz hic stili olmayan
            # hucrelere onceki ceyregin stilini uygula; hazir stilleri koru.
            for target_row in range(1, target_sheet.max_row + 1):
                new_cell = target_sheet.cell(target_row, source_column)
                prior_cell = target_sheet.cell(target_row, source_column - 3)
                if not new_cell.has_style and prior_cell.has_style:
                    new_cell._style = copy.copy(prior_cell._style)
            for source_row, target_row in row_map.items():
                source_cell = source_sheet.cell(source_row, source_column)
                row_label = _label_key(source_sheet.cell(source_row, 1).value)
                effective_target_column = source_column
                garanti_fx_row = spec.key == "garanti" and row_label in {
                    "usdcaridovizaliskuru", "usdaritmetikortalama"
                }
                if garanti_fx_row:
                    if (source_column - 2) % 3 != 0:
                        continue
                    effective_target_column = source_column + 2
                target_cell = target_sheet.cell(target_row, effective_target_column)
                # The target workbook is the layout authority.  For a newly
                # opened period, structural PDF subheaders are copied from the
                # preceding quarter in the target layout instead of accepting
                # parser-specific spellings such as Kur/Toplam or Yakin/Yakın.
                if source_row > 2 and row_label in {"kalem", "vade"}:
                    prior_cell = target_sheet.cell(target_row, effective_target_column - 3)
                    target_cell.value = STRUCTURAL_HEADER_OVERRIDES.get(spec.key, {}).get(
                        target_cell.coordinate, prior_cell.value
                    )
                    continue
                section = ""
                # Baslik satirinin kimligi tablo adini tasir; diger satirlarda
                # kaynak satirin ait oldugu en yakin PDF basligini bul.
                for scan_row in range(source_row, 2, -1):
                    if _label_key(source_sheet.cell(scan_row + 1, 1).value) in {"kalem", "vade"}:
                        section = str(source_sheet.cell(scan_row, 1).value or "")
                        break
                period = str(source_header)
                if garanti_fx_row:
                    if isinstance(source_cell.value, (int, float)):
                        target_cell.value = source_cell.value
                    else:
                        text_value = str(source_cell.value or "").strip().replace(".", "").replace(",", ".")
                        target_cell.value = float(text_value) if text_value else None
                elif source_row <= 2:
                    target_cell.value = source_cell.value
                elif source_cell.value is None and source_column > 4:
                    # A dash in KAP tables is normalized to an empty source
                    # cell by some bank parsers.  When the same breakdown was
                    # numeric in the preceding quarter it semantically means
                    # zero; otherwise retain a genuine unavailable blank.
                    explicit_zero = (
                        spec.key == "enpara" and target_cell.coordinate in ENPARA_Q2_ZERO_CELLS
                    ) or (
                        spec.key == "yapikredi" and target_cell.coordinate in YAPIKREDI_Q2_ZERO_CELLS
                    )
                    target_cell.value = 0 if explicit_zero else None
                else:
                    source_value = source_cell.value
                    structural_blank = (
                        spec.key == "qnb" and target_row in {16, 26, 39, 61, 76, 77, 167}
                    ) or (spec.key == "teb" and target_row == 167)
                    if structural_blank:
                        source_value = None
                    target_cell.value = _scale_value_for_period(spec, period, section, source_value)
                    if target_cell.coordinate in Q2_BLANK_OVERRIDES.get(spec.key, set()):
                        target_cell.value = None
                # Do not touch formatting: the integrated workbook is the
                # pattern authority and its reserved period cells already
                # carry bank/row-specific number formats, fills and fonts.
            letter = source_sheet.cell(1, source_column).column_letter
            previous_letter = target_sheet.cell(1, source_column - 3).column_letter
            if previous_letter in target_sheet.column_dimensions:
                dimension = copy.copy(target_sheet.column_dimensions[previous_letter])
                dimension.index = letter
                dimension.min = source_column
                dimension.max = source_column
                target_sheet.column_dimensions[letter] = dimension
            elif letter not in target_sheet.column_dimensions:
                # Onceki donemin genisligi B:BX gibi tek bir grup tanimi
                # altinda olabilir. Yeni sutunlar bu grubun disinda kalir.
                # Tarih ve binlik ayiracli tutarlar icin okunabilir genislik.
                prior_width = next((
                    item.width for item in target_sheet.column_dimensions.values()
                    if item.min is not None and item.max is not None
                    and item.min <= source_column - 3 <= item.max
                ), 13)
                dimension = target_sheet.column_dimensions[letter]
                dimension.width = max(prior_width or 13, 20)
                dimension.min = source_column
                dimension.max = source_column
        candidate.save(candidate_path)
    finally:
        rebuilt.close()
        candidate.close()


def _color(color) -> tuple | None:
    if color is None:
        return None
    return color.type, color.rgb, color.indexed, color.theme, color.tint


def _side(side) -> tuple:
    if side is None:
        return None, None
    return side.style, _color(side.color)


def cell_signature(cell) -> tuple:
    font, fill, alignment, border, protection = cell.font, cell.fill, cell.alignment, cell.border, cell.protection
    value = None if cell.value in {None, ""} else cell.value
    data_type = "n" if value is None else cell.data_type
    return (
        value,
        data_type,
        cell.number_format,
        (font.name, font.sz, bool(font.b), bool(font.i), font.underline, bool(font.strike), _color(font.color)),
        (fill.fill_type, _color(fill.fgColor), _color(fill.bgColor)),
        (alignment.horizontal, alignment.vertical, alignment.text_rotation, bool(alignment.wrap_text), alignment.indent),
        (_side(border.left), _side(border.right), _side(border.top), _side(border.bottom)),
        (bool(protection.locked), bool(protection.hidden)),
    )


def _dimension_signature(dimension) -> tuple:
    if dimension is None:
        return (None, None, False, 0, False)
    width = getattr(dimension, "width", None)
    height = getattr(dimension, "height", None)
    return (
        round(width, 6) if isinstance(width, float) else width,
        round(height, 6) if isinstance(height, float) else height,
        bool(dimension.hidden),
        dimension.outlineLevel or 0,
        bool(dimension.collapsed),
    )


def audit_candidate(original_path: Path, candidate_path: Path, sheet_name: str, new_periods: set[str]) -> tuple[int, int]:
    original = load_workbook(original_path, data_only=False)
    candidate = load_workbook(candidate_path, data_only=False)
    checked = allowed = 0
    try:
        if original.sheetnames != candidate.sheetnames:
            raise RuntimeError("Sheet listesi veya sirasi degisti.")
        target = candidate[sheet_name]
        allowed_columns = {cell.column for cell in target[1] if cell.value in new_periods}
        if new_periods and not allowed_columns:
            raise RuntimeError(f"Yeni donem kolonlari bulunamadi: {sorted(new_periods)}")

        for name in original.sheetnames:
            left, right = original[name], candidate[name]
            if left.freeze_panes != right.freeze_panes or left.sheet_view.showGridLines != right.sheet_view.showGridLines:
                raise RuntimeError(f"Sheet gorunum deseni degisti: {name}")
            if {str(item) for item in left.merged_cells.ranges} != {str(item) for item in right.merged_cells.ranges}:
                raise RuntimeError(f"Birlesik hucre deseni degisti: {name}")
            for index in set(left.row_dimensions) | set(right.row_dimensions):
                if _dimension_signature(left.row_dimensions.get(index)) != _dimension_signature(right.row_dimensions.get(index)):
                    raise RuntimeError(f"Satir bicimi degisti: {name}!{index}")
            for key in set(left.column_dimensions) | set(right.column_dimensions):
                right_dimension = right.column_dimensions.get(key)
                if name == sheet_name and right_dimension is not None and right_dimension.min in allowed_columns:
                    continue
                if _dimension_signature(left.column_dimensions.get(key)) != _dimension_signature(right_dimension):
                    raise RuntimeError(f"Sutun bicimi degisti: {name}!{key}")

            max_row = max(left.max_row, right.max_row)
            max_column = max(left.max_column, right.max_column)
            for row in range(1, max_row + 1):
                for column in range(1, max_column + 1):
                    checked += 1
                    before = cell_signature(left.cell(row, column))
                    after = cell_signature(right.cell(row, column))
                    if before == after:
                        continue
                    if name == sheet_name and column in allowed_columns:
                        allowed += 1
                        continue
                    raise RuntimeError(
                        f"History/veri deseni degisikligi: {name}!{right.cell(row, column).coordinate} "
                        f"{before[0]!r} -> {after[0]!r}"
                    )
        return checked, allowed
    finally:
        original.close()
        candidate.close()


def commit(candidate: Path, workbook: Path, backup_root: Path, bank_key: str) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_dir = backup_root / bank_key
    backup_dir.mkdir(parents=True, exist_ok=True)
    backup = backup_dir / f"{workbook.stem}_ONCESI_{stamp}{workbook.suffix}"
    shutil.copy2(workbook, backup)
    replacement = workbook.with_name(f".{workbook.name}.{bank_key}.tmp")
    shutil.copy2(candidate, replacement)
    os.replace(replacement, workbook)
    return backup


def write_report(report: AuditReport, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(asdict(report), ensure_ascii=False, indent=2), encoding="utf-8")


def run_bank(spec: BankSpec, args: argparse.Namespace) -> int:
    workbook = Path(args.workbook).resolve()
    config = Path(args.config).resolve()
    history_dir = Path(args.history_dir or spec.history_dir).resolve()
    cache_dir = Path(args.cache_dir or spec.cache_dir).resolve()
    report_path = Path(args.audit_report or (PROJECT_ROOT / "outputs" / "otomatik_guncelleme_audit" / f"{spec.key}_son_calisma.json")).resolve()
    report = AuditReport(spec.display_name, datetime.now().isoformat(timespec="seconds"), str(workbook), bool(args.dry_run))
    original_hash = None
    debug_dir = Path(args.debug_dir).resolve() if args.debug_dir else None

    def debug_copy(source: Path, name: str) -> None:
        if debug_dir is None:
            return
        debug_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, debug_dir / name)

    try:
        if not config.is_file():
            raise FileNotFoundError(f"KAP config bulunamadi: {config}")
        validate_workbook(workbook, spec.company_name)
        original_hash = sha256(workbook)
        old_periods = set(workbook_periods(workbook, spec.company_name))
        if not old_periods:
            raise RuntimeError(f'{spec.display_name} sayfasinda mevcut donem bulunamadi; gecmis indirme baslatilmayacak.')
        latest_existing_period = max(old_periods, key=period_to_quarter)
        expected_period = next_period(latest_existing_period)
        if args.period and args.period != expected_period:
            raise ValueError(f'Yalniz bir sonraki donem islenebilir: {expected_period}')
        log("Baslangic", f"{spec.display_name}; ana Excel gecici kopyada islenecek.")
        log('Hedef', f'Excel son donem: {latest_existing_period}; yalniz hedef: {expected_period}')

        if args.skip_download:
            downloaded: list[str] = []
            available = complete_history_periods(history_dir)
            new_periods = [expected_period] if expected_period in available else []
            log("KAP", "Indirme atlandi; lokal history dogrulandi.")
        else:
            downloaded = timed(
                report,
                "1-KAP",
                lambda: download_reports(
                    spec,
                    history_dir,
                    date.fromisoformat(args.today) if args.today else date.today(),
                    args.force_download,
                    latest_existing_period,
                ),
            )
            new_periods = [quarter_to_period(item) for item in downloaded]
            if new_periods and new_periods != [expected_period]:
                raise RuntimeError(f'KAP beklenen tek donemden farkli kaynak dondurdu: {new_periods}')

        if args.period:
            requested = {args.period}
            missing = requested - complete_history_periods(history_dir)
            if missing:
                raise RuntimeError(f"Istenen donemin lokal PDF+Excel kaynagi eksik: {sorted(missing)}")
            new_periods = sorted(requested - old_periods)

        report.downloaded_quarters = downloaded
        report.new_periods = new_periods
        if not new_periods:
            report.status = "NO_NEW_PERIOD"
            log("Sonuc", "Yeni donem yok; ana Excel degistirilmedi.")
            if sha256(workbook) != original_hash:
                raise RuntimeError("Yeni donem yokken ana Excel hash'i degisti.")
            write_report(report, report_path)
            return 0

        history_dir.mkdir(parents=True, exist_ok=True)
        cache_dir.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix=f"{spec.key}_update_") as temp_dir:
            rebuilt = Path(temp_dir) / f"rebuilt_{workbook.name}"
            candidate = Path(temp_dir) / f"candidate_{workbook.name}"
            shutil.copy2(workbook, rebuilt)
            period_sources = history_dir / period_to_quarter(expected_period)
            timed(report, "2-EXCEL", lambda: update_excel(spec, config, rebuilt, period_sources))
            debug_copy(rebuilt, "01_excel_asamasi.xlsx")
            timed(report, "3-PDF", lambda: update_pdf(spec, rebuilt, history_dir, cache_dir, new_periods, args.workers or spec.workers))
            debug_copy(rebuilt, "02_pdf_asamasi.xlsx")
            timed(report, "3B-BIRLESTIR", lambda: merge_new_period_columns(workbook, rebuilt, candidate, spec, set(new_periods)))
            debug_copy(candidate, "03_birlestirme_adayi.xlsx")
            validate_workbook(candidate, spec.company_name)
            candidate_periods = set(workbook_periods(candidate, spec.company_name))
            if set(new_periods) - candidate_periods:
                raise RuntimeError(f"Yeni donem Excel'e yazilamadi: {sorted(set(new_periods) - candidate_periods)}")

            checked, allowed = timed(report, "4-KONTROL", lambda: audit_candidate(workbook, candidate, spec.company_name, set(new_periods)))
            report.checked_cells, report.allowed_changes = checked, allowed
            if allowed == 0:
                raise RuntimeError("Yeni donem icin hicbir hucre degisikligi bulunamadi.")
            if sha256(workbook) != original_hash:
                raise RuntimeError('Islem surerken ana Excel baska bir islem tarafindan degisti; kayit iptal edildi.')

            if args.dry_run:
                output = Path(args.dry_run_output or workbook.with_name(f"{workbook.stem}_{spec.key}_DRY_RUN{workbook.suffix}")).resolve()
                if output == workbook:
                    raise ValueError('Guvenli deneme cikti yolu ana Excel ile ayni olamaz.')
                shutil.copy2(candidate, output)
                report.output = str(output)
                log("Dry-run", f"Ana Excel degismedi; kontrol kopyasi: {output}")
            else:
                backup_root = Path(
                    args.backup_root or (PROJECT_ROOT / "outputs" / "otomatik_guncelleme_yedek")
                ).resolve()
                backup = commit(candidate, workbook, backup_root, spec.key)
                report.output, report.backup = str(workbook), str(backup)
                log("Kayit", f"Atomik kayit tamamlandi; yedek: {backup}")

        if args.dry_run and sha256(workbook) != original_hash:
            raise RuntimeError("Dry-run sirasinda ana Excel degisti.")
        report.status = "OK"
        write_report(report, report_path)
        log("Tamamlandi", f"Audit raporu: {report_path}")
        return 0
    except Exception as exc:
        report.status = "ERROR"
        report.error = f"{type(exc).__name__}: {exc}"
        if original_hash and workbook.exists() and sha256(workbook) != original_hash:
            report.error += " | KRITIK: ana Excel hash'i degisti"
        write_report(report, report_path)
        log("HATA", f"Ana Excel degistirilmedi: {report.error}")
        return 1


def build_parser(default_bank: str | None = None) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Banka bazli guvenli KAP Excel ve PDF otomatik guncelleme akisi.")
    if default_bank is None:
        parser.add_argument("bank", choices=sorted(BANKS))
    parser.add_argument("--config", default=str(DEFAULT_CONFIG))
    parser.add_argument("--workbook", default=str(DEFAULT_WORKBOOK))
    parser.add_argument("--history-dir")
    parser.add_argument("--cache-dir")
    parser.add_argument("--workers", type=int)
    parser.add_argument("--today", help="KAP sorgusu icin test tarihi (YYYY-MM-DD).")
    parser.add_argument("--period", help="Yalniz bu lokal donemi isle (ornek: 30.06.2026).")
    parser.add_argument("--force-download", action="store_true")
    parser.add_argument("--skip-download", action="store_true")
    parser.add_argument("--dry-run", action="store_true", help="Ana Excel'i degistirmeden tum adimlari ve audit'i calistir.")
    parser.add_argument("--dry-run-output")
    parser.add_argument("--audit-report")
    parser.add_argument("--backup-root", help="Yedek klasoru; testlerde gecici bir dizin verilebilir.")
    parser.add_argument("--debug-dir", help="Yalniz test/teshis icin ara asama kopyalarini saklar.")
    return parser


def main_for_bank(bank_key: str) -> int:
    if bank_key not in BANKS:
        raise SystemExit(f"Bilinmeyen banka: {bank_key}")
    args = build_parser(default_bank=bank_key).parse_args()
    return run_bank(BANKS[bank_key], args)


def main() -> int:
    args = build_parser().parse_args()
    return run_bank(BANKS[args.bank], args)


if __name__ == "__main__":
    raise SystemExit(main())
