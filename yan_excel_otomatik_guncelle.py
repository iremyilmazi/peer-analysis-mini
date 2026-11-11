from __future__ import annotations

import argparse
import copy
import importlib
import json
import os
import shutil
import tempfile
from dataclasses import asdict, dataclass, field
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

import banka_otomatik_guncelle as main_update
from yan_excel_donem_ekle import new_pdf_files, workbook_periods


ROOT = Path(__file__).resolve().parent
PRODUCTS = {
    "stage": ("birinci_ikinci_asama_krediler.xlsx", "Kredi asamalari"),
    "deposit": ("mevduat_faiz.xlsx", "Mevduat faizi"),
}
SIDE_MODULES = {
    "akbank": ("akbank.akbank_yan_excel_guncelle", "akbank"),
    "denizbank": ("denizbank.denizbank_yan_excel_guncelle", "denizbank"),
    "enpara": ("enpara_bankasi.enpara_yan_excel_guncelle", "enpara"),
    "garanti": ("garanti_bankasi.garanti_yan_excel_guncelle", "garanti"),
    "halkbank": ("halkbank.halkbank_yan_excel_guncelle", "halkbank"),
    "hsbc": ("hsbc.hsbc_yan_excel_guncelle", "hsbc"),
    "isbank": ("is_bankasi.is_bankasi_yan_excel_guncelle", "is-bankasi"),
    "qnb": ("qnb_bank.qnb_yan_excel_guncelle", "qnb-bank"),
    "teb": ("teb_bankasi.teb_yan_excel_guncelle", "teb"),
    "vakifbank": ("vakifbank.vakifbank_yan_excel_guncelle", "vakifbank"),
    "yapikredi": ("yapi_kredi_bankasi.yapi_kredi_yan_excel_guncelle", "yapi-kredi"),
    "ziraat": ("ziraat_bankasi.ziraat_yan_excel_guncelle", "ziraat"),
}


@dataclass
class SideAuditReport:
    bank: str
    product: str
    workbook: str
    dry_run: bool
    started_at: str
    status: str = "RUNNING"
    downloaded_quarters: list[str] = field(default_factory=list)
    new_periods: list[str] = field(default_factory=list)
    checked_cells: int = 0
    allowed_changes: int = 0
    output: str | None = None
    backup: str | None = None
    error: str | None = None


def log(step: str, message: str) -> None:
    print(f"[{datetime.now():%H:%M:%S}] {step:<12} {message}", flush=True)


def write_report(report: SideAuditReport, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(asdict(report), ensure_ascii=False, indent=2), encoding="utf-8")


def source_engine(bank: str):
    module_name, sheet_name = SIDE_MODULES[bank]
    module = importlib.import_module(module_name)
    configure = getattr(module, "configure_base", None)
    if callable(configure):
        configure()
        module = module._base
    return module, sheet_name


def last_used_row(sheet) -> int:
    last_row = 0
    for row in sheet.iter_rows():
        if any(cell.value is not None for cell in row):
            last_row = row[0].row
    return last_row


def audit_candidate(original_path: Path, candidate_path: Path, sheet_name: str, expected_periods: list[str]) -> tuple[int, int]:
    original = load_workbook(original_path, data_only=False)
    candidate = load_workbook(candidate_path, data_only=False)
    checked = allowed = 0
    try:
        if original.sheetnames != candidate.sheetnames:
            raise RuntimeError("Sheet listesi veya sirasi degisti.")
        if sheet_name not in original.sheetnames:
            raise RuntimeError(f"Hedef banka sheet'i bulunamadi: {sheet_name}")
        original_last_row = last_used_row(original[sheet_name])
        candidate_periods = workbook_periods(candidate_path, sheet_name)
        missing = set(expected_periods) - set(candidate_periods)
        if missing:
            raise RuntimeError(f"Yeni donem yan Excel'e yazilamadi: {sorted(missing)}")

        for name in original.sheetnames:
            left, right = original[name], candidate[name]
            if left.freeze_panes != right.freeze_panes or left.sheet_view.showGridLines != right.sheet_view.showGridLines:
                raise RuntimeError(f"Sheet gorunumu degisti: {name}")
            if {str(item) for item in left.merged_cells.ranges} != {str(item) for item in right.merged_cells.ranges}:
                raise RuntimeError(f"Birlesik hucre deseni degisti: {name}")
            for row_index in set(left.row_dimensions) | set(right.row_dimensions):
                if name == sheet_name and row_index > original_last_row:
                    continue
                if main_update._dimension_signature(left.row_dimensions.get(row_index)) != main_update._dimension_signature(right.row_dimensions.get(row_index)):
                    raise RuntimeError(f"Satir bicimi degisti: {name}!{row_index}")
            for column_key in set(left.column_dimensions) | set(right.column_dimensions):
                if main_update._dimension_signature(left.column_dimensions.get(column_key)) != main_update._dimension_signature(right.column_dimensions.get(column_key)):
                    raise RuntimeError(f"Sutun bicimi degisti: {name}!{column_key}")
            max_row = max(left.max_row, right.max_row)
            max_column = max(left.max_column, right.max_column)
            for row in range(1, max_row + 1):
                for column in range(1, max_column + 1):
                    checked += 1
                    before = main_update.cell_signature(left.cell(row, column))
                    after = main_update.cell_signature(right.cell(row, column))
                    if before == after:
                        continue
                    if name == sheet_name and row > original_last_row:
                        allowed += 1
                        continue
                    raise RuntimeError(
                        f"Korunan alan degisti: {name}!{right.cell(row, column).coordinate}"
                    )
        if allowed == 0:
            raise RuntimeError("Yeni donem icin hicbir hucre degisikligi bulunamadi.")
        return checked, allowed
    finally:
        original.close()
        candidate.close()


def commit(candidate: Path, workbook: Path, bank: str, product: str) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_dir = ROOT / "outputs" / "yan_excel_guncelleme_yedek" / bank / product
    backup_dir.mkdir(parents=True, exist_ok=True)
    backup = backup_dir / f"{workbook.stem}_ONCESI_{stamp}{workbook.suffix}"
    shutil.copy2(workbook, backup)
    replacement = workbook.with_name(f".{workbook.name}.{bank}.{product}.tmp")
    shutil.copy2(candidate, replacement)
    os.replace(replacement, workbook)
    return backup


def run(args: argparse.Namespace) -> int:
    bank = args.bank
    product = args.product
    file_name, product_label = PRODUCTS[product]
    workbook = Path(args.workbook or (ROOT / file_name)).resolve()
    report_path = Path(
        args.audit_report or (ROOT / "outputs" / "yan_excel_guncelleme_audit" / f"{bank}_{product}_son_calisma.json")
    ).resolve()
    report = SideAuditReport(bank, product, str(workbook), bool(args.dry_run), datetime.now().isoformat(timespec="seconds"))
    try:
        if not workbook.is_file():
            raise FileNotFoundError(f"Yan Excel bulunamadi: {workbook}")
        spec = main_update.BANKS[bank]
        engine, sheet_name = source_engine(bank)
        history_dir = Path(args.history_dir or spec.history_dir).resolve()
        cache_dir = Path(args.cache_dir or (ROOT / spec.folder / "data" / "side_excel_cache")).resolve()

        if args.skip_download:
            log("KAP", "Indirme atlandi; lokal PDF history kullanilacak.")
        else:
            log("1-KAP", f"{spec.display_name} icin yeni raporlar kontrol ediliyor.")
            downloaded = main_update.download_reports(
                spec,
                history_dir,
                date.fromisoformat(args.today) if args.today else date.today(),
                args.force_download,
            )
            report.downloaded_quarters = downloaded
            log("1-KAP", "Tamamlandi.")

        pdf_files = engine.list_pdf_files(history_dir)
        if not pdf_files:
            raise FileNotFoundError(f"PDF history bulunamadi: {history_dir}")
        existing_periods = workbook_periods(workbook, sheet_name)
        selected_files = new_pdf_files(pdf_files, existing_periods, engine.extract_report_date)
        if args.period:
            if args.period not in main_update.complete_history_periods(history_dir):
                raise RuntimeError(f"Istenen donemin lokal PDF+Excel kaynagi eksik: {args.period}")
            selected_files = [path for path in selected_files if engine.extract_report_date(path) == args.period]

        if not selected_files:
            report.status = "NO_NEW_PERIOD"
            log("Sonuc", f"{product_label} icin yeni donem yok; dosya degistirilmedi.")
            write_report(report, report_path)
            return 0

        report.new_periods = [engine.extract_report_date(path) for path in selected_files]
        cache_dir.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix=f"{bank}_{product}_") as temp_dir:
            candidate = Path(temp_dir) / workbook.name
            shutil.copy2(workbook, candidate)
            log("2-YAN-EXCEL", f"{product_label} icin {len(selected_files)} yeni donem ayrisiyor.")
            payloads = engine.extract_history_payloads(selected_files, cache_dir, product, args.workers or getattr(engine, "DEFAULT_WORKERS", 1))
            writer = engine.write_stage_workbook if product == "stage" else engine.write_deposit_workbook
            added = writer(payloads, candidate)
            if added != len(report.new_periods):
                raise RuntimeError(f"Beklenen {len(report.new_periods)} donemin {added} adedi eklendi.")
            log("2-YAN-EXCEL", "Tamamlandi.")
            log("3-KONTROL", "Korunan alanlar denetleniyor.")
            report.checked_cells, report.allowed_changes = audit_candidate(workbook, candidate, sheet_name, report.new_periods)
            log("3-KONTROL", "Tamamlandi.")

            if args.dry_run:
                output = Path(args.dry_run_output or (ROOT / "outputs" / "yan_excel_dry_run" / f"{workbook.stem}_{bank}_DRY_RUN{workbook.suffix}")).resolve()
                output.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(candidate, output)
                report.output = str(output)
                log("Dry-run", f"Ana yan Excel degismedi; kontrol kopyasi: {output}")
            else:
                backup = commit(candidate, workbook, bank, product)
                report.output, report.backup = str(workbook), str(backup)
                log("Kayit", f"Atomik kayit tamamlandi; yedek: {backup}")
        report.status = "OK"
        write_report(report, report_path)
        log("Tamamlandi", f"Audit raporu: {report_path}")
        return 0
    except Exception as exc:
        report.status = "ERROR"
        report.error = f"{type(exc).__name__}: {exc}"
        write_report(report, report_path)
        log("HATA", f"Yan Excel degistirilmedi: {report.error}")
        return 1


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Yan Excel'lere dipnot akisi olmadan yalniz yeni donemi ekler.")
    parser.add_argument("--bank", choices=sorted(SIDE_MODULES), required=True)
    parser.add_argument("--product", choices=sorted(PRODUCTS), required=True)
    parser.add_argument("--workbook", help="Hedef yan Excel yolu.")
    parser.add_argument("--history-dir")
    parser.add_argument("--cache-dir")
    parser.add_argument("--workers", type=int)
    parser.add_argument("--today", help="KAP sorgusu icin test tarihi (YYYY-MM-DD).")
    parser.add_argument("--period", help="Yalniz bu lokal donemi isle (ornek: 30.06.2026).")
    parser.add_argument("--force-download", action="store_true")
    parser.add_argument("--skip-download", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--dry-run-output")
    parser.add_argument("--audit-report")
    return parser


if __name__ == "__main__":
    raise SystemExit(run(build_parser().parse_args()))
