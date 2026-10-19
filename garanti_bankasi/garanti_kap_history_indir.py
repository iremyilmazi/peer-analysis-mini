from __future__ import annotations

import argparse
import http.client
import json
import re
import socket
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from datetime import date
from email.message import Message
from pathlib import Path


BASE_URL = "https://www.kap.org.tr/tr"
GARANTI_MKK_MEMBER_OID = "4028e4a140f2ed720140f37cb2a601b7"
DEFAULT_OUTPUT_DIR = Path(__file__).resolve().parent / "data" / "history"
DEFAULT_START_YEAR = 2020
REQUEST_PAUSE_SECONDS = 0.35
RETRY_WAIT_SECONDS = (10, 30, 90, 180, 300)
QUARTER_RE = re.compile(r"^(?P<year>\d{4})Q(?P<quarter>[1-4])$")
# Hafif guncelleme paketinde, mevcut ana Excel'in son donemi bu degisken
# ile aktarilir. Normal/tam paket davranisinda deger None'dur.
INCREMENTAL_AFTER_QUARTER: str | None = None


@dataclass(frozen=True)
class Disclosure:
    disclosure_index: int
    publish_date: str
    year: int
    quarter: int
    ft_niteligi: str
    attachments: tuple[dict, ...]

    @property
    def quarter_key(self) -> str:
        return f"{self.year}Q{self.quarter}"


def request_json(url: str, data: dict | None = None) -> object:
    payload = None
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json, text/plain, */*",
    }
    if data is not None:
        payload = json.dumps(data).encode("utf-8")
        headers["Content-Type"] = "application/json"

    request = urllib.request.Request(url, data=payload, headers=headers)
    for attempt in range(4):
        try:
            with urllib.request.urlopen(request, timeout=60) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            if exc.code != 429 or attempt == 3:
                raise
            wait_seconds = RETRY_WAIT_SECONDS[min(attempt, len(RETRY_WAIT_SECONDS) - 1)]
            print(f"WARN: KAP 429 verdi, {wait_seconds} sn bekleniyor: {url}")
            time.sleep(wait_seconds)

    raise RuntimeError(f"KAP istegi tamamlanamadi: {url}")


def by_criteria_payload(from_date: date, to_date: date) -> dict:
    return {
        "fromDate": from_date.isoformat(),
        "toDate": to_date.isoformat(),
        "mkkMemberOidList": [GARANTI_MKK_MEMBER_OID],
        "inactiveMkkMemberOidList": [],
        "disclosureClass": "FR",
        "subjectList": [],
        "isLate": "",
        "mainSector": "",
        "sector": "",
        "market": "",
        "index": "",
        "ruleType": "",
        "period": "",
        "year": "",
        "fromSrc": False,
        "srcCategory": "",
        "disclosureIndexList": [],
    }


def fetch_year_rows(year: int, today: date) -> list[dict]:
    from_date = date(year, 1, 1)
    to_date = min(date(year, 12, 31), today)
    if from_date > today:
        return []

    payload = by_criteria_payload(from_date, to_date)
    url = f"{BASE_URL}/api/disclosure/members/byCriteria"
    rows = request_json(url, payload)
    if not isinstance(rows, list):
        raise RuntimeError(f"KAP byCriteria beklenmeyen cevap verdi: {type(rows).__name__}")
    return rows


def fetch_attachment_detail(disclosure_index: int) -> dict:
    url = f"{BASE_URL}/api/notification/attachment-detail/{disclosure_index}"
    data = request_json(url)
    if not isinstance(data, list) or not data:
        raise RuntimeError(f"attachment-detail bos geldi: {disclosure_index}")
    if not isinstance(data[0], dict):
        raise RuntimeError(f"attachment-detail beklenmeyen cevap verdi: {disclosure_index}")
    return data[0]


def iter_candidate_disclosures(start_year: int, today: date) -> list[Disclosure]:
    candidates: list[Disclosure] = []
    for year in range(start_year, today.year + 1):
        rows = fetch_year_rows(year, today)
        print(f"KAP liste: {year} -> {len(rows)} satir")
        for row in rows:
            if row.get("subject") != "Finansal Rapor":
                continue
            try:
                row_year = int(row.get("year"))
                quarter = int(row.get("period"))
                disclosure_index = int(row.get("disclosureIndex"))
            except (TypeError, ValueError):
                continue
            if row_year < start_year or quarter not in {1, 2, 3, 4}:
                continue

            try:
                detail = fetch_attachment_detail(disclosure_index)
            except Exception as exc:
                print(f"WARN: {row_year}Q{quarter} attachment-detail alinamadi: {disclosure_index} -> {type(exc).__name__}: {exc}")
                continue
            disclosure = detail.get("disclosure", {})
            disclosure_detail = disclosure.get("disclosureDetail", {})
            ft_niteligi = str(disclosure_detail.get("ftNiteligi") or "")
            attachments = tuple(detail.get("attachments") or ())
            candidates.append(
                Disclosure(
                    disclosure_index=disclosure_index,
                    publish_date=str(row.get("publishDate") or ""),
                    year=row_year,
                    quarter=quarter,
                    ft_niteligi=ft_niteligi,
                    attachments=attachments,
                )
            )
            time.sleep(REQUEST_PAUSE_SECONDS)
    return candidates


def choose_nc_by_quarter(candidates: list[Disclosure]) -> dict[str, Disclosure]:
    selected: dict[str, Disclosure] = {}
    by_quarter: dict[str, list[Disclosure]] = {}
    for item in candidates:
        by_quarter.setdefault(item.quarter_key, []).append(item)

    for quarter_key, items in sorted(by_quarter.items()):
        nc_items = [item for item in items if item.ft_niteligi.casefold() == "nc"]
        if not nc_items:
            print(f"WARN: {quarter_key} icin konsolide olmayan Finansal Rapor bulunamadi; adaylar: {[i.disclosure_index for i in items]}")
            continue
        selected[quarter_key] = max(nc_items, key=lambda item: item.disclosure_index)
    return selected


def period_end_date(year: int, quarter: int) -> str:
    if quarter == 1:
        return f"31.03.{year}"
    if quarter == 2:
        return f"30.06.{year}"
    if quarter == 3:
        return f"30.09.{year}"
    if quarter == 4:
        return f"31.12.{year}"
    raise ValueError(f"Gecersiz quarter: {quarter}")


def parse_quarter_key(quarter_key: str) -> tuple[int, int] | None:
    match = QUARTER_RE.match(quarter_key)
    if not match:
        return None
    return int(match.group("year")), int(match.group("quarter"))


def quarter_sort_key(quarter_key: str) -> tuple[int, int]:
    parsed = parse_quarter_key(quarter_key)
    if parsed is None:
        return 9999, 4
    return parsed


def next_quarter(quarter_key: str) -> str:
    year, quarter = quarter_sort_key(quarter_key)
    if quarter == 4:
        return f"{year + 1}Q1"
    return f"{year}Q{quarter + 1}"


def quarter_is_after(left: str, right: str) -> bool:
    return quarter_sort_key(left) > quarter_sort_key(right)


def quarter_has_required_files(quarter_dir: Path) -> bool:
    if not quarter_dir.exists():
        return False
    has_excel = any(path.is_file() and path.suffix.lower() in {".xls", ".xlsx", ".xlsm"} for path in quarter_dir.glob("*"))
    has_pdf = any(path.is_file() and path.suffix.lower() == ".pdf" for path in quarter_dir.glob("*"))
    return has_excel and has_pdf


def existing_complete_quarters(output_dir: Path) -> list[str]:
    if not output_dir.exists():
        return []
    quarters: list[str] = []
    for child in output_dir.iterdir():
        if not child.is_dir():
            continue
        if parse_quarter_key(child.name) is None:
            continue
        if quarter_has_required_files(child):
            quarters.append(child.name)
    return sorted(quarters, key=quarter_sort_key)


def latest_complete_quarter(output_dir: Path) -> str | None:
    quarters = existing_complete_quarters(output_dir)
    return quarters[-1] if quarters else None


def sanitize_filename(filename: str) -> str:
    filename = re.sub(r'[<>:"/\\|?*]+', "_", filename).strip()
    return filename or "kap_download"


def filename_from_headers(headers: Message, fallback: str) -> str:
    content_disposition = headers.get("Content-Disposition") or ""
    encoded = re.search(r"filename\*=UTF-8''([^;]+)", content_disposition, flags=re.IGNORECASE)
    if encoded:
        return urllib.parse.unquote(encoded.group(1))
    plain = re.search(r'filename="?([^";]+)"?', content_disposition, flags=re.IGNORECASE)
    if plain:
        return plain.group(1)
    return fallback


def download_url(url: str, target_dir: Path, fallback_name: str) -> Path:
    request = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    for attempt in range(6):
        try:
            with urllib.request.urlopen(request, timeout=300) as response:
                data = response.read()
                filename = filename_from_headers(response.headers, fallback_name)
            break
        except urllib.error.HTTPError as exc:
            if exc.code != 429 or attempt == 5:
                raise
            wait_seconds = RETRY_WAIT_SECONDS[min(attempt, len(RETRY_WAIT_SECONDS) - 1)]
            print(f"WARN: KAP 429 verdi, {wait_seconds} sn bekleniyor: {url}")
            time.sleep(wait_seconds)
        except (http.client.IncompleteRead, TimeoutError, socket.timeout) as exc:
            if attempt == 5:
                raise
            wait_seconds = RETRY_WAIT_SECONDS[min(attempt, len(RETRY_WAIT_SECONDS) - 1)]
            print(f"WARN: Indirme yarim kaldi, {wait_seconds} sn sonra tekrar denenecek: {url} -> {type(exc).__name__}: {exc}")
            time.sleep(wait_seconds)
    else:
        raise RuntimeError(f"Indirme tamamlanamadi: {url}")

    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / sanitize_filename(filename)
    if target.exists() and target.stat().st_size == len(data):
        print(f"Mevcut: {target}")
        return target
    target.write_bytes(data)
    print(f"Indirildi: {target}")
    return target


def download_disclosure(disclosure: Disclosure, output_dir: Path, force: bool = False) -> tuple[list[Path], list[str]]:
    quarter_dir = output_dir / disclosure.quarter_key
    downloaded: list[Path] = []
    warnings: list[str] = []

    excel_url = f"{BASE_URL}/api/notification/export/excel/{disclosure.disclosure_index}"
    excel_name = f"{period_end_date(disclosure.year, disclosure.quarter).replace('.', '')}_Bildirimler.xls"
    if not force and quarter_dir.exists() and any(quarter_dir.glob("*.xls*")):
        print(f"Mevcut Excel atlandi: {quarter_dir}")
    else:
        try:
            downloaded.append(download_url(excel_url, quarter_dir, excel_name))
        except Exception as exc:
            warning = f"WARN: {disclosure.quarter_key} Excel indirilemedi: {disclosure.disclosure_index} -> {type(exc).__name__}: {exc}"
            print(warning)
            warnings.append(warning)

    pdf_attachments = [
        item for item in disclosure.attachments if str(item.get("fileExtension") or "").casefold() == "pdf"
    ]
    if not pdf_attachments:
        warning = f"WARN: {disclosure.quarter_key} PDF eki bulunamadi: {disclosure.disclosure_index}"
        print(warning)
        warnings.append(warning)
        return downloaded, warnings

    if not force and quarter_dir.exists() and any(quarter_dir.glob("*.pdf")):
        print(f"Mevcut PDF atlandi: {quarter_dir}")
        return downloaded, warnings

    for attachment in pdf_attachments:
        obj_id = str(attachment.get("objId") or "")
        if not obj_id:
            warning = f"WARN: {disclosure.quarter_key} PDF objId bos: {disclosure.disclosure_index}"
            print(warning)
            warnings.append(warning)
            continue
        file_name = str(attachment.get("fileName") or f"{disclosure.quarter_key}.pdf")
        pdf_url = f"{BASE_URL}/api/file/download/{obj_id}"
        try:
            downloaded.append(download_url(pdf_url, quarter_dir, file_name))
        except Exception as exc:
            warning = f"WARN: {disclosure.quarter_key} PDF indirilemedi: {file_name} -> {type(exc).__name__}: {exc}"
            print(warning)
            warnings.append(warning)
    return downloaded, warnings


def run(start_year: int, output_dir: Path, today: date, force: bool = False) -> None:
    candidates = iter_candidate_disclosures(start_year, today)
    selected = choose_nc_by_quarter(candidates)
    print(f"Secilen konsolide olmayan rapor sayisi: {len(selected)}")

    expected_quarters = [
        f"{year}Q{quarter}"
        for year in range(start_year, today.year + 1)
        for quarter in range(1, 5)
        if date(year, ((quarter - 1) * 3) + 1, 1) <= today
    ]
    for quarter_key in expected_quarters:
        if quarter_key not in selected:
            print(f"WARN: {quarter_key} icin secilebilir rapor bulunamadi.")

    total_warnings = 0
    for quarter_key, disclosure in sorted(selected.items()):
        print(f"\n{quarter_key} indiriliyor: disclosureIndex={disclosure.disclosure_index}")
        _, warnings = download_disclosure(disclosure, output_dir, force=force)
        total_warnings += len(warnings)
        time.sleep(REQUEST_PAUSE_SECONDS)

    print(f"\nGaranti history indirme tamamlandi: {output_dir.resolve()}")
    print(f"Indirilen/secili donem: {len(selected)}; uyari: {total_warnings}")


def run_incremental(output_dir: Path, today: date, force: bool = False) -> list[str]:
    latest_quarter = latest_complete_quarter(output_dir)
    # Excel esas kaynaktir; yarim kalan indirme/prova hedefi atlatmaz.
    if INCREMENTAL_AFTER_QUARTER:
        latest_quarter = INCREMENTAL_AFTER_QUARTER
    if latest_quarter:
        start_year = quarter_sort_key(latest_quarter)[0]
        print(f"Mevcut son eksiksiz Garanti donemi: {latest_quarter}")
    else:
        start_year = DEFAULT_START_YEAR
        print("Mevcut eksiksiz Garanti history bulunamadi; ilk kurulum araligi taranacak.")

    candidates = iter_candidate_disclosures(start_year, today)
    selected = choose_nc_by_quarter(candidates)
    if latest_quarter:
        expected_quarter = next_quarter(latest_quarter)
        if expected_quarter not in selected:
            print(
                f"Yeni donem indirilmiyor: ana Excel/history son donemi {latest_quarter}; "
                f"yalniz beklenen sonraki donem {expected_quarter}."
            )
            return []
        selected = {expected_quarter: selected[expected_quarter]}

    if not selected:
        print("Yeni Garanti KAP donemi bulunamadi; indirme atlandi.")
        return []

    downloaded_quarters: list[str] = []
    total_warnings = 0
    for quarter_key, disclosure in sorted(selected.items(), key=lambda item: quarter_sort_key(item[0])):
        print(f"\nYeni donem indiriliyor: {quarter_key} disclosureIndex={disclosure.disclosure_index}")
        _, warnings = download_disclosure(disclosure, output_dir, force=force)
        total_warnings += len(warnings)
        if quarter_has_required_files(output_dir / quarter_key):
            downloaded_quarters.append(quarter_key)
        else:
            print(f"WARN: {quarter_key} klasoru eksik Excel/PDF nedeniyle guncellemeye hazir degil.")
        time.sleep(REQUEST_PAUSE_SECONDS)

    print(f"\nGaranti yeni donem indirme tamamlandi: {output_dir.resolve()}")
    print(f"Yeni indirilen/hazir donem: {len(downloaded_quarters)}; uyari: {total_warnings}")
    return downloaded_quarters


def main() -> None:
    parser = argparse.ArgumentParser(description="Garanti KAP finansal rapor history dosyalarini ceyrek klasorlerine indirir.")
    parser.add_argument("--start-year", type=int, default=DEFAULT_START_YEAR, help="Baslangic yili.")
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR), help="History ana klasoru.")
    parser.add_argument("--today", help="Test icin bugun tarihi, YYYY-MM-DD.")
    parser.add_argument("--force", action="store_true", help="Mevcut Excel/PDF dosyalari olsa bile yeniden indir.")
    parser.add_argument("--all-history", action="store_true", help="Tum history araligini yeniden tara; varsayilan sadece yeni donemleri indirir.")
    args = parser.parse_args()

    today = date.fromisoformat(args.today) if args.today else date.today()
    if args.all_history:
        run(args.start_year, Path(args.output_dir), today, force=args.force)
    else:
        run_incremental(Path(args.output_dir), today, force=args.force)


if __name__ == "__main__":
    main()
