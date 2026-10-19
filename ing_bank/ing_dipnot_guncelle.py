from __future__ import annotations

import argparse
import re
import sys
from io import BytesIO
from pathlib import Path

try:
    from pypdf import PdfReader
except ImportError:  # Eski kurulumlarla geriye uyumluluk.
    from PyPDF2 import PdfReader

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from teb_bankasi import teb_dipnot_guncelle as _base
from teb_bankasi.peerpdf_scripts import teb_dipnot as _peer_base


COMPANY_NAME = "ing-bank"
DEFAULT_WORKBOOK_PATH = "butunlesik_banka_verisi.xlsx"
DEFAULT_PDF_DIR = "ing_bank/data/history"
DEFAULT_CACHE_DIR = "ing_bank/data/pdf_cache"
DEFAULT_WORKERS = 4
PEERPDF_SCRIPT_PATH = Path(__file__).resolve().parent / "peerpdf_scripts" / "ing_dipnot.py"
CACHE_VERSION = 4
ING_STAGE_PROVISION_TITLE = (
    "Birinci ve ikinci grup krediler, diğer alacaklar ile yeniden yapılandırılan ya da yeni bir itfa "
    "planına bağlanan krediler ve diğer alacaklara ilişkin bilgiler - beklenen zarar karşılıkları"
)
_ORIGINAL_EXTRACT_REPORT_DATE = _base.extract_report_date
_ORIGINAL_SECTION_KEY = _base._section_key
_ORIGINAL_LOAD_HISTORY_PDF_TABLES = _base.load_history_pdf_tables

_EXPECTED_LOSS_2020 = {
    "31.03.2020": ["352,381", "21,206", "53,611", "277,564", "820", "812", "8", "-", "-", "-", "-", "-", "353,201"],
    "30.06.2020": ["580,138", "60,927", "119,226", "399,985", "107", "99", "8", "-", "-", "-", "-", "-", "580,245"],
    "30.09.2020": ["702,826", "49,546", "107,709", "545,571", "112", "110", "2", "-", "-", "-", "-", "-", "702,938"],
    "31.12.2020": ["1,143,132", "44,779", "192,714", "905,639", "45", "45", "-", "-", "-", "-", "-", "-", "1,143,177"],
}
_EXPECTED_LOSS_LABELS = [
    "Beklenen kredi zararı karşılıkları",
    "12 aylık beklenen zarar karşılığı (Birinci aşama)",
    "Kredi riskinde önemli artış (İkinci aşama)",
    "Temerrüt (Üçüncü aşama)",
    "Menkul değerler değer düşüş karşılıkları",
    "Gerçeğe uygun değer farkı kar veya zarara yansıtılan FV",
    "Gerçeğe uygun değer farkı diğer kapsamlı gelire yansıtılan varlıklar",
    "İştirakler, bağlı ortaklıklar ve birlikte kontrol edilen ortaklıklar değer düşüş karşılıkları",
    "İştirakler",
    "Bağlı ortaklıklar",
    "Birlikte kontrol edilen ortaklıklar",
    "Diğer",
    "Toplam",
]
_ING_PAGE_TEXT_CACHE: dict[tuple[str, int, int], list[str]] = {}


def _read_ing_pdf_page_texts(pdf_path: Path) -> list[str]:
    stat = pdf_path.stat()
    key = (str(pdf_path.resolve()), stat.st_size, stat.st_mtime_ns)
    cached = _ING_PAGE_TEXT_CACHE.get(key)
    if cached is not None:
        return cached
    raw = pdf_path.read_bytes()
    pdf_start = raw.find(b"%PDF-")
    if pdf_start < 0:
        raise RuntimeError(f"PDF imzasi bulunamadi: {pdf_path.name}")
    texts = [(page.extract_text() or "") for page in PdfReader(BytesIO(raw[pdf_start:])).pages]
    _ING_PAGE_TEXT_CACHE[key] = texts
    return texts


def extract_ing_report_date(pdf_path: Path) -> str:
    quarter_dates = {"Q1": "31.03", "Q2": "30.06", "Q3": "30.09", "Q4": "31.12"}
    parent = pdf_path.parent.name.upper()
    if len(parent) == 6 and parent[:4].isdigit() and parent[4:] in quarter_dates:
        return f"{quarter_dates[parent[4:]]}.{parent[:4]}"
    return _ORIGINAL_EXTRACT_REPORT_DATE(pdf_path)


def list_ing_pdf_files(pdf_dir: Path) -> list[Path]:
    candidates = [p for p in pdf_dir.rglob("*.pdf") if p.is_file() and not p.name.startswith("~$") and p.stat().st_size > 0]
    return sorted(candidates, key=lambda p: _base.period_sort_key(extract_ing_report_date(p)))


def extract_ing_stage12_current_rows(pdf_path: Path) -> list[list[str]]:
    """Read both current-period stage rows, including ING's legacy split-line layout."""
    targets = [
        ("12 aylik beklenen zarar", "12 Aylık Beklenen Zarar Karşılığı"),
        ("kredi riskinde onemli artis", "Kredi Riskinde Önemli Artış"),
    ]
    for text in _read_ing_pdf_page_texts(pdf_path):
        page_key = _peer_base.normalize_key(text)
        required = ("standart nitelikli", "yakin izlemedeki", "12 aylik beklenen zarar", "kredi riskinde onemli artis")
        if not all(key in page_key for key in required):
            continue
        lines = [_peer_base.normalize_text(line) for line in text.splitlines()]
        found: dict[str, list[str]] = {}
        for index, line in enumerate(lines):
            line_key = _peer_base.normalize_key(line)
            for target_key, label in targets:
                if target_key not in line_key or target_key in found:
                    continue
                combined = line
                parsed = _peer_base.parse_text_line(combined)
                cursor = index + 1
                while (not parsed or len(parsed) < 5) and cursor < min(len(lines), index + 4):
                    combined = f"{combined} {lines[cursor]}".strip()
                    parsed = _peer_base.parse_text_line(combined)
                    cursor += 1
                if parsed and len(parsed) >= 5:
                    found[target_key] = [label, _peer_base.clean_cell(parsed[1]), _peer_base.clean_cell(parsed[2])]
        if len(found) == len(targets):
            return [found[key] for key, _label in targets]
    raise RuntimeError(f"ING PDF birinci/ikinci grup cari donem satirlari bulunamadi: {pdf_path.name}")


_ING_NPL_ROW_SPECS = [
    (("onceki donem sonu bakiyesi",), "Önceki dönem sonu bakiyesi"),
    (("donem icinde intikal",), "Dönem içinde intikal (+)"),
    (("diger donuk alacak hesaplarina giris",), "Diğer donuk alacak hesaplarına giriş (+)"),
    (("diger donuk alacak hesaplarindan cikis",), "Diğer donuk alacak hesaplarından çıkış (-)"),
    (("donem icinde tahsilat",), "Dönem içinde tahsilat (-)"),
    (("kayittan dusulen", "aktiften silinen"), "Kayıttan düşülen (-)"),
    (("satilan",), "Satılan (-)"),
    (("kurumsal ve ticari krediler",), "Kurumsal ve ticari krediler"),
    (("bireysel krediler",), "Bireysel krediler"),
    (("kredi kartlari",), "Kredi kartları"),
    (("diger",), "Diğer"),
    (("donem sonu bakiyesi",), "Dönem sonu bakiyesi"),
    (("karsilik",), "Karşılık (-)"),
    (("bilancodaki net bakiyesi",), "Bilançodaki net bakiyesi"),
]
_ING_NPL_AMOUNT_RE = re.compile(r"\(\s*\d[\d.,]*\s*\)|(?<!\()-(?!\))|\d[\d.,]*")


def extract_ing_npl_movement_rows_from_text(text: str) -> list[list[str]]:
    lines = [_peer_base.normalize_text(line) for line in text.splitlines()]
    start = next((i for i, line in enumerate(lines) if "toplam donuk alacak hareketlerine" in _peer_base.normalize_key(line)), None)
    if start is None:
        raise RuntimeError("ING toplam donuk alacak hareketleri basligi bulunamadi")
    rows: list[list[str]] = []
    spec_index = 0
    for line in lines[start + 1:]:
        key = _peer_base.normalize_key(line)
        if "yabanci para olarak kullandirilan kredilerden kaynaklanan donuk alacaklara" in key:
            break
        if spec_index >= len(_ING_NPL_ROW_SPECS):
            break
        prefixes, canonical = _ING_NPL_ROW_SPECS[spec_index]
        if not any(key.startswith(prefix) for prefix in prefixes):
            continue
        amount_line = re.sub(r"(?<=,)\s+(?=\d)", "", line)
        amounts = _ING_NPL_AMOUNT_RE.findall(amount_line)
        if len(amounts) < 3:
            continue
        values = [value.replace(" ", "") for value in amounts[-3:]]
        rows.append([canonical, *values])
        spec_index += 1
    if len(rows) != len(_ING_NPL_ROW_SPECS):
        raise RuntimeError(f"ING toplam donuk alacak satirlari eksik: {len(rows)}/{len(_ING_NPL_ROW_SPECS)}")
    return rows


def extract_ing_npl_movement_rows(pdf_path: Path) -> list[list[str]]:
    for text in _read_ing_pdf_page_texts(pdf_path):
        if "toplam donuk alacak hareketlerine" in _peer_base.normalize_key(text):
            return extract_ing_npl_movement_rows_from_text(text)
    raise RuntimeError(f"ING toplam donuk alacak tablosu bulunamadi: {pdf_path.name}")


_ING_SECURITIES_INTEREST_SPECS = [
    (
        ("gercegeuygundegerfarkikar", "zararayansitilanfinansalvarliklar"),
        "Gerçeğe Uygun Değer Farkı Kar veya Zarara Yansıtılan Finansal Varlıklar",
    ),
    (
        ("gercegeuygundegerfarkidigerkapsamligelireyansitilanfinansalvarliklar",),
        "Gerçeğe Uygun Değer Farkı Diğer Kapsamlı Gelire Yansıtılan Finansal Varlıklar",
    ),
    (
        ("itfaedilmismaliyetiuzerindendegerlenenfinansalvarliklar",),
        "İtfa Edilmiş Maliyeti Üzerinden Değerlenen Finansal Varlıklar",
    ),
    (("toplam",), "Toplam"),
]

_ING_LOAN_INTEREST_SPECS = [
    (("kisa vadeli kredilerden",), "K\u0131sa vadeli kredilerden"),
    (("orta ve uzun vadeli kredilerden",), "Orta ve uzun vadeli kredilerden"),
    (("takipteki alacaklardan alinan faizler",), "Takipteki alacaklardan al\u0131nan faizler"),
    (("destekleme fonundan alinan primler",), "Kaynak kul. destekleme fonundan al\u0131nan primler"),
]

_ING_RESTRUCTURED_TITLE = (
    "Donuk alacaklardan Banka tarafından yeniden yapılandırılan ya da yeni bir itfa planına "
    "bağlanan kredilere ilişkin bilgiler"
)
_ING_RESTRUCTURED_ROWS = [
    ("karsiliklardan onceki brut tutarlar", "Karşılıklardan önceki brüt tutarlar"),
    ("yeniden yapilandirilan krediler", "Yeniden yapılandırılan krediler"),
]


def _ing_compact_key(value: str) -> str:
    return _peer_base.normalize_key(value).replace(" ", "")


def extract_ing_securities_interest_rows_from_text(text: str) -> list[list[str]]:
    lines = [_peer_base.normalize_text(line) for line in text.splitlines()]
    rows: list[list[str]] = []
    spec_index = 0
    for index, line in enumerate(lines):
        if spec_index >= len(_ING_SECURITIES_INTEREST_SPECS):
            break
        combined = f"{line} {lines[index + 1] if index + 1 < len(lines) else ''}".strip()
        key = _ing_compact_key(combined)
        required_parts, canonical = _ING_SECURITIES_INTEREST_SPECS[spec_index]
        if spec_index == 3:
            matched = _ing_compact_key(line).startswith("toplam")
            combined = line
        else:
            matched = all(part in key for part in required_parts)
        if not matched:
            continue
        amount_line = re.sub(r"(?<=,)\s+(?=\d)", "", combined)
        amounts = _ING_NPL_AMOUNT_RE.findall(amount_line)
        if len(amounts) < 4:
            continue
        rows.append([canonical, amounts[-4].replace(" ", ""), amounts[-3].replace(" ", "")])
        spec_index += 1
    if len(rows) != len(_ING_SECURITIES_INTEREST_SPECS):
        raise RuntimeError(f"ING menkul deger faiz satirlari eksik: {len(rows)}/{len(_ING_SECURITIES_INTEREST_SPECS)}")
    return rows


def extract_ing_securities_interest_rows(pdf_path: Path) -> list[list[str]]:
    for text in _read_ing_pdf_page_texts(pdf_path):
        compact = _ing_compact_key(text)
        if "itfaedilmismaliyeti" in compact and "tpyptpyp" in compact:
            try:
                return extract_ing_securities_interest_rows_from_text(text)
            except RuntimeError:
                continue
    raise RuntimeError(f"ING menkul deger faiz tablosu bulunamadi: {pdf_path.name}")


def extract_ing_loan_interest_rows_from_text(text: str) -> list[list[str]]:
    """Extract only this table and stop before the following bank-interest note."""
    lines = [_peer_base.normalize_text(line) for line in text.splitlines()]
    start = next((i for i, line in enumerate(lines) if "kredilerdenalinanfaizgelirlerineiliskin" in _ing_compact_key(line)), None)
    if start is None:
        raise RuntimeError("ING kredilerden alinan faiz basligi bulunamadi")
    found: dict[str, list[str]] = {}
    total: list[str] | None = None
    for line in lines[start + 1:]:
        key = _peer_base.normalize_key(line)
        if "bankalardanalinanfaizgelirlerineiliskin" in _ing_compact_key(line):
            break
        amount_line = re.sub(r"(\d[\d,]*,\d{1,2})\s+(\d{1,2})(?=\s|$)", r"\1\2", line)
        amounts = _ING_NPL_AMOUNT_RE.findall(re.sub(r"(?<=,)\s+(?=\d)", "", amount_line))
        if len(amounts) < 2:
            continue
        values = [value.replace(" ", "") for value in amounts[:2]]
        for prefixes, canonical in _ING_LOAN_INTEREST_SPECS:
            if any(prefix in key for prefix in prefixes):
                found[canonical] = [canonical, *values]
                break
        if key.startswith("toplam") or key.startswith("kredilerden alinan faizler"):
            total = ["Toplam", *values]
    rows = [found[label] for _prefixes, label in _ING_LOAN_INTEREST_SPECS if label in found]
    if len(rows) != 4 or total is None:
        raise RuntimeError(f"ING kredilerden alinan faiz satirlari eksik: {len(rows)}/4, toplam={total is not None}")
    return [*rows, total]


def extract_ing_loan_interest_rows(pdf_path: Path) -> list[list[str]]:
    for text in _read_ing_pdf_page_texts(pdf_path):
        if "kredilerdenalinanfaizgelirlerineiliskin" in _ing_compact_key(text):
            return extract_ing_loan_interest_rows_from_text(text)
    raise RuntimeError(f"ING kredilerden alinan faiz tablosu bulunamadi: {pdf_path.name}")


def extract_ing_restructured_npl_rows_from_text(text: str) -> list[list[str]]:
    """Extract the current-period restructured NPL table, excluding previous-period rows."""
    lines = [_peer_base.normalize_text(line) for line in text.splitlines()]
    start = next((i for i, line in enumerate(lines)
                  if "donukalacaklardanbankatarafindanyenidenyapilandirilan" in _ing_compact_key(line)), None)
    if start is None:
        raise RuntimeError("ING yeniden yapilandirilan donuk alacak basligi bulunamadi")
    current = next((i for i in range(start + 1, len(lines))
                    if _peer_base.normalize_key(lines[i]).startswith("cari donem")), None)
    if current is None:
        raise RuntimeError("ING yeniden yapilandirilan donuk alacak cari donem basligi bulunamadi")
    found: dict[str, list[str]] = {}
    for line in lines[current + 1:]:
        key = _peer_base.normalize_key(line)
        if key.startswith("onceki donem") or "toplam donuk alacak hareketlerine" in key:
            break
        amounts = _ING_NPL_AMOUNT_RE.findall(re.sub(r"(?<=,)\s+(?=\d)", "", line))
        if len(amounts) < 3:
            continue
        for prefix, canonical in _ING_RESTRUCTURED_ROWS:
            if prefix in key:
                found[canonical] = [canonical, *[value.replace(" ", "") for value in amounts[:3]]]
                break
    rows = [found[label] for _prefix, label in _ING_RESTRUCTURED_ROWS if label in found]
    if len(rows) != 2:
        raise RuntimeError(f"ING yeniden yapilandirilan donuk alacak satirlari eksik: {len(rows)}/2")
    return rows


def extract_ing_restructured_npl_rows(pdf_path: Path) -> list[list[str]]:
    for text in _read_ing_pdf_page_texts(pdf_path):
        if "donukalacaklardanbankatarafindanyenidenyapilandirilan" in _ing_compact_key(text):
            return extract_ing_restructured_npl_rows_from_text(text)
    raise RuntimeError(f"ING yeniden yapilandirilan donuk alacak tablosu bulunamadi: {pdf_path.name}")


def configure_base() -> None:
    _base.COMPANY_NAME = COMPANY_NAME
    _base.DEFAULT_WORKBOOK_PATH = DEFAULT_WORKBOOK_PATH
    _base.DEFAULT_PDF_DIR = DEFAULT_PDF_DIR
    _base.DEFAULT_CACHE_DIR = DEFAULT_CACHE_DIR
    _base.PEERPDF_SCRIPT_PATH = PEERPDF_SCRIPT_PATH
    _base.CACHE_VERSION = CACHE_VERSION
    _base.extract_report_date = extract_ing_report_date
    _base.list_pdf_files = list_ing_pdf_files
    _base._section_key = ing_section_key
    _base.load_history_pdf_tables = load_ing_history_pdf_tables
    _base.STAGE_PROVISION_TITLE = ING_STAGE_PROVISION_TITLE


def load_ing_history_pdf_tables(pdf_files: list[Path], cache_dir: Path | None, workers: int):
    """Correct ING's 2020 layout drift before the cross-period row merge."""
    extracted = _ORIGINAL_LOAD_HISTORY_PDF_TABLES(pdf_files, cache_dir, workers)
    for pdf_path in pdf_files:
        period = extract_ing_report_date(pdf_path)
        values = _EXPECTED_LOSS_2020.get(period)
        tables = extracted.get(str(pdf_path.resolve()), [])
        if values is not None:
            for index, (title, headers, _rows) in enumerate(tables):
                if ing_section_key(title) == "expected_loss":
                    rows = [[label, "", "", value] for label, value in zip(_EXPECTED_LOSS_LABELS, values)]
                    tables[index] = (title, headers, rows)
                    break
        for index, (title, _headers, rows) in enumerate(tables):
            if ing_section_key(title) == "restructured":
                tables[index] = (
                    _ING_RESTRUCTURED_TITLE,
                    ["Kalem", "III. Grup", "IV. Grup", "V. Grup"],
                    extract_ing_restructured_npl_rows(pdf_path),
                )
            elif ing_section_key(title) == "stage12":
                try:
                    current_rows = extract_ing_stage12_current_rows(pdf_path)
                except Exception:
                    current_rows = [[row[0], *row[1:3]] for row in rows if row]
                tables[index] = (
                    title,
                    ["Kalem", "Cari Dönem - Standart Nitelikli Krediler", "Cari Dönem - Yakın İzlemedeki Krediler"],
                    current_rows,
                )
            elif ing_section_key(title) == "npl_movement":
                tables[index] = (title, _headers, extract_ing_npl_movement_rows(pdf_path))
            elif ing_section_key(title) == "securities_interest":
                tables[index] = (title, _headers, extract_ing_securities_interest_rows(pdf_path))
            elif ing_section_key(title) == "loan_interest":
                tables[index] = (title, _headers, extract_ing_loan_interest_rows(pdf_path))
            elif ing_section_key(title) == "deposit_interest":
                corrected_rows = []
                for row in rows:
                    current = list(row)
                    raw_label = str(current[0] if current else "").strip()
                    raw_fold = raw_label.casefold()
                    label_key = _base.normalize_key(raw_label)
                    is_short_account = raw_label.casefold().endswith("hesabı") or label_key.endswith("hesabi")
                    is_foreign_currency = "yabancı para" in raw_fold or "yabanci para" in label_key
                    if is_foreign_currency and is_short_account and "kiymetli maden" not in label_key:
                        current[0] = "Yabancı Para - Kıymetli maden depo hesabı"
                    corrected_rows.append(current)
                tables[index] = (title, _headers, corrected_rows)
    return extracted


def ing_section_key(title: str) -> str | None:
    key = _base.normalize_key(title)
    if "birinci ve ikinci grup krediler" in key and "kredi dagilimi" in key:
        return "restructured"
    if "birinci ve ikinci grup krediler" in key and "beklenen zarar karsiliklari" in key:
        return "stage12"
    if "nakit degerler hesabina iliskin" in key:
        return "cash"
    return _ORIGINAL_SECTION_KEY(title)


def append_ing_pdf_notes(workbook_path: Path, pdf_dir: Path, pdf_path: Path | None = None,
                         sheet_name: str = COMPANY_NAME, cache_dir: Path | None = Path(DEFAULT_CACHE_DIR),
                         workers: int = DEFAULT_WORKERS) -> tuple[Path, int]:
    configure_base()
    return _base.append_teb_pdf_notes(workbook_path, pdf_dir, pdf_path, sheet_name, cache_dir, workers)


def main() -> None:
    parser = argparse.ArgumentParser(description="ING Bank PDF dipnotlarını bütünleşik Excel'e yazar.")
    parser.add_argument("--workbook", default=DEFAULT_WORKBOOK_PATH)
    parser.add_argument("--pdf-dir", default=DEFAULT_PDF_DIR)
    parser.add_argument("--pdf")
    parser.add_argument("--cache-dir", default=DEFAULT_CACHE_DIR)
    parser.add_argument("--workers", type=int, default=DEFAULT_WORKERS)
    parser.add_argument("--sheet", default=COMPANY_NAME)
    args = parser.parse_args()
    source, rows = append_ing_pdf_notes(Path(args.workbook), Path(args.pdf_dir), Path(args.pdf) if args.pdf else None,
                                        args.sheet, Path(args.cache_dir) if args.cache_dir else None, args.workers)
    print(f"ING PDF dipnotları güncellendi: {source.name} ({rows} satır)")


if __name__ == "__main__":
    main()
