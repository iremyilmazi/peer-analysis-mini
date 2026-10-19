from __future__ import annotations

import re
import sys
from io import BytesIO
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from teb_bankasi.peerpdf_scripts import teb_dipnot as _base

_ORIGINAL_EXTRACT_RAW_ROWS = _base.extract_raw_rows
_AMOUNT_WITH_PAREN_RE = re.compile(r"^(?:-|\(?\d+(?:[.,]\d{3})*(?:[.,]\d+)?\)?)$")
_ACTIVE_PDF_PATH: Path | None = None


CONSUMER_CREDIT_TITLE = (
    "Tüketici kredileri, bireysel kredi kartları, personel kredileri "
    "ve personel kredi kartlarına ilişkin bilgiler"
)
INSTALLMENT_COMMERCIAL_TITLE = "Taksitli ticari krediler ve kurumsal kredi kartlarına ilişkin bilgiler"
DEPOSIT_MATURITY_TITLE = "Mevduatın vade yapısına ilişkin bilgiler"
DEPOSIT_TOTAL_BY_MATURITY_TITLE = "Mevduatın vade yapısına ilişkin toplam bilgileri"
DEPOSIT_INTEREST_TITLE = "Mevduata Ödenen Faizin Vade Yapısına Göre Gösterimi"
RESTRUCTURED_NPL_TITLE = "Donuk alacaklar ve yeniden yapılandırılan kredilere ilişkin bilgiler"
NPL_MOVEMENT_TITLE = "Toplam donuk alacak hareketlerine ilişkin bilgiler"
BANKS_INTEREST_INCOME_TITLE = "Bankalardan alınan faiz gelirlerine ilişkin bilgiler"
SECURITIES_INTEREST_INCOME_TITLE = "Menkul değerlerden alınan faiz gelirlerine ilişkin bilgiler"
EQUITY_ITEMS_TITLE = "Özkaynak kalemlerine ilişkin bilgiler"
EXPECTED_LOSS_PROVISIONS_TITLE = "Bankaların kredi ve diğer alacaklara ilişkin değer düşüş karşılıkları:"
STAGE_PROVISION_TITLE = (
    "Standart nitelikli ve yakın izlemedeki krediler ile yeniden yapılandırılan "
    "yakın izlemedeki kredilere ilişkin bilgiler"
)
LIQUIDITY_COVERAGE_RATIO_TITLE = "LİKİDİTE KARŞILAMA ORANI"
STAGE3_PROVISIONS_TITLE = "Kredilere ilişkin olarak ayrılan özel karşılıklar:"
OTHER_OPERATING_INCOME_TITLE = "Diğer faaliyet gelirlerine ilişkin açıklamalar:"
CASH_AND_TCMB_TITLE = "Nakit değerler ve TCMB’ye ilişkin bilgiler"
USED_LOANS_INTEREST_EXPENSE_TITLE = "Kullanılan Kredilere Verilen Faizlere İlişkin Bilgiler"
LOAN_INTEREST_INCOME_TITLE = "Kredilerden alınan faiz gelirlerine ilişkin bilgiler"

_DEPOSIT_INTEREST_LABELS = [
    "Türk Parası - Bankalar Mevduatı", "Türk Parası - Tasarruf Mevduatı",
    "Türk Parası - Resmî Mevduat", "Türk Parası - Ticari Mevduat",
    "Türk Parası - Diğer Mevduat", "Türk Parası - 7 Gün İhbarlı Mevduat",
    "Türk Parası - Toplam", "Yabancı Para - Döviz Tevdiat Hesabı",
    "Yabancı Para - Bankalar Mevduatı", "Yabancı Para - 7 Gün İhbarlı Mevduat",
    "Yabancı Para - Kıymetli Maden D.Hs.", "Yabancı Para - Toplam", "Genel Toplam",
]


def _valid_deposit_interest_rows(rows: list[list[str]]) -> bool:
    return len(rows) == len(_DEPOSIT_INTEREST_LABELS) and all(
        len(row) == 9 and _base.normalize_key(row[0]) == _base.normalize_key(label)
        for row, label in zip(rows, _DEPOSIT_INTEREST_LABELS)
    )


def _deposit_interest_rows_from_pypdf(pdf_path: Path) -> list[list[str]]:
    from pypdf import PdfReader
    payload = pdf_path.read_bytes(); offset = payload.find(b"%PDF-")
    if offset < 0: return []
    labels = [label.split(" - ", 1)[-1] for label in _DEPOSIT_INTEREST_LABELS]
    for page in PdfReader(BytesIO(payload[offset:])).pages:
        text = page.extract_text() or ""
        if _base.normalize_key(DEPOSIT_INTEREST_TITLE) not in _base.normalize_key(text):
            continue
        section = ""; output: list[list[str]] = []; started = False
        for raw in text.splitlines():
            line = _base.normalize_text(raw); key = _base.normalize_key(line)
            if not started:
                if _base.normalize_key(DEPOSIT_INTEREST_TITLE) in key: started = True
                continue
            if re.fullmatch(r"\d{1,3}", line) and output and str(output[-1][-1]).endswith("."):
                output[-1][-1] += line
                continue
            if key == "turk parasi": section = "Türk Parası"; continue
            if key == "yabanci para": section = "Yabancı Para"; continue
            if not section: continue
            candidates = ["Genel Toplam"] if key.startswith("genel toplam") else labels
            for suffix in candidates:
                if key.startswith(_base.normalize_key(suffix)):
                    tail = line[len(suffix):].strip()
                    values = re.findall(r"(?<!\S)(?:-|\d[\d.]*)", tail)
                    if not values: break
                    label = "Genel Toplam" if suffix == "Genel Toplam" else f"{section} - {suffix}"
                    output.append([label, *values])
                    break
            if output and output[-1][0] == "Genel Toplam":
                return output if _valid_deposit_interest_rows(output) else []
    return []


def extract_deposit_interest_rows(page_texts: list[str]) -> list[list[str]]:
    try: rows = _base.extract_deposit_interest_rows(page_texts)
    except Exception: rows = []
    if _valid_deposit_interest_rows(rows): return rows
    pdf_path = getattr(_base, "_ACTIVE_PDF_PATH", None) or globals().get("_ACTIVE_PDF_PATH")
    if pdf_path:
        rows = _deposit_interest_rows_from_pypdf(Path(pdf_path))
    if not rows: raise RuntimeError(f"Hedef cari dönem satırları bulunamadı: {DEPOSIT_INTEREST_TITLE}")
    return rows


def _patch_base() -> None:
    for name in [
        "CONSUMER_CREDIT_TITLE",
        "INSTALLMENT_COMMERCIAL_TITLE",
        "DEPOSIT_MATURITY_TITLE",
        "DEPOSIT_TOTAL_BY_MATURITY_TITLE",
        "DEPOSIT_INTEREST_TITLE",
        "RESTRUCTURED_NPL_TITLE",
        "NPL_MOVEMENT_TITLE",
        "BANKS_INTEREST_INCOME_TITLE",
        "SECURITIES_INTEREST_INCOME_TITLE",
        "EQUITY_ITEMS_TITLE",
        "EXPECTED_LOSS_PROVISIONS_TITLE",
        "STAGE_PROVISION_TITLE",
        "LIQUIDITY_COVERAGE_RATIO_TITLE",
        "STAGE3_PROVISIONS_TITLE",
        "OTHER_OPERATING_INCOME_TITLE",
        "CASH_AND_TCMB_TITLE",
        "USED_LOANS_INTEREST_EXPENSE_TITLE",
        "LOAN_INTEREST_INCOME_TITLE",
    ]:
        setattr(_base, name, globals()[name])
    _base.STAGE_PROVISION_TARGET_LABELS = [
        "12 Aylık Beklenen Zarar Karşılığı",
        "Kredi Riskinde Önemli Artış",
    ]


def _is_title_line(line: str, title: str) -> bool:
    title_key = _base.normalize_key(title)
    line_key = _base.normalize_key(line)
    title_prefix = " ".join(title_key.split()[:4])
    return title_key in line_key or (title_prefix and title_prefix in line_key)


def _lines_after_title(page_texts: list[str], title: str, max_pages: int = 2) -> list[str]:
    lines: list[str] = []
    for start_page in _base.title_pages(page_texts, title):
        started = False
        for text in page_texts[start_page : min(len(page_texts), start_page + max_pages)]:
            for raw_line in text.splitlines():
                line = _base.normalize_text(raw_line)
                if not started:
                    if _is_title_line(line, title):
                        started = True
                    continue
                if line:
                    lines.append(line)
    return lines


def _parse_amount_line(line: str) -> list[str] | None:
    cleaned = _base.normalize_text(line)
    parts = cleaned.split()
    values: list[str] = []
    while parts and _base.AMOUNT_TOKEN_RE.match(parts[-1]):
        values.insert(0, parts.pop())
    if len(values) < 2 or not parts:
        return None
    return [" ".join(parts), *values]


def _current_text_rows(page_texts: list[str], title: str, max_pages: int = 2) -> list[list[str]]:
    for start_page in _base.title_pages(page_texts, title):
        rows: list[list[str]] = []
        seen_title = False
        in_current = False
        for text in page_texts[start_page : min(len(page_texts), start_page + max_pages)]:
            for raw_line in text.splitlines():
                line = _base.normalize_text(raw_line)
                key = _base.normalize_key(line)
                if not seen_title:
                    if _is_title_line(line, title):
                        seen_title = True
                    continue
                if "onceki donem" in key or "previous period" in key:
                    return rows
                if line.startswith("(*)") or line.startswith("(**)"):
                    if rows:
                        return rows
                    continue
                if "cari donem" in key or "current period" in key:
                    in_current = True
                    continue
                if not in_current:
                    continue
                parsed = _parse_amount_line(line)
                if parsed:
                    rows.append(parsed)
        if rows:
            return rows
    return []


_DEPOSIT_LABELS = [
    "Tasarruf Mevduatı", "Döviz Tevdiat Hesabı", "Yurtiçinde Yerleşik Kişiler",
    "Yurtdışında Yerleşik Kişiler", "Resmi Kuruluşlar Mevduatı",
    "Ticari Kuruluşlar Mevduatı", "Diğer Kuruluşlar Mevduatı",
    "Kıymetli Maden Depo Hesabı", "Bankalararası Mevduat", "T.C. Merkez Bankası",
    "Yurtiçi Bankalar", "Yurtdışı Bankalar", "Katılım Bankaları", "Diğer", "Toplam",
]
_DEPOSIT_TOP_LEVEL = [0, 1, 4, 5, 6, 7, 8, 13]


def _canonical_deposit_rows(rows: list[list[str]]) -> list[list[str]]:
    slots: list[list[int] | None] = [None] * len(_DEPOSIT_LABELS)
    originals: list[list[str] | None] = [None] * len(_DEPOSIT_LABELS)
    for row in rows:
        if len(row) < 3:
            continue
        key = _base.normalize_key(row[0])
        match = next((i for i, label in enumerate(_DEPOSIT_LABELS)
                      if key.startswith(_base.normalize_key(label))), None)
        if match is None or slots[match] is not None:
            continue
        vadesiz = _base.parse_amount_text(row[1]) or 0
        total = _base.parse_amount_text(row[-1]) or 0
        slots[match] = [vadesiz, total - vadesiz, total]
        originals[match] = [str(value) for value in row[1:]]
    if any(value is None for value in slots):
        return []
    total_vadesiz = sum(slots[i][0] for i in _DEPOSIT_TOP_LEVEL)
    total_amount = sum(slots[i][2] for i in _DEPOSIT_TOP_LEVEL)
    slots[-1] = [total_vadesiz, total_amount - total_vadesiz, total_amount]
    if any(min(value) < 0 or value[0] + value[1] != value[2] for value in slots):
        return []
    return [
        [label, *(original or [_base.format_amount(value) for value in values])]
        for label, values, original in zip(_DEPOSIT_LABELS, slots, originals)
    ]


def _deposit_current_rows_from_text(page_texts: list[str]) -> list[list[str]]:
    """Read the current-period HSBC maturity table from page text.

    In the 2026Q2 PDF the ruled-table extractor returns dashes for detail
    rows even though the text layer contains the amounts.  Restrict parsing
    to the Cari Dönem subsection so previous-period values cannot leak in.
    """
    title_key = _base.normalize_key(DEPOSIT_MATURITY_TITLE)
    for text in page_texts:
        if title_key not in _base.normalize_key(text):
            continue
        in_current = False
        rows: list[list[str]] = []
        for line in text.splitlines():
            key = _base.normalize_key(line)
            if "cari donem" in key:
                in_current = True
                continue
            if in_current and "onceki donem" in key:
                break
            if not in_current:
                continue
            parsed = _parse_amount_line(line)
            if parsed and any(key.startswith(_base.normalize_key(label)) for label in _DEPOSIT_LABELS):
                rows.append(parsed)
                if key.startswith("toplam"):
                    break
        if len(rows) == len(_DEPOSIT_LABELS):
            return rows
    return []


def _combine_wrapped_rows(lines: list[str], stop_keys: tuple[str, ...] = ()) -> list[list[str]]:
    rows: list[list[str]] = []
    pending_label = ""
    for line in lines:
        key = _base.normalize_key(line)
        if any(key.startswith(stop_key) for stop_key in stop_keys):
            break
        if (
            not line
            or "cari donem" in key
            or "onceki donem" in key
            or key.startswith("31 mart")
            or key in {"tp yp tp yp"}
        ):
            continue
        parsed = _parse_amount_line(line)
        if parsed:
            label = " ".join(part for part in [pending_label, parsed[0]] if part).strip()
            rows.append([label, *parsed[1:]])
            pending_label = ""
            if _base.normalize_key(label).startswith("toplam"):
                break
            continue
        pending_label = " ".join(part for part in [pending_label, line] if part).strip()
    return rows


def _plain_rows_after_title(
    page_texts: list[str],
    title: str,
    stop_keys: tuple[str, ...],
    max_pages: int = 1,
) -> list[list[str]]:
    rows: list[list[str]] = []
    for line in _lines_after_title(page_texts, title, max_pages=max_pages):
        key = _base.normalize_key(line)
        if rows and (line.startswith("(*)") or any(key.startswith(stop_key) for stop_key in stop_keys)):
            break
        if (
            not line
            or "cari donem" in key
            or key.startswith("31 mart")
            or key in {"kisa vadeli orta ve uzun vadeli toplam", "iii grup iv grup v grup"}
            or key in {"tahsil imkani tahsili supheli niteligindeki", "zarar", "sinirli krediler krediler krediler"}
        ):
            continue
        parsed = _parse_amount_line(line)
        if parsed:
            rows.append(parsed)
    return rows


def extract_consumer_credit_rows(page_texts: list[str]) -> list[list[str]]:
    """Extract the HSBC consumer-credit table even when PDF text splits title words."""
    rows: list[list[str]] = []
    started = False
    for text in page_texts:
        for raw_line in text.splitlines():
            line = _base.normalize_text(raw_line)
            compact = re.sub(r"[^a-z0-9çğıöşü]", "", _base.normalize_key(line))
            if not started:
                if (
                    "tuketicikredileri" in compact
                    and "bireyselkredikartlari" in compact
                    and "personelkredileri" in compact
                ):
                    started = True
                continue
            key = _base.normalize_key(line)
            if rows and (
                key.startswith("5 taksitli")
                or "taksitli ticari krediler ve kurumsal kredi kartlari" in key
            ):
                if len(rows) != 45:
                    raise RuntimeError(f"HSBC tüketici kredileri satır sayısı beklenenden farklı: {len(rows)}")
                return rows
            if (
                not line
                or "kisa vadeli" in key
                or key.startswith("31 mart")
                or key.startswith("30 haziran")
                or key.startswith("30 eylul")
                or key.startswith("31 aralik")
            ):
                continue
            parsed = _parse_amount_line(line)
            if parsed:
                rows.append(parsed[:4])
                if _base.normalize_key(parsed[0]).replace(" ", "").startswith("toplamtuketici"):
                    if len(rows) != 45:
                        raise RuntimeError(f"HSBC tüketici kredileri satır sayısı beklenenden farklı: {len(rows)}")
                    return rows
    return rows


def _fix_npl_date_label_rows(rows: list[list[str]], page_texts: list[str] | None = None) -> list[list[str]]:
    labels = [
        "Önceki Dönem Sonu Bakiyesi", "Dönem İçinde İntikal (+)",
        "Diğer Donuk Alacak Hesaplarından Giriş (+)", "Diğer Donuk Alacak Hesaplarına Çıkış (-)",
        "Dönem İçinde Tahsilat (-)", "Kayıttan Düşülen (-)", "Satılan (-)",
        "Kurumsal ve Ticari Krediler", "Bireysel Krediler", "Kredi Kartları", "Diğer",
        "Dönem Sonu Bakiyesi", "Karşılık (-)", "Bilançodaki Net Bakiyesi",
    ]
    if len(rows) != len(labels):
        return rows
    fixed = [[label, *(row[1:4])] for label, row in zip(labels, rows)]
    if page_texts:
        targets = {0: "onceki donem sonu bakiyesi", 11: "donem sonu bakiyesi"}
        for index, target in targets.items():
            for text in page_texts:
                for line in text.splitlines():
                    key = _base.normalize_key(line)
                    if not key.startswith(target) or ":" not in line:
                        continue
                    parsed = _parse_amount_line(line)
                    if parsed and len(parsed) >= 4:
                        fixed[index] = [labels[index], *parsed[-3:]]
                        break
                else:
                    continue
                break
    return fixed


def extract_deposit_total_by_maturity_rows(page_texts: list[str]) -> list[list[str]]:
    labels = ["Vadesiz", "7 Gün İhbarlı", "1 Aya Kadar", "1-3 Ay", "3-6 Ay", "6 Ay - 1 Yıl", "1 Yıl ve Üstü", "Birikimli Mevduat", "Toplam"]
    title_key = _base.normalize_key(DEPOSIT_MATURITY_TITLE)
    for text in page_texts:
        if title_key not in _base.normalize_key(text):
            continue
        in_current = False
        for line in text.splitlines():
            key = _base.normalize_key(line)
            if "cari donem" in key:
                in_current = True
                continue
            if in_current and key.startswith("toplam"):
                parsed = _parse_amount_line(line)
                if parsed and len(parsed) >= 10:
                    return [[label, "", "", value] for label, value in zip(labels, parsed[1:10])]
    return []


def _canonical_securities_interest_rows(rows: list[list[str]]) -> list[list[str]]:
    labels = [
        "Gerçeğe Uygun Değ. Farkı Kar/Zarar’a Yansıtılan Finansal Varlıklar",
        "Gerçeğe Uygun Değ. Farkı Diğer Kapsamlı Gelire Yansıtılan Finansal Varlıklar",
        "İtfa Edilmiş Maliyeti Üzerinden Değerlenen Finansal Varlıklar", "Toplam",
    ]
    if len(rows) != 4:
        return rows
    return [[label, *(row[1:3])] for label, row in zip(labels, rows)]


def _parse_current_previous_amount_line(line: str) -> list[str] | None:
    cleaned = _base.normalize_text(line)
    parts = cleaned.split()
    values: list[str] = []
    while parts and _AMOUNT_WITH_PAREN_RE.match(parts[-1]):
        values.insert(0, parts.pop())
    if not values or not parts:
        return None
    return [" ".join(parts), *values[:2]]


def _split_hsbc_expected_loss_line(line: str) -> list[str]:
    second_label = "Kredi Riskinde Önemli Artış (İkinci Aşama)"
    index = line.find(second_label)
    if index <= 0:
        return [line]
    return [line[:index].strip(), line[index:].strip()]


def extract_expected_loss_rows(page_texts: list[str]) -> list[list[str]]:
    rows: list[list[str]] = []
    pending_label = ""
    for line in _lines_after_title(page_texts, EXPECTED_LOSS_PROVISIONS_TITLE, max_pages=1):
        key = _base.normalize_key(line)
        if key.startswith("g diger faaliyet giderlerine"):
            break
        if not line or "cari donem" in key or "onceki donem" in key or key.startswith("31 mart"):
            continue
        for segment in _split_hsbc_expected_loss_line(line):
            parsed = _parse_current_previous_amount_line(segment)
            if parsed:
                label = " ".join(part for part in [pending_label, parsed[0]] if part).strip()
                rows.append([label, *parsed[1:]])
                pending_label = ""
                continue
            pending_label = " ".join(part for part in [pending_label, segment] if part).strip()
    labels = [
        "Beklenen Kredi Zarar Karşılıkları", "12 Aylık Beklenen Zarar Karşılığı (Birinci Aşama)",
        "Kredi Riskinde Önemli Artış (İkinci Aşama)", "Temerrüt (Üçüncü Aşama)",
        "Menkul Değerler Değer Düşüş Karşılıkları",
        "Gerçeğe Uygun Değer Farkı Kar veya Zarara Yansıtılan Finansal Varlıklar",
        "Gerçeğe Uygun Değer Farkı Diğer Kapsamlı Gelire Yansıtılan Varlıklar",
        "İştirakler, Bağlı Ortaklıklar ve Birlikte Kontrol Edilen Ortaklıklar Değer Düşüş Karşılıkları",
        "İştirakler", "Bağlı Ortaklıklar", "Birlikte Kontrol Edilen Ortaklıklar", "Diğer", "Toplam",
    ]
    clean = [row for row in rows if not re.match(r"^(?:30|31)\s", row[0])]
    if len(clean) == len(labels):
        return [[label, *(row[1:])] for label, row in zip(labels, clean)]
    return clean


def extract_raw_rows(
    page_texts: list[str],
    page_tables: list[list[list[list[str | None]]]],
    title: str,
    include_empty_amount_rows: bool = False,
) -> list[list[str]]:
    title_key = _base.normalize_key(title)
    if title_key == _base.normalize_key(CONSUMER_CREDIT_TITLE):
        rows = extract_consumer_credit_rows(page_texts)
        if rows:
            return rows
    if title_key == _base.normalize_key(INSTALLMENT_COMMERCIAL_TITLE):
        rows = _plain_rows_after_title(page_texts, title, ("6 kredilerin",), max_pages=1)
        if rows:
            return rows
    if title_key == _base.normalize_key(DEPOSIT_MATURITY_TITLE):
        rows = _deposit_current_rows_from_text(page_texts)
        if not rows:
            rows = _ORIGINAL_EXTRACT_RAW_ROWS(page_texts, page_tables, title, True)
        rows = _canonical_deposit_rows(rows)
        if rows:
            return rows
    if title_key == _base.normalize_key(NPL_MOVEMENT_TITLE):
        rows = _plain_rows_after_title(page_texts, title, ("10 iii",), max_pages=1)
        if rows:
            return _fix_npl_date_label_rows(rows, page_texts)
    if title_key == _base.normalize_key(SECURITIES_INTEREST_INCOME_TITLE):
        rows = _ORIGINAL_EXTRACT_RAW_ROWS(page_texts, page_tables, title, include_empty_amount_rows)
        return _canonical_securities_interest_rows(rows)
    if title_key == _base.normalize_key(CASH_AND_TCMB_TITLE):
        rows = _ORIGINAL_EXTRACT_RAW_ROWS(page_texts, page_tables, title, include_empty_amount_rows)
        labels = ["Kasa/Efektif", "TCMB", "Diğer (*)", "Toplam"]
        clean = [row for row in rows if _base.normalize_key(row[0]).startswith(("kasa efektif", "tcmb", "diger", "toplam"))]
        if len(clean) == 3:
            clean.insert(2, ["Diğer (*)", "-", "-"])
        if len(clean) == 4:
            return [[label, *(row[1:3])] for label, row in zip(labels, clean)]
        return clean
    if title_key == _base.normalize_key(STAGE3_PROVISIONS_TITLE):
        rows = _plain_rows_after_title(page_texts, title, ("10 donuk",), max_pages=1)
        if rows:
            return rows
    if title_key == _base.normalize_key(USED_LOANS_INTEREST_EXPENSE_TITLE):
        rows = _plain_rows_after_title(page_texts, title, ("2 istirakler",), max_pages=1)
        if rows:
            return rows
    if title_key == _base.normalize_key(EXPECTED_LOSS_PROVISIONS_TITLE):
        rows = extract_expected_loss_rows(page_texts)
        if rows:
            return rows
    if title_key in {
        _base.normalize_key(SECURITIES_INTEREST_INCOME_TITLE),
        _base.normalize_key(OTHER_OPERATING_INCOME_TITLE),
    }:
        rows = _combine_wrapped_rows(_lines_after_title(page_texts, title, max_pages=1))
        if rows:
            if title_key == _base.normalize_key(OTHER_OPERATING_INCOME_TITLE) and len(rows) == 5:
                labels = [
                    "Geçmiş Yıllar Giderlerine Ait Düzeltmeler (*)",
                    "Aktiflerin Satışından Elde Edilen Gelirler",
                    "Haberleşme Giderleri Karşılığı", "Diğer Gelirler", "Toplam",
                ]
                return [[label, *(row[1:])] for label, row in zip(labels, rows)]
            return rows
    return _ORIGINAL_EXTRACT_RAW_ROWS(
        page_texts,
        page_tables,
        title,
        include_empty_amount_rows=include_empty_amount_rows,
    )


def extract_restructured_npl_rows(page_texts: list[str]) -> list[list[str]]:
    rows = _current_text_rows(page_texts, RESTRUCTURED_NPL_TITLE, max_pages=1)
    targets = ["Karşılıklardan Önceki Brüt Tutarlar", "Yeniden Yapılandırılan Krediler"]
    found: dict[str, list[str]] = {}
    for row in rows:
        key = _base.normalize_key(row[0])
        for label in targets:
            if key.startswith(_base.normalize_key(label)):
                found[label] = [label, *(row[1:4])]
    if len(found) == 2:
        return [found[label] for label in targets]
    for start_page in _base.title_pages(page_texts, RESTRUCTURED_NPL_TITLE):
        text = page_texts[start_page]
        if "Bulunmamaktadır" in text:
            return [[label, "-", "-", "-"] for label in targets]
    return [found.get(label, [label, "-", "-", "-"]) for label in targets]


def extract_stage_provision_rows(page_texts: list[str]) -> list[list[str]]:
    targets = [
        ("12 aylik beklenen zarar karsiligi", "12 Aylik Beklenen Zarar Karsiligi"),
        ("kredi riskinde onemli artis", "Kredi Riskinde Onemli Artis"),
    ]
    found: dict[str, list[str]] = {}
    for start_page in _base.title_pages(page_texts, STAGE_PROVISION_TITLE):
        for text in page_texts[start_page : min(len(page_texts), start_page + 2)]:
            for raw_line in text.splitlines():
                parsed = _parse_amount_line(raw_line)
                if not parsed:
                    continue
                label_key = _base.normalize_key(parsed[0])
                for target_key, label in targets:
                    if target_key in label_key and target_key not in found:
                        found[target_key] = [label, parsed[1], parsed[2] if len(parsed) > 2 else ""]
            if len(found) == len(targets):
                return [found[target_key] for target_key, _ in targets]
    return [found[target_key] for target_key, _ in targets if target_key in found]


def extract_liquidity_coverage_rows(page_texts: list[str]) -> list[list[str]]:
    for start_page in _base.title_pages(page_texts, LIQUIDITY_COVERAGE_RATIO_TITLE):
        for text in page_texts[start_page : min(len(page_texts), start_page + 4)]:
            for raw_line in text.splitlines():
                line = _base.normalize_text(raw_line)
                if "likidite karsilama orani" not in _base.normalize_key(line):
                    continue
                values = re.findall(r"%\s*([0-9]+(?:[.,][0-9]+)?)", line)
                if len(values) >= 2:
                    return [
                        [
                            "Likidite Karsilama Orani (%)",
                            values[0].replace(",", "."),
                            values[1].replace(",", "."),
                        ]
                    ]
                parsed = _parse_amount_line(line)
                if parsed and len(parsed) >= 3:
                    return [["Likidite Karsilama Orani (%)", parsed[1], parsed[2]]]
    return []


def extract_equity_rows(page_texts: list[str]) -> list[list[str]]:
    targets = {
        "cekirdek sermaye toplami": "Çekirdek Sermaye Toplamı",
        "katki sermaye toplami": "Katkı Sermaye Toplamı",
        "toplam ozkaynak ana sermaye ve katki sermaye toplami": "Toplam Özkaynak",
        "toplam risk agirlikli tutarlar": "Toplam Risk Ağırlıklı Tutarlar",
    }
    found: dict[str, str] = {}
    for text in page_texts:
        for raw_line in text.splitlines():
            parsed = _parse_amount_line(raw_line)
            if not parsed:
                continue
            key = _base.normalize_key(parsed[0])
            for target, label in targets.items():
                if key.startswith(target):
                    # Later occurrences are the regulatory final totals and
                    # supersede the pre-deduction subtotal on the prior page.
                    found[target] = parsed[1]
    missing = [label for target, label in targets.items() if target not in found]
    if missing:
        raise RuntimeError("Hedef özkaynak kalemleri bulunamadı: " + ", ".join(missing))
    return [[label, "", "", found[target]] for target, label in targets.items()]


def _patch_hsbc_overrides() -> None:
    _base.extract_raw_rows = extract_raw_rows
    _base.extract_restructured_npl_rows = extract_restructured_npl_rows
    _base.extract_stage_provision_rows = extract_stage_provision_rows
    _base.extract_liquidity_coverage_rows = extract_liquidity_coverage_rows
    _base.extract_equity_rows = extract_equity_rows
    _base.extract_deposit_total_by_maturity_rows = extract_deposit_total_by_maturity_rows


_patch_base()
_patch_hsbc_overrides()

for _name in dir(_base):
    if not _name.startswith("__") and _name not in globals():
        globals()[_name] = getattr(_base, _name)


def main() -> None:
    _patch_base()
    _patch_hsbc_overrides()
    _base.main()


if __name__ == "__main__":
    main()
