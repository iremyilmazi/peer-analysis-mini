from __future__ import annotations

import re
import sys
from collections import defaultdict
from pathlib import Path

import pdfplumber

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from teb_bankasi.peerpdf_scripts import teb_dipnot as _base

_ORIGINAL_EXTRACT_RAW_ROWS = _base.extract_raw_rows
_ORIGINAL_EXTRACT_ALL_TABLES = _base.extract_all_tables
_ORIGINAL_EXTRACT_OTHER_EXPENSE_WORDS = (
    _base.extract_other_operating_expense_rows_from_words
)
_ACTIVE_PDF_PATH: Path | None = None


CONSUMER_CREDIT_TITLE = (
    "Tüketici kredileri, bireysel kredi kartları, personel kredileri "
    "ve personel kredi kartlarına ilişkin bilgiler"
)
INSTALLMENT_COMMERCIAL_TITLE = "Taksitli ticari krediler ve kurumsal kredi kartlarına ilişkin bilgiler"
DEPOSIT_MATURITY_TITLE = "Mevduatın vade yapısına ilişkin bilgiler"
DEPOSIT_TOTAL_BY_MATURITY_TITLE = "Mevduatın vade yapısına ilişkin toplam bilgileri"
DEPOSIT_INTEREST_TITLE = "Mevduata ödenen faizin vade yapısına göre gösterimi"
RESTRUCTURED_NPL_TITLE = "Donuk alacaklardan Banka’ca yeniden yapılandırılan kredilere ilişkin bilgiler"
NPL_MOVEMENT_TITLE = "Toplam donuk alacak hareketlerine ilişkin bilgiler"
BANKS_INTEREST_INCOME_TITLE = "Bankalardan alınan faiz gelirlerine ilişkin bilgiler"
SECURITIES_INTEREST_INCOME_TITLE = "Menkul değerlerden alınan faizlere ilişkin bilgiler"
EQUITY_ITEMS_TITLE = "Özkaynak kalemlerine ilişkin bilgiler"
EXPECTED_LOSS_PROVISIONS_TITLE = "Beklenen zarar karşılıkları ve diğer karşılık giderleri"
STAGE_PROVISION_TITLE = (
    "Standart nitelikli ve yakın izlemedeki krediler ile yeniden yapılandırılan "
    "yakın izlemedeki kredilere ilişkin bilgiler"
)
STAGE_PROVISION_HEADERS = ["Kalem", "Standart Nitelikli Krediler", "Yakın İzlemedeki Krediler"]
LIQUIDITY_COVERAGE_RATIO_TITLE = "LİKİDİTE KARŞILAMA ORANI"
STAGE3_PROVISIONS_TITLE = "Kredilere ilişkin olarak ayrılan temerrüt (Üçüncü Aşama) karşılıkları"
OTHER_OPERATING_INCOME_TITLE = "Diğer faaliyet gelirlerine ilişkin açıklamalar"
OTHER_OPERATING_EXPENSE_TITLE = "Diğer faaliyet giderlerine ilişkin bilgiler"
CASH_AND_TCMB_TITLE = "Nakit değerler ve TCMB’ye ilişkin bilgiler"
USED_LOANS_INTEREST_EXPENSE_TITLE = "Kullanılan Kredilere Verilen Faizlere İlişkin Bilgiler"
LOAN_INTEREST_INCOME_TITLE = "Kredilerden alınan faiz gelirlerine ilişkin bilgiler"


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
        "OTHER_OPERATING_EXPENSE_TITLE",
        "CASH_AND_TCMB_TITLE",
        "USED_LOANS_INTEREST_EXPENSE_TITLE",
        "LOAN_INTEREST_INCOME_TITLE",
    ]:
        setattr(_base, name, globals()[name])
    _base.STAGE_PROVISION_HEADERS = STAGE_PROVISION_HEADERS
    _base.STAGE_PROVISION_TARGET_LABELS = [
        "12 aylık beklenen zarar karşılığı",
        "Kredi riskinde önemli artış",
    ]


def _is_title_line(line: str, title: str) -> bool:
    title_key = _base.normalize_key(title)
    line_key = _base.normalize_key(line)
    title_prefix = " ".join(title_key.split()[:4])
    return title_key in line_key or (title_prefix and title_prefix in line_key)


def _lines_after_title(page_texts: list[str], title: str, max_pages: int = 2, include_title_tail: bool = False) -> list[str]:
    for start_page in _base.title_pages(page_texts, title):
        lines: list[str] = []
        started = False
        for text in page_texts[start_page : min(len(page_texts), start_page + max_pages)]:
            for raw_line in text.splitlines():
                line = _base.normalize_text(raw_line)
                if not started:
                    if _is_title_line(line, title):
                        started = True
                        if include_title_tail:
                            lines.append(line)
                    continue
                if line:
                    lines.append(line)
        if lines:
            return lines
    return []


def _parse_amount_line(line: str) -> list[str] | None:
    parts = _base.normalize_text(line).split()
    values: list[str] = []
    while parts and _base.SIGNED_AMOUNT_TOKEN_RE.match(parts[-1]):
        if re.fullmatch(r"\(\d+\)", parts[-1]):
            break
        values.insert(0, parts.pop())
    if len(values) < 2 or not parts:
        return None
    return [" ".join(parts), *values]


def _parse_variable_amount_line(line: str, min_values: int = 1) -> list[str] | None:
    parts = _base.normalize_text(line).split()
    values: list[str] = []
    while parts and _base.SIGNED_AMOUNT_TOKEN_RE.match(parts[-1]):
        if re.fullmatch(r"\(\d+\)", parts[-1]):
            break
        values.insert(0, parts.pop())
    if len(values) < min_values or not parts:
        return None
    return [" ".join(parts), *values]


def _is_amount_only_line(line: str) -> bool:
    parts = _base.normalize_text(line).split()
    return bool(parts) and all(_base.SIGNED_AMOUNT_TOKEN_RE.match(part) for part in parts)


def _current_pair(values: list[str]) -> list[str]:
    current = values[:2]
    while len(current) < 2:
        current.append("")
    return current


def _current_rows(rows: list[list[str]], value_count: int = 2) -> list[list[str]]:
    output: list[list[str]] = []
    for row in rows:
        values = row[1 : 1 + value_count]
        while len(values) < value_count:
            values.append("")
        output.append([row[0], *values])
    return output


def to_interest_row(row: list[str]) -> list[str]:
    output = [_base.normalize_text(value) for value in row[:3]]
    while len(output) < 3:
        output.append("")
    return output


def _clean_label(label: str, title: str | None = None) -> str:
    cleaned = _base.normalize_text(label)
    while ":" in cleaned:
        prefix, tail = cleaned.split(":", 1)
        prefix_key = _base.normalize_key(prefix)
        if (title is not None and _base.normalize_key(title) in prefix_key) or "iliskin bilgiler" in prefix_key:
            cleaned = tail.strip()
            continue
        break
    cleaned = re.sub(r"\s*:?\s*31\s+(?:Mart|Aralık)\s+20\d{2}\s*$", "", cleaned).strip()
    if title is not None and _base.normalize_key(title) == _base.normalize_key(INSTALLMENT_COMMERCIAL_TITLE):
        cleaned = re.sub(r"\s+\(", "(", cleaned)
    if title is not None and _base.normalize_key(title) == _base.normalize_key(DEPOSIT_MATURITY_TITLE):
        key = _base.normalize_key(cleaned)
        if key == "yurt disinda yer k":
            cleaned = "Yurt Dışında Yer. K."
        elif key == "resmi kur mevduati":
            cleaned = "Resmî Kur. Mevduatı"
    return cleaned


def _combine_wrapped_rows(
    lines: list[str],
    stop_keys: tuple[str, ...] = (),
    skip_period_markers: bool = True,
) -> list[list[str]]:
    rows: list[list[str]] = []
    pending_label = ""
    for line in lines:
        key = _base.normalize_key(line)
        if rows and any(key.startswith(stop_key) for stop_key in stop_keys):
            break
        if (
            not line
            or (skip_period_markers and "cari donem" in key)
            or (skip_period_markers and "onceki donem" in key)
            or key.startswith("31 mart")
            or key in {"tp yp tp yp", "tp yp"}
            or line.startswith("(")
            or line.startswith("Yetkili ")
        ):
            continue
        parsed = _parse_amount_line(line)
        if parsed:
            label = _clean_label(" ".join(part for part in [pending_label, parsed[0]] if part).strip())
            rows.append([label, *parsed[1:]])
            pending_label = ""
            if _base.normalize_key(label).startswith(("toplam", "genel toplam")):
                break
            continue
        pending_label = " ".join(part for part in [pending_label, line] if part).strip()
    return rows


def _plain_current_rows(page_texts: list[str], title: str, stop_keys: tuple[str, ...] = (), max_pages: int = 1) -> list[list[str]]:
    rows: list[list[str]] = []
    in_current = True
    for line in _lines_after_title(page_texts, title, max_pages=max_pages, include_title_tail=True):
        key = _base.normalize_key(line)
        if rows and any(key.startswith(stop_key) for stop_key in stop_keys):
            break
        if "onceki donem" in key:
            if rows:
                break
            in_current = False
            continue
        if "cari donem" in key:
            in_current = True
            continue
        if not in_current:
            continue
        if key in {"kisa vadeli orta ve uzun vadeli toplam", "iii grup iv grup v grup"}:
            continue
        parsed = _parse_amount_line(line)
        if parsed:
            rows.append([_clean_label(parsed[0], title), *parsed[1:]])
    return rows


def _credit_values_without_accrual(values: list[str]) -> list[str]:
    if len(values) >= 4:
        return [values[0], values[1], values[-1]]
    if len(values) == 3:
        return values
    if len(values) == 2:
        return [values[0], "", values[1]]
    if len(values) == 1:
        return [values[0], "", ""]
    return ["", "", ""]


def _credit_display_labels(title: str) -> list[str]:
    """Return the row captions exactly as printed in the İş Bankası tables.

    Repeated child captions must stay repeated.  Prefixing them with their
    parent (for example ``Tüketici Kredileri-TP - Konut Kredisi``) creates a
    label that does not exist in the source PDF.
    """
    if _base.normalize_key(title) == _base.normalize_key(CONSUMER_CREDIT_TITLE):
        return [
            "Tüketici Kredileri-TP",
            "Konut Kredisi",
            "Taşıt Kredisi",
            "İhtiyaç Kredisi",
            "Diğer",
            "Tüketici Kredileri-Dövize Endeksli",
            "Konut Kredisi",
            "Taşıt Kredisi",
            "İhtiyaç Kredisi",
            "Diğer",
            "Tüketici Kredileri-YP",
            "Konut Kredisi",
            "Taşıt Kredisi",
            "İhtiyaç Kredisi",
            "Diğer",
            "Bireysel Kredi Kartları-TP",
            "Taksitli",
            "Taksitsiz",
            "Bireysel Kredi Kartları-YP",
            "Taksitli",
            "Taksitsiz",
            "Personel Kredileri-TP",
            "Konut Kredisi",
            "Taşıt Kredisi",
            "İhtiyaç Kredisi",
            "Diğer",
            "Personel Kredileri-Dövize Endeksli",
            "Konut Kredisi",
            "Taşıt Kredisi",
            "İhtiyaç Kredisi",
            "Diğer",
            "Personel Kredileri-YP",
            "Konut Kredisi",
            "Taşıt Kredisi",
            "İhtiyaç Kredisi",
            "Diğer",
            "Personel Kredi Kartları-TP",
            "Taksitli",
            "Taksitsiz",
            "Personel Kredi Kartları-YP",
            "Taksitli",
            "Taksitsiz",
            "Kredili Mevduat Hesabı-TP (Gerçek Kişi)",
            "Kredili Mevduat Hesabı-YP (Gerçek Kişi)",
            "Toplam",
        ]
    return [
        "Taksitli Ticari Krediler-TP",
        "İşyeri Kredileri",
        "Taşıt Kredileri",
        "İhtiyaç Kredileri",
        "Diğer",
        "Taksitli Ticari Krediler-Dövize Endeksli",
        "İşyeri Kredileri",
        "Taşıt Kredileri",
        "İhtiyaç Kredileri",
        "Diğer",
        "Taksitli Ticari Krediler-YP",
        "İşyeri Kredileri",
        "Taşıt Kredileri",
        "İhtiyaç Kredileri",
        "Diğer",
        "Kurumsal Kredi Kartları-TP",
        "Taksitli",
        "Taksitsiz",
        "Kurumsal Kredi Kartları-YP",
        "Taksitli",
        "Taksitsiz",
        "Kredili Mevduat Hesabı-TP (Tüzel Kişi)",
        "Kredili Mevduat Hesabı-YP (Tüzel Kişi)",
        "Toplam",
    ]


def _credit_current_rows(page_texts: list[str], title: str, stop_keys: tuple[str, ...], max_pages: int = 1) -> list[list[str]]:
    consumer_schema = [
        "Tüketici Kredileri-TP",
        "Tüketici Kredileri-TP - Konut Kredisi",
        "Tüketici Kredileri-TP - Taşıt Kredisi",
        "Tüketici Kredileri-TP - İhtiyaç Kredisi",
        "Tüketici Kredileri-TP - Diğer",
        "Tüketici Kredileri-Dövize Endeksli",
        "Tüketici Kredileri-Dövize Endeksli - Konut Kredisi",
        "Tüketici Kredileri-Dövize Endeksli - Taşıt Kredisi",
        "Tüketici Kredileri-Dövize Endeksli - İhtiyaç Kredisi",
        "Tüketici Kredileri-Dövize Endeksli - Diğer",
        "Tüketici Kredileri-YP",
        "Tüketici Kredileri-YP - Konut Kredisi",
        "Tüketici Kredileri-YP - Taşıt Kredisi",
        "Tüketici Kredileri-YP - İhtiyaç Kredisi",
        "Tüketici Kredileri-YP - Diğer",
        "Bireysel Kredi Kartları-TP",
        "Bireysel Kredi Kartları-TP - Taksitli",
        "Bireysel Kredi Kartları-TP - Taksitsiz",
        "Bireysel Kredi Kartları-YP",
        "Bireysel Kredi Kartları-YP - Taksitli",
        "Bireysel Kredi Kartları-YP - Taksitsiz",
        "Personel Kredileri-TP",
        "Personel Kredileri-TP - Konut Kredisi",
        "Personel Kredileri-TP - Taşıt Kredisi",
        "Personel Kredileri-TP - İhtiyaç Kredisi",
        "Personel Kredileri-TP - Diğer",
        "Personel Kredileri-Dövize Endeksli",
        "Personel Kredileri-Dövize Endeksli - Konut Kredisi",
        "Personel Kredileri-Dövize Endeksli - Taşıt Kredisi",
        "Personel Kredileri-Dövize Endeksli - İhtiyaç Kredisi",
        "Personel Kredileri-Dövize Endeksli - Diğer",
        "Personel Kredileri-YP",
        "Personel Kredileri-YP - Konut Kredisi",
        "Personel Kredileri-YP - Taşıt Kredisi",
        "Personel Kredileri-YP - İhtiyaç Kredisi",
        "Personel Kredileri-YP - Diğer",
        "Personel Kredi Kartları-TP",
        "Personel Kredi Kartları-TP - Taksitli",
        "Personel Kredi Kartları-TP - Taksitsiz",
        "Personel Kredi Kartları-YP",
        "Personel Kredi Kartları-YP - Taksitli",
        "Personel Kredi Kartları-YP - Taksitsiz",
        "Kredili Mevduat Hesabı-TP (Gerçek Kişi)",
        "Kredili Mevduat Hesabı-YP (Gerçek Kişi)",
        "Toplam",
    ]
    installment_schema = [
        "Taksitli Ticari Krediler-TP",
        "Taksitli Ticari Krediler-TP - İşyeri Kredileri",
        "Taksitli Ticari Krediler-TP - Taşıt Kredileri",
        "Taksitli Ticari Krediler-TP - İhtiyaç Kredileri",
        "Taksitli Ticari Krediler-TP - Diğer",
        "Taksitli Ticari Krediler-Dövize Endeksli",
        "Taksitli Ticari Krediler-Dövize Endeksli - İşyeri Kredileri",
        "Taksitli Ticari Krediler-Dövize Endeksli - Taşıt Kredileri",
        "Taksitli Ticari Krediler-Dövize Endeksli - İhtiyaç Kredileri",
        "Taksitli Ticari Krediler-Dövize Endeksli - Diğer",
        "Taksitli Ticari Krediler-YP",
        "Taksitli Ticari Krediler-YP - İşyeri Kredileri",
        "Taksitli Ticari Krediler-YP - Taşıt Kredileri",
        "Taksitli Ticari Krediler-YP - İhtiyaç Kredileri",
        "Taksitli Ticari Krediler-YP - Diğer",
        "Kurumsal Kredi Kartları-TP",
        "Kurumsal Kredi Kartları-TP - Taksitli",
        "Kurumsal Kredi Kartları-TP - Taksitsiz",
        "Kurumsal Kredi Kartları-YP",
        "Kurumsal Kredi Kartları-YP - Taksitli",
        "Kurumsal Kredi Kartları-YP - Taksitsiz",
        "Kredili Mevduat Hesabı-TP (Tüzel Kişi)",
        "Kredili Mevduat Hesabı-YP (Tüzel Kişi)",
        "Toplam",
    ]
    schema = (
        consumer_schema
        if _base.normalize_key(title) == _base.normalize_key(CONSUMER_CREDIT_TITLE)
        else installment_schema
    )
    parent_by_key = {
        _base.normalize_key(label): label
        for label in schema
        if " - " not in label
        and not label.startswith("Kredili Mevduat")
        and label != "Toplam"
    }
    schema_by_key = {_base.normalize_key(label): label for label in schema}
    child_names = {
        "konut kredisi": "Konut Kredisi",
        "tasit kredisi": "Taşıt Kredisi",
        "ihtiyac kredisi": "İhtiyaç Kredisi",
        "isyeri kredileri": "İşyeri Kredileri",
        "tasit kredileri": "Taşıt Kredileri",
        "ihtiyac kredileri": "İhtiyaç Kredileri",
        "taksitli": "Taksitli",
        "taksitsiz": "Taksitsiz",
        "diger": "Diğer",
    }
    values_by_label: dict[str, list[str]] = {}
    current_parent = ""

    for line in _lines_after_title(page_texts, title, max_pages=max_pages, include_title_tail=False):
        key = _base.normalize_key(line)
        if values_by_label and any(key.startswith(stop_key) for stop_key in stop_keys):
            break
        if (
            not line
            or "cari donem" in key
            or "onceki donem" in key
            or key.startswith("31 mart")
            or key in {
                "faiz ve gelir tahakkuk",
                "kisa vadeli orta ve uzun vadeli toplam",
                "ve reeskontlari",
            }
            or line.startswith("(")
            or line.startswith("Yetkili ")
        ):
            continue
        parsed = _parse_amount_line(line)
        raw_label = _clean_label(parsed[0] if parsed else line, title)
        label_key = _base.normalize_key(raw_label)
        values = _credit_values_without_accrual(parsed[1:] if parsed else [])

        if label_key in parent_by_key:
            current_parent = parent_by_key[label_key]
            values_by_label[current_parent] = values
            continue
        if label_key in child_names and current_parent:
            contextual_label = f"{current_parent} - {child_names[label_key]}"
            if contextual_label in schema:
                values_by_label[contextual_label] = values
            continue
        direct_label = schema_by_key.get(label_key)
        if direct_label is not None:
            values_by_label[direct_label] = values
            if direct_label == "Toplam":
                break

    return [[label, *values_by_label.get(label, ["", "", ""])] for label in schema]


def _credit_rows_by_position(title: str, stop_keys: tuple[str, ...]) -> list[list[str]]:
    """Preserve blank İş Bankası credit columns using PDF x coordinates."""
    if _ACTIVE_PDF_PATH is None:
        return []
    title_prefix = " ".join(_base.normalize_key(title).split()[:5])
    with pdfplumber.open(str(_ACTIVE_PDF_PATH)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            if title_prefix not in _base.normalize_key(text):
                continue
            words = page.extract_words(x_tolerance=1, y_tolerance=2)
            width = float(page.width)
            grouped: dict[float, list[dict[str, object]]] = defaultdict(list)
            for word in words:
                grouped[round(float(word["top"]), 1)].append(word)
            rows: list[list[str]] = []
            started = False
            pending_label = ""
            width = float(page.width)
            for top in sorted(grouped):
                line_words = sorted(grouped[top], key=lambda item: float(item["x0"]))
                line = " ".join(str(item["text"]) for item in line_words)
                key = _base.normalize_key(line)
                if not started:
                    if title_prefix in key:
                        started = True
                    continue
                if rows and (any(key.startswith(stop) for stop in stop_keys) or line.startswith("(Yetkili")):
                    break
                if (
                    not line or "cari donem" in key or "onceki donem" in key
                    or key.startswith("31 mart") or "kisa vadeli orta ve uzun vadeli" in key
                    or key in {"faiz ve gelir tahakkuk", "ve reeskontlari"}
                ):
                    continue
                numeric = [item for item in line_words if _base.SIGNED_AMOUNT_TOKEN_RE.match(str(item["text"]))]
                if not numeric:
                    label = _clean_label(line, title)
                    if label and not label.startswith("("):
                        # Is Bankasi prints genuinely empty credit rows without any
                        # dash/zero.  Treating those rows as wrapped prose joins them
                        # to the next populated row and shifts every later label.
                        # Known credit row labels must therefore be emitted in place.
                        label_key = _base.normalize_key(label)
                        credit_row_keys = {
                            "tuketici kredileri tp", "konut kredisi", "tasit kredisi",
                            "ihtiyac kredisi", "diger", "tuketici kredileri dovize endeksli",
                            "diger tuketici kredileri dovize endeksli",
                            "tuketici kredileri yp", "diger tuketici kredileri yp",
                            "bireysel kredi kartlari tp",
                            "bireysel kredi kartlari yp", "taksitli", "taksitsiz",
                            "personel kredileri tp", "personel kredileri dovize endeksli",
                            "personel kredileri yp",
                            "personel kredi kartlari tp", "personel kredi kartlari yp",
                            "diger personel kredileri dovize endeksli",
                            "diger personel kredileri yp", "diger personel kredileri tp",
                            "taksitli ticari krediler tp",
                            "taksitli ticari krediler dovize endeksli",
                            "taksitli ticari krediler yp", "isyeri kredileri",
                            "tasit kredileri", "ihtiyac kredileri",
                            "kurumsal kredi kartlari tp", "kurumsal kredi kartlari yp",
                            "kredili mevduat hesabi tp gercek kisi",
                            "kredili mevduat hesabi yp gercek kisi",
                            "kredili mevduat hesabi tp tuzel kisi",
                            "kredili mevduat hesabi yp tuzel kisi",
                            "toplam",
                        }
                        if label_key in credit_row_keys:
                            rows.append([label, "", "", ""])
                            pending_label = ""
                        else:
                            pending_label = " ".join(part for part in [pending_label, label] if part)
                    continue
                first_x = min(float(item["x0"]) for item in numeric)
                label_words = [str(item["text"]) for item in line_words if float(item["x0"]) < first_x]
                label = _clean_label(" ".join(part for part in [pending_label, " ".join(label_words)] if part), title)
                pending_label = ""
                values = ["", "", ""]
                for item in numeric:
                    x = float(item["x0"]); value = str(item["text"])
                    if x < width * 0.49: values[0] = value
                    elif x < width * 0.68: values[1] = value
                    elif x >= width * 0.80: values[2] = value
                    # 0.68-0.80 is the accrual/re-discount column and is intentionally excluded.
                if label:
                    rows.append([label, *values])
                    if _base.normalize_key(label).startswith("toplam"):
                        return rows
            if rows:
                return rows
    return []


def _deposit_maturity_rows(page_texts: list[str]) -> list[list[str]]:
    rows: list[list[str]] = []
    for line in _lines_after_title(page_texts, DEPOSIT_MATURITY_TITLE, max_pages=1, include_title_tail=True):
        key = _base.normalize_key(line)
        if "onceki donem" in key:
            break
        if (
            not line
            or "cari donem" in key
            or _is_deposit_maturity_header_fragment(key)
            or key in {
                "7 gun birikimli",
                "ihbarli mevduat",
            }
            or line.startswith("(")
            or line.startswith("Yetkili ")
        ):
            continue
        parsed = _parse_variable_amount_line(line, min_values=1)
        if parsed:
            rows.append([_clean_label(parsed[0], DEPOSIT_MATURITY_TITLE), *parsed[1:]])
            if _base.normalize_key(parsed[0]).startswith("toplam"):
                break
            continue
        label = _clean_label(line, DEPOSIT_MATURITY_TITLE)
        if label:
            rows.append([label])
    return rows


def _is_deposit_maturity_header_fragment(key: str) -> bool:
    header_tokens = (
        "vadesiz",
        "1 aya kadar",
        "1 3 ay",
        "3 6 ay",
        "6 ay 1 yil",
        "1 yil",
        "birikimli",
        "ihbarli",
        "mevduat",
    )
    if key in {"ustuu mevduat", "ustu mevduat"}:
        return True
    if "7 gun" in key and "birikimli" in key:
        return True
    if key.startswith("ihbarli") and "mevduat" in key:
        return True
    return "vadesiz" in key and sum(1 for token in header_tokens if token in key) >= 3


def _expected_loss_rows(page_texts: list[str]) -> list[list[str]]:
    schema = [
        "Beklenen Kredi Zararı Karşılıkları",
        "12 Aylık Beklenen Zarar Karşılığı (Birinci Aşama)",
        "Kredi Riskinde Önemli Artış (İkinci Aşama)",
        "Temerrüt (Üçüncü Aşama)",
        "Menkul Değerler Değer Düşüş Karşılıkları",
        "Gerçeğe Uygun Değer Farkı Kar veya Zarara Yansıtılan Finansal Varlıklar",
        "Gerçeğe Uygun Değer Farkı Diğer Kapsamlı Gelire Yansıtılan Finansal Varlıklar",
        (
            "İştirakler, Bağlı Ortaklıklar ve Birlikte Kontrol Edilen "
            "Ortaklıklar Değer Düşüş Karşılıkları"
        ),
        "İştirakler",
        "Bağlı Ortaklıklar",
        "Birlikte Kontrol Edilen Ortaklıklar (İş Ortaklıkları)",
        "Diğer (*)",
        "Toplam",
    ]
    schema_by_key = {_base.normalize_key(label): label for label in schema}

    def canonical_label(text: str) -> str | None:
        key = _base.normalize_key(_clean_label(text))
        if key in schema_by_key:
            return schema_by_key[key]
        if key.startswith("diger"):
            return "Diğer (*)"
        return None

    rows: list[list[str]] = []
    pending_parts: list[str] = []
    pending_value = ""
    for line in _lines_after_title(page_texts, EXPECTED_LOSS_PROVISIONS_TITLE, max_pages=1):
        key = _base.normalize_key(line)
        if key.startswith(("4 5 diger", "4 6 diger", "diger faaliyet", "e diger faaliyet")):
            break
        if (
            not line
            or "cari donem" in key
            or "onceki donem" in key
            or key.startswith("31 mart")
            or line.startswith("Yetkili ")
        ):
            continue

        parts = line.split()
        line_values: list[str] = []
        while parts and _base.SIGNED_AMOUNT_TOKEN_RE.match(parts[-1]):
            line_values.insert(0, parts.pop())
        label_part = " ".join(parts).strip()

        if not label_part and line_values:
            # Older PDFs place the amounts between the two halves of a wrapped
            # label. Keep the current-period amount until the label completes.
            if pending_parts:
                pending_value = line_values[0]
            continue

        direct_label = canonical_label(label_part)
        if direct_label is not None and not pending_parts:
            rows.append([direct_label, "", "", line_values[0] if line_values else ""])
            if direct_label == "Toplam":
                break
            continue

        if label_part:
            pending_parts.append(label_part)
        combined = " ".join(pending_parts).strip()
        completed_label = canonical_label(combined)
        if completed_label is None:
            continue
        if line_values and not pending_value:
            pending_value = line_values[0]
        rows.append([completed_label, "", "", pending_value])
        pending_parts = []
        pending_value = ""
        if completed_label == "Toplam":
            break

    return rows


def _npl_movement_rows(page_texts: list[str]) -> list[list[str]]:
    rows: list[list[str]] = []
    current_section = ""
    seen_movement_section = False
    column_heading_fragments = {"zarar", "niteligindeki", "krediler"}
    for line in _lines_after_title(page_texts, NPL_MOVEMENT_TITLE, max_pages=1, include_title_tail=True):
        key = _base.normalize_key(line)
        if (
            key.startswith(("e 8 3", "1 7 8 3", "yabanci para", "banka yonetim"))
            or line.lstrip().startswith(("(*)", "(**)", "(Yetkili"))
        ):
            break
        if (
            not line
            or key in {"iii grup iv grup v grup"}
            or key.startswith(("tahsil imkani", "sinirli krediler", "krediler krediler"))
            or line.startswith("Yetkili ")
        ):
            continue
        parsed = _parse_variable_amount_line(line)
        if parsed:
            label = _clean_label(parsed[0], NPL_MOVEMENT_TITLE)
            label_key = _base.normalize_key(label)
            values = parsed[1:]
            # A few periods print the opening balance as "Dönem Sonu
            # Bakiyesi". Only the first movement heading has this meaning;
            # the later closing-balance row must retain its original label.
            if label_key == "donem sonu bakiyesi" and not seen_movement_section:
                label = "Önceki Dönem Sonu Bakiyesi"
                label_key = "onceki donem sonu bakiyesi"
            if label_key in {
                "onceki donem sonu bakiyesi",
                "donem icinde intikal",
                "diger donuk alacak hesaplarindan giris",
                "diger donuk alacak hesaplarina cikis",
                "donem icinde tahsilat",
                "kayittan dusulen",
                "satilan",
                "kur degisimi etkisi",
                "donem sonu bakiyesi",
                "karsilik",
                "bilancodaki net bakiyesi",
            }:
                current_section = label_key
                seen_movement_section = True
            rows.append([label, *_align_npl_movement_values(label_key, values, current_section)])
            # This is the final row of the movement table.  Everything below it
            # is explanatory prose, even in periods where the footnote marker is
            # omitted on the first continuation line.
            if label_key == "bilancodaki net bakiyesi":
                break
            continue
        label = _clean_label(line, NPL_MOVEMENT_TITLE)
        label_key = _base.normalize_key(label)
        if (
            not label
            or label_key in {"e 8 2 toplam donuk alacak hareketlerine iliskin bilgiler"}
            or label_key in column_heading_fragments
        ):
            continue
        rows.append([label, "", "", ""])
    return rows


def _align_npl_movement_values(label_key: str, values: list[str], current_section: str) -> list[str]:
    if len(values) >= 3:
        return values[:3]
    if len(values) == 2:
        if current_section == "diger donuk alacak hesaplarindan giris":
            return ["", values[0], values[1]]
        if current_section == "kur degisimi etkisi":
            return [values[0], "", values[1]]
        return [values[0], values[1], ""]
    if len(values) == 1:
        if current_section in {
            "onceki donem sonu bakiyesi",
            "donem icinde intikal",
            "donem icinde tahsilat",
            "kayittan dusulen",
            "satilan",
            "kur degisimi etkisi",
            "donem sonu bakiyesi",
            "karsilik",
        }:
            return ["", "", values[0]]
        return [values[0], "", ""]
    return ["", "", ""]


def _stage3_rows(page_texts: list[str]) -> list[list[str]]:
    rows: list[list[str]] = []
    for line in _lines_after_title(page_texts, STAGE3_PROVISIONS_TITLE, max_pages=1):
        key = _base.normalize_key(line)
        if key.startswith("1 7 8"):
            break
        if (
            not line
            or "cari donem" in key
            or "onceki donem" in key
            or key.startswith("31 mart")
            or line.startswith("Yetkili ")
        ):
            continue
        parsed = _parse_amount_line(line)
        if parsed:
            rows.append([_clean_label(parsed[0], STAGE3_PROVISIONS_TITLE), *parsed[1:]])
            if _base.normalize_key(parsed[0]).startswith("toplam"):
                break
    return rows


def _other_income_rows(page_texts: list[str]) -> list[list[str]]:
    return []


def extract_raw_rows(
    page_texts: list[str],
    page_tables: list[list[list[list[str | None]]]],
    title: str,
    include_empty_amount_rows: bool = False,
) -> list[list[str]]:
    title_key = _base.normalize_key(title)
    if title_key == _base.normalize_key(CONSUMER_CREDIT_TITLE):
        schema_rows = _credit_current_rows(page_texts, title, ("1 7 4",), max_pages=2)
        position_rows = _credit_rows_by_position(title, ("1 7 4",))
        display_labels = _credit_display_labels(title)
        if len(position_rows) == len(display_labels):
            return [
                [label, *position_row[1:4]]
                for label, position_row in zip(display_labels, position_rows)
            ]
        if len(schema_rows) == len(display_labels):
            return [
                [label, *schema_row[1:4]]
                for label, schema_row in zip(display_labels, schema_rows)
            ]
    if title_key == _base.normalize_key(INSTALLMENT_COMMERCIAL_TITLE):
        schema_rows = _credit_current_rows(page_texts, title, ("e 5", "yetkili"), max_pages=1)
        position_rows = _credit_rows_by_position(title, ("e 5", "yetkili"))
        display_labels = _credit_display_labels(title)
        if len(position_rows) == len(display_labels):
            return [
                [label, *position_row[1:4]]
                for label, position_row in zip(display_labels, position_rows)
            ]
        if len(schema_rows) == len(display_labels):
            return [
                [label, *schema_row[1:4]]
                for label, schema_row in zip(display_labels, schema_rows)
            ]
    if title_key == _base.normalize_key(DEPOSIT_MATURITY_TITLE):
        rows = _deposit_maturity_rows(page_texts)
        if rows:
            return rows
    if title_key == _base.normalize_key(NPL_MOVEMENT_TITLE):
        rows = _npl_movement_rows(page_texts)
        if rows:
            return rows
    if title_key == _base.normalize_key(EXPECTED_LOSS_PROVISIONS_TITLE):
        rows = _expected_loss_rows(page_texts)
        if rows:
            return rows
    if title_key == _base.normalize_key(OTHER_OPERATING_INCOME_TITLE):
        return _other_income_rows(page_texts)
    if title_key == _base.normalize_key(STAGE3_PROVISIONS_TITLE):
        rows = _stage3_rows(page_texts)
        if rows:
            return rows
    if title_key == _base.normalize_key(BANKS_INTEREST_INCOME_TITLE):
        rows = extract_banks_interest_income_rows(page_texts)
        if rows:
            return rows
    if title_key == _base.normalize_key(LOAN_INTEREST_INCOME_TITLE):
        rows = extract_loan_interest_income_rows(page_texts)
        if rows:
            return rows
    if title_key == _base.normalize_key(CASH_AND_TCMB_TITLE):
        rows = extract_cash_and_tcmb_rows(page_texts)
        if rows:
            return rows
    if title_key == _base.normalize_key(SECURITIES_INTEREST_INCOME_TITLE):
        rows = extract_securities_interest_rows(page_texts)
        if rows:
            return rows
    if title_key == _base.normalize_key(USED_LOANS_INTEREST_EXPENSE_TITLE):
        rows = extract_used_loans_interest_rows(page_texts)
        if rows:
            return rows
    return _ORIGINAL_EXTRACT_RAW_ROWS(
        page_texts,
        page_tables,
        title,
        include_empty_amount_rows=include_empty_amount_rows,
    )


def extract_restructured_npl_rows(page_texts: list[str]) -> list[list[str]]:
    """Extract current-period restructured NPL rows across title variants.

    Reports through 2021Q3 spell ``bankaca`` as one word, while later reports
    use ``Banka’ca``.  Matching the disclosure semantically also tolerates the
    embedded fonts that replace Turkish characters in older PDFs.
    """
    canonical_labels = [
        "(Karşılıklardan Önceki Brüt Tutarlar)",
        "Yeniden Yapılandırılan Krediler",
    ]
    for text in page_texts:
        in_section = False
        output: list[list[str]] = []
        for raw_line in text.splitlines():
            line = _base.normalize_text(raw_line)
            key = _base.normalize_key(line)
            if not in_section:
                if "donuk alacaklardan" in key and "yeniden" in key:
                    in_section = True
                continue
            if "onceki donem" in key:
                break
            if line.startswith(("Yetkili ", "(Yetkili ")):
                break
            parsed = _parse_amount_line(line)
            if not parsed:
                continue
            values = parsed[1:]
            if len(values) == 2:
                # 2021Q2 has no III Group amount; preserve the leading blank.
                values = ["", *values]
            elif len(values) > 3:
                values = values[-3:]
            while len(values) < 3:
                values.append("")
            output.append([canonical_labels[len(output)], *values[:3]])
            if len(output) == len(canonical_labels):
                return output
    return []


def extract_stage_provision_rows(page_texts: list[str]) -> list[list[str]]:
    """Extract the current-period standard/watch-list provision amounts.

    The e.2 disclosure contains two consecutive tables. The workbook target is
    the current/prior-period provision table. The ageing table below it
    (31-60 / 61-90 days) is a different disclosure and must not be mapped here.
    """
    targets = (
        (
            "12 aylik beklenen zarar karsiligi",
            "12 Aylık Beklenen Zarar Karşılığı",
            0,
        ),
        (
            "kredi riskinde onemli artis",
            "Kredi Riskinde Önemli Artış",
            1,
        ),
    )
    for text in page_texts:
        text_key = _base.normalize_key(text)
        if not (
            "standart nitelikli" in text_key
            and "yakin izlemedeki" in text_key
            and "12 aylik beklenen zarar karsiligi" in text_key
            and "kredi riskinde onemli artis" in text_key
        ):
            continue

        output: list[list[str]] = []
        for raw_line in text.splitlines():
            line = _base.normalize_text(raw_line)
            parsed = _parse_amount_line(line)
            if not parsed or len(parsed) < 2:
                continue
            label_key = _base.normalize_key(parsed[0])
            for target_key, canonical_label, target_column in targets:
                if label_key != target_key:
                    continue
                values = ["", ""]
                # First amount is current period; the second, if present, is
                # the comparative period printed alongside it.
                values[target_column] = parsed[1]
                output.append([canonical_label, *values])
                break
        if len(output) == len(targets):
            return output
    return []


def extract_stage_ageing_rows(page_texts: list[str]) -> list[list[str]]:
    """Extract current-period 31-60 / 61-90 day ageing table totals.

    This table belongs to the separate "Birinci ve ikinci grup krediler..."
    workbook block and must not be confused with the provision table above it.
    """
    target_labels = {
        "kurumsal ticari krediler": "Kurumsal / Ticari Krediler",
        "tuketici kredileri": "Tüketici Kredileri",
        "kredi kartlari": "Kredi Kartları",
        "toplam": "Toplam",
    }
    for text in page_texts:
        text_key = _base.normalize_key(text)
        if not (
            "31 60 gun" in text_key
            and "61 90 gun" in text_key
            and "kurumsal ticari krediler" in text_key
        ):
            continue

        output: list[list[str]] = []
        in_current_period = False
        for raw_line in text.splitlines():
            line = _base.normalize_text(raw_line)
            key = _base.normalize_key(line)
            if "cari donem" in key and "31 60 gun" in key and "61 90 gun" in key:
                in_current_period = True
                continue
            if in_current_period and "onceki donem" in key:
                break
            if not in_current_period:
                continue

            parsed = _parse_amount_line(line)
            if not parsed:
                continue
            label_key = _base.normalize_key(parsed[0])
            canonical_label = target_labels.get(label_key)
            if canonical_label is None or len(parsed) < 2:
                continue
            output.append([canonical_label, "", "", parsed[-1]])
            if label_key == "toplam":
                return output
        if output:
            return output
    return []


def extract_deposit_interest_rows(page_texts: list[str]) -> list[list[str]]:
    positioned = _deposit_interest_rows_by_position()
    if positioned:
        return positioned
    rows: list[list[str]] = []
    section = ""
    for line in _lines_after_title(page_texts, DEPOSIT_INTEREST_TITLE, max_pages=1):
        key = _base.normalize_key(line)
        if key.startswith("4 3 ticari"):
            break
        if (
            not line
            or "cari donem" in key
            or "onceki donem" in key
            or key.startswith(("vadeli mevduat", "vadesiz", "hesap adi"))
            or line.startswith("Yetkili ")
        ):
            continue
        if key in {"turk parasi", "yabanci para"}:
            section = line
            continue
        if key == "7 gun ihbarli mevduat":
            rows.append([f"{section} - {line}" if section else line, "", "", "", "", "", "", "", ""])
            continue
        parsed = _parse_amount_line(line)
        if not parsed:
            continue
        label = _clean_label(parsed[0], DEPOSIT_INTEREST_TITLE)
        if section and not _base.normalize_key(label).startswith(("toplam", "genel toplam")):
            label = f"{section} - {label}"
        values = _align_deposit_interest_values(_base.normalize_key(label), parsed[1:])
        rows.append([label, *values])
        if _base.normalize_key(label) == "genel toplam":
            break
    return rows


def _deposit_interest_rows_by_position() -> list[list[str]]:
    """Map the eight maturity columns from their PDF x coordinates."""
    if _ACTIVE_PDF_PATH is None:
        return []
    schema = [
        "Türk Parası - Bankalar Mevduatı",
        "Türk Parası - Tasarruf Mevduatı",
        "Türk Parası - Resmi Mevduat",
        "Türk Parası - Ticari Mevduat",
        "Türk Parası - Diğer Mevduat",
        "Türk Parası - 7 Gün İhbarlı Mevduat",
        "Türk Parası - Toplam",
        "Yabancı Para - DTH",
        "Yabancı Para - Bankalar Mevduatı",
        "Yabancı Para - 7 Gün İhbarlı Mevduat",
        "Yabancı Para - Kıymetli Maden D. Hs.",
        "Yabancı Para - Toplam",
        "Genel Toplam",
    ]
    canonical_by_key = {_base.normalize_key(label): label for label in schema}
    title_prefix = " ".join(_base.normalize_key(DEPOSIT_INTEREST_TITLE).split()[:5])

    with pdfplumber.open(str(_ACTIVE_PDF_PATH)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            if title_prefix not in _base.normalize_key(text):
                continue
            words = page.extract_words(x_tolerance=1, y_tolerance=2)
            width = float(page.width)
            grouped: dict[float, list[dict[str, object]]] = defaultdict(list)
            for word in words:
                grouped[round(float(word["top"]), 1)].append(word)

            title_top = None
            section_top = None
            for top in sorted(grouped):
                line_words = sorted(grouped[top], key=lambda item: float(item["x0"]))
                line_key = _base.normalize_key(
                    " ".join(str(item["text"]) for item in line_words)
                )
                if title_top is None and title_prefix in line_key:
                    title_top = top
                elif title_top is not None and line_key == "turk parasi":
                    section_top = top
                    break
            header_words = [
                item
                for item in words
                if title_top is not None
                and section_top is not None
                and title_top < float(item["top"]) < section_top
            ]
            vadesiz_words = [
                item for item in header_words
                if _base.normalize_key(str(item["text"])) == "vadesiz"
            ]
            toplam_words = [
                item for item in header_words
                if _base.normalize_key(str(item["text"])) == "toplam"
            ]
            left_center = (
                (float(vadesiz_words[-1]["x0"]) + float(vadesiz_words[-1]["x1"])) / 2
                if vadesiz_words else width * 0.326
            )
            right_center = (
                (float(toplam_words[-1]["x0"]) + float(toplam_words[-1]["x1"])) / 2
                if toplam_words else width * 0.862
            )
            column_centers = [
                left_center + (right_center - left_center) * index / 7
                for index in range(8)
            ]

            started = False
            section = ""
            found: dict[str, list[str]] = {}
            positional_rows: list[list[str]] = []
            numeric_only_rows: list[list[str]] = []
            for top in sorted(grouped):
                line_words = sorted(grouped[top], key=lambda item: float(item["x0"]))
                line = " ".join(str(item["text"]) for item in line_words)
                key = _base.normalize_key(line)
                if not started:
                    if title_prefix in key:
                        started = True
                    continue
                if line.startswith("(Yetkili"):
                    break
                if key in {"turk parasi", "yabanci para"}:
                    section = "Türk Parası" if key == "turk parasi" else "Yabancı Para"
                    continue
                if (
                    not line
                    or "cari donem" in key
                    or key.startswith(("vadeli mevduat", "vadesiz"))
                    or key in {"kadar kadar kadar kadar uzun mev"}
                ):
                    continue

                numeric = [
                    item
                    for item in line_words
                    if _base.SIGNED_AMOUNT_TOKEN_RE.match(str(item["text"]))
                ]
                if (
                    title_top is not None
                    and top > title_top
                    and numeric
                    and "kadar" not in key
                    and "cari donem" not in key
                ):
                    positional_values = [""] * 8
                    for item in numeric:
                        x = (float(item["x0"]) + float(item["x1"])) / 2
                        column = min(
                            range(8),
                            key=lambda index: abs(x - column_centers[index]),
                        )
                        positional_values[column] = str(item["text"])
                    positional_rows.append(positional_values)
                    if len(numeric) == len(line_words):
                        numeric_only_rows.append(positional_values)
                label_words = [
                    str(item["text"])
                    for item in line_words
                    if float(item["x0"]) < width * 0.31
                    and not _base.SIGNED_AMOUNT_TOKEN_RE.match(str(item["text"]))
                ]
                raw_label = _clean_label(" ".join(label_words), DEPOSIT_INTEREST_TITLE)
                raw_key = _base.normalize_key(raw_label)
                if raw_key == "genel toplam":
                    canonical = "Genel Toplam"
                else:
                    canonical = canonical_by_key.get(
                        _base.normalize_key(f"{section} - {raw_label}")
                    )
                if canonical is None:
                    continue

                values = [""] * 8
                for item in numeric:
                    x = (float(item["x0"]) + float(item["x1"])) / 2
                    value = str(item["text"])
                    column = min(
                        range(8),
                        key=lambda index: abs(x - column_centers[index]),
                    )
                    values[column] = value
                found[canonical] = values
                if canonical == "Genel Toplam" and any(values):
                    return [[label, *found.get(label, [""] * 8)] for label in schema]
            fallback_rows = numeric_only_rows if len(numeric_only_rows) >= 11 else positional_rows
            if not any(found.get("Genel Toplam", [])) and len(fallback_rows) >= 11:
                # Some 2021 PDFs use an embedded font whose Turkish labels are
                # extracted as replacement characters.  The row order is stable,
                # so retain the PDF's blank 7-day rows and map numeric rows only.
                numeric_schema_indexes = [0, 1, 2, 3, 4, 6, 7, 8, 10, 11, 12]
                for schema_index, values in zip(numeric_schema_indexes, fallback_rows[-11:]):
                    found[schema[schema_index]] = values
            if found:
                return [[label, *found.get(label, [""] * 8)] for label in schema]
    return []


def _align_deposit_interest_values(label_key: str, values: list[str]) -> list[str]:
    if len(values) >= 8:
        return values[:8]
    if not values:
        return [""] * 8
    if label_key == "turk parasi tasarruf mevduati" and len(values) == 7:
        return ["", *values]
    if label_key == "turk parasi resmi mevduat" and len(values) == 5:
        return ["", values[0], values[1], values[2], values[3], "", "", values[4]]
    if label_key in {"turk parasi ticari mevduat", "turk parasi diger mevduat", "yabanci para bankalar mevduati"} and len(values) == 6:
        return ["", values[0], values[1], values[2], values[3], values[4], "", values[5]]
    if label_key == "yabanci para kiymetli maden d hs":
        return ["", "", values[0], values[1], values[2], values[3], "", values[4]] if len(values) == 5 else [""] * 8
    if label_key in {"yabanci para dth", "toplam"} and len(values) == 7:
        return [*values[:-1], "", values[-1]]
    if label_key == "turk parasi bankalar mevduati" and len(values) == 5:
        return [values[0], values[1], values[2], "", values[3], "", "", values[4]]
    output = values[:]
    while len(output) < 8:
        output.insert(-1 if output else 0, "")
    return output[:8]


def to_deposit_summary(rows: list[list[str]]) -> list[list[str]]:
    output: list[list[str]] = []
    for row in rows:
        label = _base.normalize_text(row[0])
        vadesiz = _base.normalize_text(row[1] if len(row) > 1 else "")
        toplam = _base.normalize_text(row[-1] if len(row) > 1 else "")
        toplam_amount = _base.parse_amount_text(toplam)
        vadesiz_amount = _base.parse_amount_text(vadesiz)
        if toplam_amount is None:
            output.append([label, vadesiz, "", ""])
            continue
        vadeli = toplam_amount - (vadesiz_amount or 0)
        output.append([label, vadesiz, _base.format_amount(vadeli), toplam])
    return output


def _rows_with_amount_continuation(page_texts: list[str], title: str, stop_keys: tuple[str, ...]) -> list[list[str]]:
    rows: list[list[str]] = []
    pending_label = ""
    pending_values: list[str] | None = None
    for line in _lines_after_title(page_texts, title, max_pages=1):
        key = _base.normalize_key(line)
        if rows and any(key.startswith(stop_key) for stop_key in stop_keys):
            break
        if (
            not line
            or "cari donem" in key
            or "onceki donem" in key
            or key in {"tp yp tp yp", "tp yp"}
            or line.startswith("Yetkili ")
        ):
            continue
        if pending_values is not None:
            if not _is_amount_only_line(line):
                rows.append([_clean_label(f"{pending_label} {line}".strip(), title), *pending_values])
                pending_label = ""
                pending_values = None
                continue
            rows.append([_clean_label(pending_label, title), *pending_values])
            pending_label = ""
            pending_values = None

        if _is_amount_only_line(line) and pending_label:
            pending_values = line.split()
            continue
        parsed = _parse_amount_line(line)
        if parsed:
            label = _clean_label(" ".join(part for part in [pending_label, parsed[0]] if part), title)
            rows.append([label, *parsed[1:]])
            pending_label = ""
            if _base.normalize_key(label).startswith("toplam"):
                break
            continue
        pending_label = " ".join(part for part in [pending_label, line] if part).strip()
    if pending_values is not None:
        rows.append([_clean_label(pending_label, title), *pending_values])
    return rows


def extract_banks_interest_income_rows(page_texts: list[str]) -> list[list[str]]:
    rows: list[list[str]] = []
    raw_lines = _lines_after_title(
        page_texts,
        BANKS_INTEREST_INCOME_TITLE,
        max_pages=1,
        include_title_tail=True,
    )
    tcmb_candidates: list[str] = []
    for line in raw_lines:
        parsed = _parse_variable_amount_line(line, min_values=1)
        if parsed and _base.normalize_key(parsed[0]) == "t c merkez bankasindan":
            tcmb_candidates = parsed[1:]
            break
    raw_rows = _combine_wrapped_rows(
        raw_lines,
        stop_keys=("a 3",),
    )
    for row in raw_rows:
        label = row[0]
        key = _base.normalize_key(label)
        values = row[1:]
        if key.startswith("t c merkez bankasindan") and " yurt ici bankalardan" in key:
            rows.append(["T.C. Merkez Bankasından", "", ""])
            rows.append(["Yurt İçi Bankalardan", *_current_pair(values)])
            continue
        if key == "t c merkez bankasindan" and len(values) == 2:
            rows.append([label, "", values[0]])
            continue
        if key == "yurt disi merkez ve subelerden toplam":
            rows.append(["Yurt Dışı Merkez ve Şubelerden", "", ""])
            rows.append(["Toplam", *_current_pair(values)])
            continue
        rows.append([label, *_current_pair(values)])

    # A single amount on the T.C. Central Bank row can belong either to the
    # current or the previous period. Plain PDF text loses its x-coordinate,
    # so reconcile the current-period pair to the published total. Only fill
    # a residual when that exact amount is visibly present on the TCMB row.
    by_key = {_base.normalize_key(row[0]): row for row in rows}
    tcmb_row = by_key.get("t c merkez bankasindan")
    total_row = by_key.get("toplam")
    if tcmb_row and total_row and tcmb_candidates:
        candidate_amounts = {
            amount
            for amount in (_base.parse_amount_text(value) for value in tcmb_candidates)
            if amount is not None
        }
        component_keys = (
            "yurt ici bankalardan",
            "yurt disi bankalardan",
            "yurt disi merkez ve subelerden",
        )
        for value_index in (1, 2):
            if _base.parse_amount_text(tcmb_row[value_index]) is not None:
                continue
            total = _base.parse_amount_text(total_row[value_index])
            if total is None:
                continue
            component_total = sum(
                _base.parse_amount_text(by_key.get(key, ["", "", ""])[value_index]) or 0
                for key in component_keys
            )
            residual = total - component_total
            if residual > 0 and residual in candidate_amounts:
                tcmb_row[value_index] = f"{residual:,}".replace(",", ".")
    return rows


def extract_securities_interest_rows(page_texts: list[str]) -> list[list[str]]:
    return _current_rows(_rows_with_amount_continuation(page_texts, SECURITIES_INTEREST_INCOME_TITLE, ("a 4 istirak",)))


def extract_loan_interest_income_rows(page_texts: list[str]) -> list[list[str]]:
    rows: list[list[str]] = []
    pending_label = ""
    for line in _lines_after_title(page_texts, LOAN_INTEREST_INCOME_TITLE, max_pages=1, include_title_tail=True):
        key = _base.normalize_key(line)
        if rows and key.startswith("a 2"):
            break
        if (
            not line
            or "cari donem" in key
            or "onceki donem" in key
            or key.startswith("31 mart")
            or key in {"tp yp tp yp", "tp yp"}
            or line.startswith("(")
            or line.startswith("Yetkili ")
        ):
            continue
        if key == "kredilerden alinan faizler":
            rows.append(["Kredilerden Alınan Faizler (*)", "", ""])
            continue
        parsed = _parse_amount_line(line)
        if parsed:
            if _base.normalize_key(parsed[0]).startswith("toplam"):
                if pending_label:
                    rows.append([_clean_label(pending_label, LOAN_INTEREST_INCOME_TITLE), "", ""])
                pending_label = ""
            label = _clean_label(" ".join(part for part in [pending_label, parsed[0]] if part).strip(), LOAN_INTEREST_INCOME_TITLE)
            rows.append([label, *_current_pair(parsed[1:])])
            pending_label = ""
            if _base.normalize_key(label).startswith("toplam"):
                break
            continue
        pending_label = " ".join(part for part in [pending_label, line] if part).strip()
    return rows


def extract_used_loans_interest_rows(page_texts: list[str]) -> list[list[str]]:
    rows: list[list[str]] = []
    blank_labels = {
        "t c merkez bankasina",
        "yurt disi merkez ve subelere",
    }
    for line in _lines_after_title(page_texts, USED_LOANS_INTEREST_EXPENSE_TITLE, max_pages=1):
        key = _base.normalize_key(line)
        if key.startswith("b 2 istirak"):
            break
        if (
            not line
            or "cari donem" in key
            or "onceki donem" in key
            or key in {"tp yp tp yp", "tp yp"}
            or line.startswith("Yetkili ")
        ):
            continue
        parsed = _parse_amount_line(line)
        if parsed:
            label = _clean_label(parsed[0], USED_LOANS_INTEREST_EXPENSE_TITLE)
            if _base.normalize_key(label) == "diger kuruluslara" and len(parsed) == 3:
                rows.append([label, "", parsed[1]])
            else:
                rows.append([label, *_current_pair(parsed[1:])])
        elif key in blank_labels:
            rows.append([line, "", ""])
    return rows


def extract_cash_and_tcmb_rows(page_texts: list[str]) -> list[list[str]]:
    """Extract current TP/YP while preserving the blank TP cell of ``Diğer``."""
    canonical = {
        "kasa efektif": "Kasa/Efektif",
        "tcmb": "TCMB",
        "diger": "Diğer",
        "toplam": "Toplam",
    }
    for text in page_texts:
        text_key = _base.normalize_key(text)
        if "nakit degerler ve tcmb" not in text_key or "kasa efektif" not in text_key:
            continue
        in_section = False
        rows: list[list[str]] = []
        for raw_line in text.splitlines():
            line = _base.normalize_text(raw_line)
            key = _base.normalize_key(line)
            if not in_section:
                if "nakit degerler ve tcmb" in key and "iliskin bilgiler" in key:
                    in_section = True
                continue
            if key.startswith(("a 2", "t c merkez bankasi hesabina")):
                break
            parsed = _parse_variable_amount_line(line, min_values=1)
            if not parsed:
                continue
            label_key = _base.normalize_key(parsed[0])
            label = next(
                (value for target, value in canonical.items() if label_key.endswith(target)),
                None,
            )
            if label is None:
                continue
            values = parsed[1:]
            if label == "Diğer" and len(values) == 2:
                current = ["", values[0]]
            else:
                current = values[:2]
                while len(current) < 2:
                    current.append("")
            rows.append([label, *current])
            if label == "Toplam":
                return rows
        if rows:
            return rows
    return []


def extract_is_bankasi_other_operating_expense_rows(
    page_texts: list[str],
    page_words: list[list[dict[str, object]]],
) -> list[list[str]]:
    """Join wrapped İş Bankası expense labels without merging adjacent rows."""
    schema = [
        "Kıdem Tazminatı Karşılığı",
        "Banka Sosyal Yardım Sandığı Varlık Açıkları Karşılığı",
        "Maddi Duran Varlık Değer Düşüş Giderleri",
        "Maddi Duran Varlık Amortisman Giderleri",
        "Maddi Olmayan Duran Varlık Değer Düşüş Giderleri",
        "Şerefiye Değer Düşüş Gideri",
        "Maddi Olmayan Duran Varlık Amortisman Giderleri",
        "Özkaynak Yöntemi Uygulanan Ortaklık Payları Değer Düşüş Gideri",
        "Elden Çıkarılacak Kıymetler Değer Düşüş Giderleri",
        "Elden Çıkarılacak Kıymetler Amortisman Giderleri",
        (
            "Satış Amaçlı Elde Tutulan ve Durdurulan Faaliyetlere İlişkin "
            "Duran Varlıklar Değer Düşüş Giderleri"
        ),
        "Diğer İşletme Giderleri",
        "TFRS 16 İstisnalarına İlişkin Kiralama Giderleri",
        "Bakım ve Onarım Giderleri",
        "Reklam ve İlan Giderleri",
        "Diğer Giderler",
        "Aktiflerin Satışından Doğan Zararlar",
        "Diğer (*)",
        "Toplam",
    ]
    schema_by_key = {_base.normalize_key(label): label for label in schema}

    # Keep the positioned extractor's correct current/previous column
    # distinction, then split any adjacent İş Bankası labels it joined.
    positioned = _ORIGINAL_EXTRACT_OTHER_EXPENSE_WORDS(page_texts, page_words)
    if positioned:
        expanded: list[list[str]] = []
        schema_index = 0
        for source_row in positioned:
            source_key = _base.normalize_key(source_row[0])
            matched_end: int | None = None
            for end in range(schema_index + 1, len(schema) + 1):
                joined_key = _base.normalize_key(" ".join(schema[schema_index:end]))
                if source_key == joined_key:
                    matched_end = end
                    break
            if matched_end is None:
                for end in range(schema_index + 1, len(schema) + 1):
                    joined_key = _base.normalize_key(" ".join(schema[schema_index:end]))
                    if source_key.startswith(joined_key) or joined_key.startswith(source_key):
                        matched_end = end
                        break
            if matched_end is None:
                continue
            # PDF word extraction can concatenate a value-bearing amortisation
            # row with the following blank impairment row.  In that layout the
            # amount belongs to the first label, not the final concatenated
            # label.  This is the pattern used by Is Bank from 2021Q4 onward.
            value_on_first_label = (
                schema_index < len(schema)
                and schema[schema_index]
                in {
                    "Maddi Duran Varlık Amortisman Giderleri",
                    "Maddi Olmayan Duran Varlık Amortisman Giderleri",
                }
                and matched_end > schema_index + 1
            )
            for index in range(schema_index, matched_end):
                value_index = schema_index if value_on_first_label else matched_end - 1
                value = source_row[3] if index == value_index else "-"
                expanded.append([schema[index], "", "", value or "-"])
            schema_index = matched_end
            if schema_index == len(schema):
                break
        if schema_index == len(schema):
            return expanded

    def canonical_label(text: str) -> str | None:
        key = _base.normalize_key(text)
        if key in schema_by_key:
            return schema_by_key[key]
        if key.startswith("reklam ve ilan giderleri"):
            return "Reklam ve İlan Giderleri"
        if key.startswith("diger giderler"):
            return "Diğer Giderler"
        if key == "diger":
            return "Diğer (*)"
        return None

    rows: list[list[str]] = []
    pending: list[str] = []
    for line in _lines_after_title(
        page_texts,
        OTHER_OPERATING_EXPENSE_TITLE,
        max_pages=1,
    ):
        key = _base.normalize_key(line)
        if line.startswith(("(*)", "(**)", "(***)", "(Yetkili", "Yetkili ")):
            break
        if not line or "cari donem" in key or "onceki donem" in key:
            continue

        parts = line.split()
        values: list[str] = []
        while parts and _base._expense_amount_token(parts[-1]):
            values.insert(0, parts.pop().rstrip(".,;"))
        label_part = " ".join(parts).strip()
        if not label_part:
            continue

        if pending:
            combined = " ".join([*pending, label_part]).strip()
            completed = canonical_label(combined)
            if completed is not None:
                rows.append([completed, "", "", values[0] if values else "-"])
                pending = []
                if completed == "Toplam":
                    break
                continue

        direct = canonical_label(label_part)
        if direct is not None:
            rows.append([direct, "", "", values[0] if values else "-"])
            pending = []
            if direct == "Toplam":
                break
            continue
        pending.append(label_part)

    return rows


def extract_equity_rows(page_texts: list[str]) -> list[list[str]]:
    targets = {
        "cekirdek sermaye toplami": ("Çekirdek Sermaye Toplamı", None),
        "katki sermaye toplami": ("Katkı Sermaye Toplamı", None),
        "toplam risk agirlikli tutarlar": ("Toplam Risk Ağırlıklı Tutarlar", None),
    }
    contextual_total_equity: str | None = None
    for start_page in _base.title_pages(page_texts, EQUITY_ITEMS_TITLE):
        found_any = False
        section_total_equity: str | None = None
        for text in page_texts[start_page : min(len(page_texts), start_page + 4)]:
            for raw_line in text.splitlines():
                line = _base.normalize_text(raw_line)
                key = _base.normalize_key(line)
                parsed = _parse_variable_amount_line(line, min_values=1)
                if not parsed:
                    continue
                if "toplam ozkaynak ana sermaye ve katki sermaye toplami" in key:
                    section_total_equity = parsed[1]
                    continue
                if "toplam risk agirlikli tutarlar" in key and section_total_equity is not None:
                    contextual_total_equity = section_total_equity
                for target_key, (label, value) in list(targets.items()):
                    if value is None and target_key in key:
                        targets[target_key] = (label, parsed[1])
                        found_any = True
        if found_any and contextual_total_equity is not None and all(value is not None for _, value in targets.values()):
            break

    rows: list[list[str]] = []
    missing: list[str] = []
    ordered_values = [
        targets["cekirdek sermaye toplami"],
        targets["katki sermaye toplami"],
        ("Toplam Özkaynak (Ana sermaye ve katkı sermaye toplamı)", contextual_total_equity),
        targets["toplam risk agirlikli tutarlar"],
    ]
    for label, value in ordered_values:
        if value is None:
            missing.append(label)
        else:
            rows.append([label, "", "", value])
    if missing:
        raise RuntimeError("Hedef özkaynak kalemleri bulunamadı: " + ", ".join(missing))
    return rows


def extract_liquidity_coverage_rows(page_texts: list[str]) -> list[list[str]]:
    for start_page in _base.title_pages(page_texts, LIQUIDITY_COVERAGE_RATIO_TITLE):
        for text in page_texts[start_page : min(len(page_texts), start_page + 3)]:
            for raw_line in text.splitlines():
                line = _base.normalize_text(raw_line)
                key = _base.normalize_key(line)
                if "likidite karsilama orani" not in key:
                    continue
                values = [token for token in line.split() if re.fullmatch(r"[0-9]+(?:\.[0-9]{3})*,[0-9]+", token)]
                if len(values) >= 2:
                    return [["Likidite Karşılama Oranı (%)", values[-2], values[-1]]]
    return []


def to_transposed_total(rows: list[list[str]], fallback_headers: list[str]) -> list[list[str]]:
    total_row = next((row for row in rows if row and _base.normalize_key(row[0]).startswith(("toplam", "genel toplam"))), rows[-1])
    if len(total_row) - 1 == 8:
        values = total_row[1:]
        return [
            ["Vadesiz", "", "", _base.clean_cell(values[0])],
            ["7 Gün İhbarlı", "", "", ""],
            ["1 Aya Kadar", "", "", _base.clean_cell(values[1])],
            ["1-3 Ay", "", "", _base.clean_cell(values[2])],
            ["3-6 Ay", "", "", _base.clean_cell(values[3])],
            ["6 Ay-1 Yıl", "", "", _base.clean_cell(values[4])],
            ["1 Yıl ve Üstü", "", "", _base.clean_cell(values[5])],
            ["Birikimli Mevduat", "", "", _base.clean_cell(values[6])],
            ["Toplam", "", "", _base.clean_cell(values[7])],
        ]
    else:
        headers = [
            "Vadesiz",
            "1 Aya Kadar",
            "1-3 Ay",
            "3-6 Ay",
            "6 Ay-1 Yıl",
            "1 Yıl ve Üstü",
            "7 Gün İhbarlı",
            "Birikimli Mevduat",
            "Toplam",
        ][: len(total_row) - 1]
    return [[header, "", "", _base.clean_cell(value)] for header, value in zip(headers, total_row[1:])]


def to_transposed_general_total(rows: list[list[str]], fallback_headers: list[str]) -> list[list[str]]:
    total_row = next((row for row in rows if row and _base.normalize_key(row[0]) == "genel toplam"), None)
    if total_row is None:
        total_row = next((row for row in reversed(rows) if row and _base.normalize_key(row[0]).endswith("toplam")), rows[-1])
    headers = ["Vadesiz Mevduat", "1 Aya Kadar", "3 Aya Kadar", "6 Aya Kadar", "1 Yıla Kadar", "1 Yıldan Uzun", "Birikimli Mevduat", "Toplam"][
        : len(total_row) - 1
    ]
    return [[header, "", "", _base.clean_cell(value)] for header, value in zip(headers, total_row[1:])]


def _patch_is_bankasi_overrides() -> None:
    _base.extract_raw_rows = extract_raw_rows
    _base.extract_restructured_npl_rows = extract_restructured_npl_rows
    _base.extract_stage_provision_rows = extract_stage_provision_rows
    _base.extract_deposit_interest_rows = extract_deposit_interest_rows
    _base.extract_equity_rows = extract_equity_rows
    _base.extract_liquidity_coverage_rows = extract_liquidity_coverage_rows
    _base.extract_other_operating_expense_rows_from_words = (
        extract_is_bankasi_other_operating_expense_rows
    )
    _base.to_interest_row = to_interest_row
    _base.to_deposit_summary = to_deposit_summary
    _base.to_transposed_total = to_transposed_total
    _base.to_transposed_general_total = to_transposed_general_total


_patch_base()
_patch_is_bankasi_overrides()


def extract_all_tables(pdf_path: Path):
    global _ACTIVE_PDF_PATH
    previous = _ACTIVE_PDF_PATH
    _ACTIVE_PDF_PATH = Path(pdf_path)
    try:
        _patch_base()
        _patch_is_bankasi_overrides()
        tables = _ORIGINAL_EXTRACT_ALL_TABLES(Path(pdf_path))
        with pdfplumber.open(str(pdf_path)) as pdf:
            page_texts = [page.extract_text(x_tolerance=1, y_tolerance=3) or "" for page in pdf.pages]
        ageing_rows = extract_stage_ageing_rows(page_texts)
        if ageing_rows:
            tables.append(
                (
                    "Birinci ve ikinci grup krediler, diğer alacaklar ile yeniden yapılandırılan ya da yeni bir itfa planına bağlanan krediler ve bunlar için ayrılan karşılıklara ilişkin bilgiler",
                    ["Kalem", "", "", "Toplam"],
                    ageing_rows,
                )
            )
        return tables
    finally:
        _ACTIVE_PDF_PATH = previous

for _name in dir(_base):
    if not _name.startswith("__") and _name not in globals():
        globals()[_name] = getattr(_base, _name)


def main() -> None:
    _patch_base()
    _patch_is_bankasi_overrides()
    _base.main()


if __name__ == "__main__":
    main()
