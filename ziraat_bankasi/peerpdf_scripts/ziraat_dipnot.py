from __future__ import annotations

import re
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from teb_bankasi.peerpdf_scripts import teb_dipnot as _base

_ORIGINAL_EXTRACT_RAW_ROWS = _base.extract_raw_rows
_ORIGINAL_EXTRACT_STAGE_PROVISION_ROWS = _base.extract_stage_provision_rows
_ORIGINAL_EXTRACT_OTHER_OPERATING_EXPENSE_ROWS = _base.extract_other_operating_expense_rows
_ORIGINAL_EXTRACT_OTHER_OPERATING_EXPENSE_ROWS_FROM_WORDS = _base.extract_other_operating_expense_rows_from_words


CONSUMER_CREDIT_TITLE = (
    "Tüketici kredileri, bireysel kredi kartları, personel kredileri "
    "ve personel kredi kartlarına ilişkin bilgiler"
)
INSTALLMENT_COMMERCIAL_TITLE = "Taksitli ticari krediler ve kurumsal kredi kartlarına ilişkin bilgiler"
DEPOSIT_MATURITY_TITLE = "Mevduatın Vade Yapısına İlişkin Bilgiler"
DEPOSIT_TOTAL_BY_MATURITY_TITLE = "Mevduatın Vade Yapısına İlişkin Toplam Bilgileri"
DEPOSIT_INTEREST_TITLE = "Mevduata Ödenen Faizin Vade Yapısına Göre Gösterimi"
RESTRUCTURED_NPL_TITLE = "Donuk Alacaklardan Yeniden Yapılandırılan Kredilere İlişkin Bilgiler"
NPL_MOVEMENT_TITLE = "Toplam donuk alacak hareketlerine ilişkin bilgiler"
BANKS_INTEREST_INCOME_TITLE = "Bankalardan alınan faiz gelirlerine ilişkin bilgiler"
SECURITIES_INTEREST_INCOME_TITLE = "Menkul Değerlerden Alınan Faizlere İlişkin Bilgiler"
EQUITY_ITEMS_TITLE = "Özkaynak kalemlerine ilişkin bilgiler"
EXPECTED_LOSS_PROVISIONS_TITLE = "Beklenen zarar karşılıkları ve diğer karşılık giderleri"
STAGE_PROVISION_TITLE = (
    "Standart Nitelikli ve Yakın İzlemedeki (Birinci ve İkinci Grup Krediler) "
    "İle Yeniden Yapılandırılan Yakın İzlemedeki Kredilere İlişkin Bilgiler"
)
STAGE_PROVISION_HEADERS = ["Kalem", "Standart Nitelikli Krediler", "Yakın İzlemedeki Krediler"]
LIQUIDITY_COVERAGE_RATIO_TITLE = "Likidite karşılama oranı"
STAGE3_PROVISIONS_TITLE = "Kredilere İlişkin Olarak Ayrılan Beklenen Zarar Karşılıkları (Üçüncü Aşama)"
OTHER_OPERATING_INCOME_TITLE = "Diğer faaliyet gelirlerine ilişkin açıklamalar"
CASH_AND_TCMB_TITLE = "Nakit Değerler ve T.C. Merkez Bankası’na İlişkin Bilgiler"
USED_LOANS_INTEREST_EXPENSE_TITLE = "Kullanılan Kredilere Verilen Faizlere İlişkin Bilgiler"
LOAN_INTEREST_INCOME_TITLE = "Kredilerden alınan faiz gelirlerine ilişkin bilgiler"
OTHER_OPERATING_EXPENSE_TITLE = "Diğer faaliyet giderlerine ilişkin bilgiler"

OTHER_OPERATING_EXPENSE_LABELS = [
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
    "Satış Amaçlı Elde Tutulan ve Durdurulan Faaliyetlere İlişkin Duran Varlıklar Değer Düşüş Giderleri",
    "Diğer İşletme Giderleri",
    "TFRS 16 İstisnalarına İlişkin Kiralama Giderleri",
    "Bakım ve Onarım Giderleri",
    "Reklam ve İlan Giderleri",
    "Diğer Giderler",
    "Aktiflerin Satışından Doğan Zararlar",
    "Diğer",
    "Toplam",
]

_CONSUMER_PARENT_LABELS = {
    "tuketici kredileri tp",
    "tuketici kredileri dovize endeksli",
    "tuketici kredileri yp",
    "bireysel kredi kartlari tp",
    "bireysel kredi kartlari yp",
    "personel kredileri tp",
    "personel kredileri dovize endeksli",
    "personel kredileri yp",
    "personel kredi kartlari tp",
    "personel kredi kartlari yp",
}
_CONSUMER_CHILD_LABELS = {
    "konut kredisi", "tasit kredisi", "ihtiyac kredisi", "diger", "taksitli", "taksitsiz",
}


def canonicalize_consumer_credit_rows(rows: list[list[str]]) -> list[list[str]]:
    """Make repeated child labels unique by retaining their PDF parent-group context."""
    canonical: list[list[str]] = []
    parent = ""
    for raw in rows:
        if not raw:
            continue
        row = list(raw)
        label = _base.normalize_text(row[0])
        label = re.sub(r"\s*\([12]\)\s*$", "", label).strip()
        key = _base.normalize_key(label)
        if key in _CONSUMER_PARENT_LABELS:
            parent = label
            row[0] = label
        elif key in _CONSUMER_CHILD_LABELS and parent:
            row[0] = f"{parent} - {label}"
        elif key.startswith("toplam"):
            parent = ""
            row[0] = "Toplam"
        else:
            row[0] = label
        canonical.append(row)
    labels = [_base.normalize_key(row[0]) for row in canonical]
    if len(labels) != len(set(labels)):
        raise RuntimeError("Ziraat tuketici kredileri kanonik satir adlari tekil degil")
    return canonical


def _patch_base() -> None:
    names = [
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
        "OTHER_OPERATING_EXPENSE_TITLE",
    ]
    for name in names:
        setattr(_base, name, globals()[name])
    _base.STAGE_PROVISION_HEADERS = STAGE_PROVISION_HEADERS
    _base.STAGE_PROVISION_TARGET_LABELS = [
        "12 Aylık Beklenen Zarar Karşılığı",
        "Kredi Riskinde Önemli Artış",
    ]


def extract_raw_rows(
    page_texts: list[str],
    page_tables: list[list[list[list[str | None]]]],
    title: str,
    include_empty_amount_rows: bool = False,
) -> list[list[str]]:
    if _base.normalize_key(title) == _base.normalize_key(CONSUMER_CREDIT_TITLE):
        rows = _ORIGINAL_EXTRACT_RAW_ROWS(
            page_texts,
            page_tables,
            title,
            include_empty_amount_rows=True,
        )
        return canonicalize_consumer_credit_rows(rows)
    if _base.normalize_key(title) == _base.normalize_key(INSTALLMENT_COMMERCIAL_TITLE):
        rows = _ORIGINAL_EXTRACT_RAW_ROWS(
            page_texts,
            page_tables,
            title,
            include_empty_amount_rows=True,
        )
        for row in rows:
            if row and _base.normalize_key(row[0]).startswith("toplam"):
                row[0] = "Toplam"
        return rows
    if _base.normalize_key(title) == _base.normalize_key(DEPOSIT_MATURITY_TITLE):
        rows = _ORIGINAL_EXTRACT_RAW_ROWS(
            page_texts,
            page_tables,
            title,
            include_empty_amount_rows=False,
        )
        if not any(_base.normalize_key(row[0]) == "diger" for row in rows if row):
            insert_at = next(
                (
                    index
                    for index, row in enumerate(rows)
                    if row and _base.normalize_key(row[0]).startswith("toplam")
                ),
                len(rows),
            )
            rows.insert(insert_at, ["Diğer", "-", "-", "-"])
        for row in rows:
            if row and _base.normalize_key(row[0]).startswith("toplam"):
                row[0] = "Toplam"
        return rows
    if _base.normalize_key(title) == _base.normalize_key(NPL_MOVEMENT_TITLE):
        rows = extract_npl_movement_rows(page_texts)
        if rows:
            return rows
    if _base.normalize_key(title) == _base.normalize_key(BANKS_INTEREST_INCOME_TITLE):
        rows = extract_current_interest_rows(page_texts, BANKS_INTEREST_INCOME_TITLE, "1 3")
        if rows:
            return rows
    if _base.normalize_key(title) == _base.normalize_key(SECURITIES_INTEREST_INCOME_TITLE):
        rows = extract_current_interest_rows(page_texts, SECURITIES_INTEREST_INCOME_TITLE, "1 4")
        if rows:
            return rows
    if _base.normalize_key(title) == _base.normalize_key(EXPECTED_LOSS_PROVISIONS_TITLE):
        rows = extract_expected_loss_rows(page_texts)
        if rows:
            return rows
    if _base.normalize_key(title) == _base.normalize_key(STAGE3_PROVISIONS_TITLE):
        rows = extract_stage3_provision_rows(page_texts, page_tables)
        if rows:
            return rows
    if _base.normalize_key(title) == _base.normalize_key(CASH_AND_TCMB_TITLE):
        rows = extract_current_tp_yp_rows(
            page_texts,
            CASH_AND_TCMB_TITLE,
            [
                "zorunlu karsiliklara iliskin aciklamalar",
                "t c merkez bankasi hesabina iliskin bilgiler",
                "2 gercege uygun deger",
            ],
        )
        if rows:
            return rows
    if _base.normalize_key(title) == _base.normalize_key(USED_LOANS_INTEREST_EXPENSE_TITLE):
        rows = extract_current_tp_yp_rows(
            page_texts,
            USED_LOANS_INTEREST_EXPENSE_TITLE,
            ["2 2 istirakler", "2 3 ihrac edilen"],
        )
        if rows:
            return rows
    if _base.normalize_key(title) == _base.normalize_key(LOAN_INTEREST_INCOME_TITLE):
        rows = extract_current_tp_yp_rows(
            page_texts,
            LOAN_INTEREST_INCOME_TITLE,
            ["1 2 bankalardan", "1 3 menkul"],
        )
        if rows:
            return rows
    if _base.normalize_key(title) == _base.normalize_key(OTHER_OPERATING_INCOME_TITLE):
        rows = extract_other_operating_income_rows(page_texts)
        if rows:
            return rows
    return _ORIGINAL_EXTRACT_RAW_ROWS(
        page_texts,
        page_tables,
        title,
        include_empty_amount_rows=include_empty_amount_rows,
    )


def extract_stage_provision_rows(page_texts: list[str]) -> list[list[str]]:
    return [row[:3] for row in _ORIGINAL_EXTRACT_STAGE_PROVISION_ROWS(page_texts)]


def canonicalize_other_operating_expense_rows(rows: list[list[str]]) -> list[list[str]]:
    target_keys = [
        re.sub(r"[^a-z0-9çğıöşü]", "", _base.normalize_key(label))
        for label in OTHER_OPERATING_EXPENSE_LABELS
    ]
    values: list[str | None] = [None] * len(OTHER_OPERATING_EXPENSE_LABELS)
    for raw in rows:
        if not raw:
            continue
        label = re.sub(r"\([12]\)", "", _base.normalize_text(raw[0]))
        source_key = re.sub(r"[^a-z0-9çğıöşü]", "", _base.normalize_key(label))
        amount = raw[-1] if len(raw) > 1 else "-"
        exact = [index for index, key in enumerate(target_keys) if source_key == key]
        matches = exact or [
            index for index, key in enumerate(target_keys[:-2])
            if len(key) > 12 and key in source_key
        ]
        if not matches:
            continue
        for index in matches[:-1]:
            if values[index] is None:
                values[index] = "-"
        values[matches[-1]] = amount
    result = [
        [label, "", "", values[index] or "-"]
        for index, label in enumerate(OTHER_OPERATING_EXPENSE_LABELS)
    ]
    if result[-1][-1] == "-":
        raise RuntimeError("Ziraat diğer faaliyet giderleri toplam satırı bulunamadı")
    return result


def extract_other_operating_expense_rows(page_texts: list[str]) -> list[list[str]]:
    return canonicalize_other_operating_expense_rows(
        _ORIGINAL_EXTRACT_OTHER_OPERATING_EXPENSE_ROWS(page_texts)
    )


def extract_other_operating_expense_rows_from_words(page_texts, page_words):
    return canonicalize_other_operating_expense_rows(
        _ORIGINAL_EXTRACT_OTHER_OPERATING_EXPENSE_ROWS_FROM_WORDS(page_texts, page_words)
    )


def extract_deposit_interest_rows(page_texts: list[str]) -> list[list[str]]:
    pages = _base.title_pages(page_texts, DEPOSIT_INTEREST_TITLE)
    if not pages:
        raise RuntimeError(f"Hedef baslik PDF icinde bulunamadi: {DEPOSIT_INTEREST_TITLE}")

    rows: list[list[str]] = []
    section = ""
    started = False
    pending_label = ""
    for page_index in pages:
        for raw_line in page_texts[page_index].splitlines():
            line = _base.normalize_text(raw_line)
            key = _base.normalize_key(line)
            if _base.normalize_key(DEPOSIT_INTEREST_TITLE) in key:
                started = True
                continue
            if not started:
                continue
            if key.startswith("onceki donem"):
                return rows
            if key == "turk parasi":
                section = "Türk Parası"
                pending_label = ""
                continue
            if key == "yabanci para":
                section = "Yabancı Para"
                pending_label = ""
                continue
            if not line or key in {"cari donem vadeli mevduat", "hesap adi mevduat 1 aya kadar 3 aya kadar 6 aya kadar kadar uzun mevduat toplam"}:
                continue

            parsed = _base.parse_text_line(line)
            if not parsed:
                if section and key not in {"vadesiz 1 yila 1 yildan birikimli"}:
                    pending_label = " ".join(part for part in [pending_label, line] if part).strip()
                continue

            label = " ".join(part for part in [pending_label, parsed[0]] if part).strip()
            pending_label = ""
            label_key = _base.normalize_key(label)
            if label_key == "genel toplam":
                rows.append(["Genel Toplam", *parsed[1:]])
                return rows
            if section and label_key not in {"hesap adi", "vadeli mevduat"}:
                rows.append([f"{section} - {label}", *parsed[1:]])
        if rows:
            break

    if not rows:
        raise RuntimeError(f"Hedef cari donem satirlari bulunamadi: {DEPOSIT_INTEREST_TITLE}")
    return rows


def extract_stage3_provision_rows(
    page_texts: list[str],
    page_tables: list[list[list[list[str | None]]]],
) -> list[list[str]]:
    for start_page in _base.title_pages(page_texts, STAGE3_PROVISIONS_TITLE):
        for tables in page_tables[start_page : min(len(page_tables), start_page + 2)]:
            for table in tables:
                parsed = _parse_current_previous_table(table)
                if (
                    parsed
                    and any(_base.normalize_key(row[0]).startswith("toplam") for row in parsed)
                    and any("ayrilanlar" in _base.normalize_key(row[0]) for row in parsed)
                ):
                    return parsed
        rows = extract_current_previous_rows_from_text(page_texts, STAGE3_PROVISIONS_TITLE)
        if rows:
            return rows
    return []


def _parse_current_previous_table(table: list[list[str | None]]) -> list[list[str]]:
    rows: list[list[str]] = []
    seen_period_header = False
    pending_label = ""
    for raw_row in table:
        cells = [_base.clean_cell(cell) for cell in raw_row]
        if not any(cells):
            continue
        row_key = _base.normalize_key(" ".join(cells))
        if "cari donem" in row_key and "onceki donem" in row_key:
            seen_period_header = True
            continue
        if not seen_period_header:
            continue
        label = cells[0] if cells else ""
        current_value = cells[1] if len(cells) > 1 else ""
        if not current_value and label:
            pending_label = " ".join(part for part in [pending_label, label] if part).strip()
            continue
        if not label and not current_value:
            continue
        full_label = " ".join(part for part in [pending_label, label] if part).strip()
        pending_label = ""
        if full_label and current_value:
            rows.append([full_label, current_value])
    return rows


def extract_current_previous_rows_from_text(page_texts: list[str], title: str) -> list[list[str]]:
    for start_page in _base.title_pages(page_texts, title):
        rows: list[list[str]] = []
        started = False
        pending_label = ""
        for text in page_texts[start_page : min(len(page_texts), start_page + 2)]:
            for raw_line in text.splitlines():
                line = _base.normalize_text(raw_line)
                key = _base.normalize_key(line)
                if not started:
                    if _is_title_line(line, title):
                        started = True
                    continue
                if rows and key.startswith("7 8"):
                    return rows
                if (
                    not line
                    or key in {"cari donem onceki donem"}
                    or key.startswith(("cari donem", "onceki donem"))
                ):
                    continue
                parsed = _parse_current_interest_line(line)
                if parsed:
                    label = " ".join(part for part in [pending_label, parsed[0]] if part).strip()
                    rows.append([label, parsed[1]])
                    pending_label = ""
                    continue
                pending_label = " ".join(part for part in [pending_label, line] if part).strip()
        if rows:
            return rows
    return []


def _is_title_line(line: str, title: str) -> bool:
    title_key = _base.normalize_key(title)
    line_key = _base.normalize_key(line)
    title_prefix = " ".join(title_key.split()[:5])
    return title_key in line_key or (title_prefix and title_prefix in line_key)


def _is_strict_title_line(line: str, title: str) -> bool:
    return _base.normalize_key(title) in _base.normalize_key(line)


def _parse_three_amount_line(line: str) -> list[str] | None:
    parts = _base.normalize_text(line).split()
    values: list[str] = []
    while parts and _base.AMOUNT_TOKEN_RE.match(parts[-1]):
        values.insert(0, parts.pop())
    if len(values) < 3 or not parts:
        return None
    return [" ".join(parts), *values[:3]]


def _parse_current_interest_line(line: str) -> list[str] | None:
    parts = _base.normalize_text(line).split()
    values: list[str] = []
    while parts and _base.AMOUNT_TOKEN_RE.match(parts[-1]):
        values.insert(0, parts.pop())
    if len(values) < 2 or not parts:
        return None
    return [" ".join(parts), *values[:2]]


def extract_current_interest_rows(page_texts: list[str], title: str, stop_key: str) -> list[list[str]]:
    for start_page in _base.title_pages(page_texts, title):
        rows: list[list[str]] = []
        started = False
        pending_label = ""
        for text in page_texts[start_page : min(len(page_texts), start_page + 2)]:
            for raw_line in text.splitlines():
                line = _base.normalize_text(raw_line)
                key = _base.normalize_key(line)
                if not started:
                    if _is_title_line(line, title):
                        started = True
                    continue
                if rows and key.startswith(stop_key):
                    return rows
                if (
                    not line
                    or "cari donem" in key
                    or "onceki donem" in key
                    or key == "tp yp tp yp"
                ):
                    continue
                parsed = _parse_current_interest_line(line)
                if parsed:
                    label = " ".join(part for part in [pending_label, parsed[0]] if part).strip()
                    rows.append([label, *parsed[1:3]])
                    pending_label = ""
                    continue
                pending_label = " ".join(part for part in [pending_label, line] if part).strip()
        if rows:
            return rows
    return []


def extract_current_tp_yp_rows(page_texts: list[str], title: str, stop_keys: list[str]) -> list[list[str]]:
    normalized_stops = [_base.normalize_key(stop_key) for stop_key in stop_keys]
    for start_page in _base.title_pages(page_texts, title):
        rows: list[list[str]] = []
        started = False
        pending_label = ""
        for text in page_texts[start_page : min(len(page_texts), start_page + 2)]:
            for raw_line in text.splitlines():
                line = _base.normalize_text(raw_line)
                key = _base.normalize_key(line)
                if not started:
                    if _is_strict_title_line(line, title):
                        started = True
                    continue
                if rows and any(key.startswith(stop_key) for stop_key in normalized_stops):
                    return rows
                if (
                    not line
                    or key in {"cari donem onceki donem", "tp yp tp yp"}
                    or key.startswith(("cari donem", "onceki donem"))
                    or line.startswith("(")
                ):
                    continue
                parsed = _parse_current_interest_line(line)
                if parsed:
                    label = " ".join(part for part in [pending_label, parsed[0]] if part).strip()
                    rows.append([label, *parsed[1:3]])
                    pending_label = ""
                    continue
                pending_label = " ".join(part for part in [pending_label, line] if part).strip()
        if rows:
            return rows
    return []


def extract_expected_loss_rows(page_texts: list[str]) -> list[list[str]]:
    canonical_labels = {
        "beklenenkredizararikarsiliklari": "Beklenen Kredi Zararı Karşılıkları",
        "12aylikbeklenenzararkarsiligibirinciasama": "12 Aylık Beklenen Zarar Karşılığı (Birinci Aşama)",
        "krediriskindeonemliartisikinciasama": "Kredi Riskinde Önemli Artış (İkinci Aşama)",
        "temerrutucuncuasama": "Temerrüt (Üçüncü Aşama)",
        "menkuldegerlerdegerdususkarsiliklari": "Menkul Değerler Değer Düşüş Karşılıkları",
        "gercegeuygundegerfarkikarveyazararayansitilanfinansalvarliklar": "Gerçeğe Uygun Değer Farkı Kâr veya Zarara Yansıtılan Finansal Varlıklar",
        "gercegeuygundegerfarkidigerkapsamligelireyansitilanvarliklar": "Gerçeğe Uygun Değer Farkı Diğer Kapsamlı Gelire Yansıtılan Varlıklar",
        "istiraklerbagliortakliklarvebirliktekontroledilenortakliklardegerdususkarsiliklari": "İştirakler, Bağlı Ortaklıklar ve Birlikte Kontrol Edilen Ortaklıklar Değer Düşüş Karşılıkları",
        "istirakler": "İştirakler",
        "bagliortakliklar": "Bağlı Ortaklıklar",
        "birliktekontroledilenortakliklar": "Birlikte Kontrol Edilen Ortaklıklar",
        "diger": "Diğer",
        "toplam": "Toplam",
    }

    def canonicalize_label(label: str) -> str:
        key = _base.normalize_key(re.sub(r"\(\*\)", "", label))
        compact = re.sub(r"[^a-z0-9çğıöşü]", "", key)
        return canonical_labels.get(compact, _base.normalize_text(label))

    for start_page in _base.title_pages(page_texts, EXPECTED_LOSS_PROVISIONS_TITLE):
        rows: list[list[str]] = []
        started = False
        pending_label = ""
        for text in page_texts[start_page : min(len(page_texts), start_page + 2)]:
            for raw_line in text.splitlines():
                line = _base.normalize_text(raw_line)
                key = _base.normalize_key(line)
                if not started:
                    if _is_title_line(line, EXPECTED_LOSS_PROVISIONS_TITLE):
                        started = True
                    continue
                if rows and key.startswith("6 diger faaliyet"):
                    return rows
                if (
                    not line
                    or key in {"cari donem onceki donem", "cari donem onceki donem (1)"}
                    or key.startswith("cari donem")
                    or key.startswith("onceki donem")
                ):
                    continue
                parsed = _parse_current_interest_line(line)
                if parsed:
                    label = " ".join(part for part in [pending_label, parsed[0]] if part).strip()
                    label = canonicalize_label(label)
                    rows.append([label, *parsed[1:3]])
                    pending_label = ""
                    if label == "Toplam":
                        keys = [_base.normalize_key(row[0]) for row in rows]
                        if len(keys) != len(set(keys)):
                            raise RuntimeError("Ziraat beklenen zarar karsiliklari satirlari tekil degil")
                        return rows
                    continue
                pending_label = " ".join(part for part in [pending_label, line] if part).strip()
        if rows:
            return rows
    return []


def extract_other_operating_income_rows(page_texts: list[str]) -> list[list[str]]:
    for start_page in _base.title_pages(page_texts, OTHER_OPERATING_INCOME_TITLE):
        started = False
        paragraph_lines: list[str] = []
        for text in page_texts[start_page : min(len(page_texts), start_page + 2)]:
            for raw_line in text.splitlines():
                line = _base.normalize_text(raw_line)
                key = _base.normalize_key(line)
                if not started:
                    if _is_strict_title_line(line, OTHER_OPERATING_INCOME_TITLE):
                        started = True
                    continue
                if key.startswith("5 beklenen zarar"):
                    paragraph = " ".join(paragraph_lines)
                    paragraph_key = _base.normalize_key(paragraph)
                    # Ziraat 2026Q2 metninde "kredi" sözcüğü kullanılmadan
                    # "beklenen zarar karşılığı iptallerinden" deniyor.
                    if "beklenen zarar" in paragraph_key and "iptallerinden" in paragraph_key:
                        amount = next((part for part in paragraph.split() if _base.AMOUNT_TOKEN_RE.match(part)), "")
                        if amount:
                            return [["Beklenen kredi zararı karşılığı iptalleri", amount]]
                    return []
                paragraph_lines.append(line)
    return []


def extract_npl_movement_rows(page_texts: list[str]) -> list[list[str]]:
    canonical_labels = {
        "oncekidonemsonubakiyesi": "Önceki Dönem Sonu Bakiyesi",
        "donemicindeintikal": "Dönem İçinde İntikal (+)",
        "digerdonukalacakhesaplarindangiris": "Diğer Donuk Alacak Hesaplarından Giriş (+)",
        "digerdonukalacakhesaplarinacikis": "Diğer Donuk Alacak Hesaplarına Çıkış (-)",
        "donemicindetahsilat": "Dönem İçinde Tahsilat (-) (1)",
        "kayittandusulen": "Kayıttan düşülen (-)",
        "satilan": "Satılan (-)",
        "kurumsalveticarikrediler": "Kurumsal ve Ticari Krediler",
        "bireyselkrediler": "Bireysel Krediler",
        "kredikartlari": "Kredi Kartları",
        "diger": "Diğer",
        "donemsonubakiyesi": "Dönem Sonu Bakiyesi",
        "karsilik": "Karşılık (-)",
        "bilancodakinetbakiyesi": "Bilançodaki Net Bakiyesi",
    }

    def canonicalize_label(label: str) -> str:
        key = _base.normalize_key(re.sub(r"\([12]\)", "", label))
        compact = re.sub(r"[^a-z0-9çğıöşü]", "", key)
        return canonical_labels.get(compact, _base.normalize_text(label))

    for start_page in _base.title_pages(page_texts, NPL_MOVEMENT_TITLE):
        rows: list[list[str]] = []
        started = False
        for text in page_texts[start_page : min(len(page_texts), start_page + 2)]:
            for raw_line in text.splitlines():
                line = _base.normalize_text(raw_line)
                key = _base.normalize_key(line)
                if not started:
                    if _is_title_line(line, NPL_MOVEMENT_TITLE):
                        started = True
                    continue
                if rows and (key.startswith("7 8 3") or line.startswith("(")):
                    return rows
                if (
                    not line
                    or key in {"iii grup iv grup v grup", "zarar"}
                    or key.startswith(("tahsil imkani", "sinirli krediler"))
                ):
                    continue
                parsed = _parse_three_amount_line(line)
                if parsed:
                    parsed[0] = canonicalize_label(parsed[0])
                    rows.append(parsed)
        if rows:
            keys = [_base.normalize_key(row[0]) for row in rows]
            if len(keys) != len(set(keys)):
                raise RuntimeError("Ziraat toplam donuk alacak hareketleri satirlari tekil degil")
            return rows
    return []


def _patch_ziraat_overrides() -> None:
    _base.extract_raw_rows = extract_raw_rows
    _base.extract_stage_provision_rows = extract_stage_provision_rows
    _base.extract_deposit_interest_rows = extract_deposit_interest_rows
    _base.extract_other_operating_expense_rows = extract_other_operating_expense_rows
    _base.extract_other_operating_expense_rows_from_words = extract_other_operating_expense_rows_from_words


_patch_base()
_patch_ziraat_overrides()

for _name in dir(_base):
    if not _name.startswith("__") and _name not in globals():
        globals()[_name] = getattr(_base, _name)


def main() -> None:
    _patch_base()
    _patch_ziraat_overrides()
    _base.main()


if __name__ == "__main__":
    main()
