from __future__ import annotations

import re
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from teb_bankasi.peerpdf_scripts import teb_dipnot as _base

_ORIGINAL_EXTRACT_RAW_ROWS = _base.extract_raw_rows
_ORIGINAL_EXTRACT_DEPOSIT_INTEREST_ROWS = _base.extract_deposit_interest_rows
_ORIGINAL_EXTRACT_OTHER_EXPENSE_WORDS = _base.extract_other_operating_expense_rows_from_words


CONSUMER_CREDIT_TITLE = (
    "Tüketici kredileri, bireysel kredi kartları, personel kredileri "
    "ve personel kredi kartlarına ilişkin bilgiler"
)
INSTALLMENT_COMMERCIAL_TITLE = "Taksitli ticari krediler ve kurumsal kredi kartlarına ilişkin bilgiler"
DEPOSIT_MATURITY_TITLE = "Mevduatın vade yapısına ilişkin bilgiler"
DEPOSIT_TOTAL_BY_MATURITY_TITLE = "Mevduatın vade yapısına ilişkin toplam bilgileri"
DEPOSIT_INTEREST_TITLE = "Mevduata Ödenen Faizin Vade Yapısına Göre Gösterimi"
RESTRUCTURED_NPL_TITLE = (
    "Donuk alacaklar için hesaplanan faiz tahakkukları, reeskontları ve değerleme "
    "farkları ile bunların karşılıklarına ilişkin bilgiler"
)
NPL_MOVEMENT_TITLE = "Toplam donuk alacak hareketlerine ilişkin bilgiler"
BANKS_INTEREST_INCOME_TITLE = "Bankalardan alınan faiz gelirlerine ilişkin bilgiler"
SECURITIES_INTEREST_INCOME_TITLE = "Menkul değerlerden alınan faizlere ilişkin bilgiler"
EQUITY_ITEMS_TITLE = "Özkaynak kalemlerine ilişkin bilgiler"
EXPECTED_LOSS_PROVISIONS_TITLE = "Beklenen zarar karşılıkları ve diğer karşılık giderleri"
STAGE_PROVISION_TITLE = (
    "Standart nitelikli ve yakın izlemedeki krediler ile yeniden yapılandırılan "
    "yakın izlemedeki kredilere ilişkin bilgiler"
)
LIQUIDITY_COVERAGE_RATIO_TITLE = "LİKİDİTE KARŞILAMA ORANI"
STAGE3_PROVISIONS_TITLE = "Kredilere ilişkin olarak ayrılan temerrüt (Üçüncü Aşama) karşılıkları"
OTHER_OPERATING_INCOME_TITLE = "Diğer faaliyet gelirlerine ilişkin açıklamalar"
CASH_AND_TCMB_TITLE = "Nakit değerler ve T.C. Merkez Bankası hesabına ilişkin bilgiler"
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
    title_prefix = " ".join(title_key.split()[:5])
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


def _plain_rows_after_title(
    page_texts: list[str],
    title: str,
    stop_keys: tuple[str, ...],
    max_pages: int = 1,
    require_current: bool = True,
) -> list[list[str]]:
    rows: list[list[str]] = []
    in_current = not require_current
    for line in _lines_after_title(page_texts, title, max_pages=max_pages):
        key = _base.normalize_key(line)
        if rows and any(key.startswith(stop_key) for stop_key in stop_keys):
            break
        if "onceki donem" in key or "previous period" in key:
            if rows:
                break
            in_current = False
            continue
        if "cari donem" in key or "current period" in key:
            in_current = True
            continue
        if not in_current:
            continue
        if (
            not line
            or key.startswith("31 mart")
            or key in {
                "kisa vadeli orta ve uzun vadeli toplam",
                "iii grup iv grup v grup",
                "standart nitelikli krediler yakin izlemedeki krediler",
            }
            or key.startswith(("tahsil imkani tahsili supheli", "krediler icin krediler icin"))
        ):
            continue
        parsed = _parse_amount_line(line)
        if parsed:
            rows.append(parsed)
    return rows


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


def _npl_movement_rows(page_texts: list[str]) -> list[list[str]]:
    rows: list[list[str]] = []
    for line in _lines_after_title(page_texts, NPL_MOVEMENT_TITLE, max_pages=1):
        key = _base.normalize_key(line)
        if rows and key.startswith("j 3"):
            break
        if (
            not line
            or key in {"iii grup iv grup v grup"}
            or key.startswith(("tahsil imkani tahsili supheli", "krediler icin krediler icin", "ayrilanlar ayrilanlar"))
        ):
            continue
        parsed = _parse_amount_line(line)
        if parsed:
            rows.append(parsed)
    return rows


def extract_raw_rows(
    page_texts: list[str],
    page_tables: list[list[list[list[str | None]]]],
    title: str,
    include_empty_amount_rows: bool = False,
) -> list[list[str]]:
    title_key = _base.normalize_key(title)
    if title_key == _base.normalize_key(CONSUMER_CREDIT_TITLE):
        rows = _plain_rows_after_title(page_texts, title, (), max_pages=1)
        if rows:
            return rows
    if title_key == _base.normalize_key(INSTALLMENT_COMMERCIAL_TITLE):
        rows = _plain_rows_after_title(page_texts, title, ("onceki donem",), max_pages=1)
        if rows:
            return rows
    if title_key == _base.normalize_key(DEPOSIT_MATURITY_TITLE):
        rows = _plain_rows_after_title(page_texts, title, ("onceki donem", "1 1 mevduat"), max_pages=1)
        if rows:
            return rows
    if title_key == _base.normalize_key(NPL_MOVEMENT_TITLE):
        rows = _npl_movement_rows(page_texts)
        if rows:
            return rows
    if title_key == _base.normalize_key(STAGE3_PROVISIONS_TITLE):
        rows = _plain_rows_after_title(page_texts, title, ("j donuk",), max_pages=1, require_current=True)
        if rows:
            return rows
    if title_key == _base.normalize_key(LOAN_INTEREST_INCOME_TITLE):
        rows = _combine_wrapped_rows(
            _lines_after_title(page_texts, title, max_pages=1),
            stop_keys=("b bankalardan",),
        )
        if rows:
            return rows
    if title_key in {
        _base.normalize_key(SECURITIES_INTEREST_INCOME_TITLE),
        _base.normalize_key(EXPECTED_LOSS_PROVISIONS_TITLE),
    }:
        rows = _combine_wrapped_rows(_lines_after_title(page_texts, title, max_pages=1))
        if rows:
            return rows
    if title_key == _base.normalize_key(OTHER_OPERATING_INCOME_TITLE):
        lines = _lines_after_title(page_texts, title, max_pages=1)
        text_lines = []
        for line in lines:
            if _base.normalize_key(line).startswith("6 beklenen zarar"):
                break
            if line and not line.startswith("ENPARA "):
                text_lines.append(line)
        if text_lines:
            return [[" ".join(text_lines), "", "", ""]]
    rows = _ORIGINAL_EXTRACT_RAW_ROWS(
        page_texts,
        page_tables,
        title,
        include_empty_amount_rows=include_empty_amount_rows,
    )
    if title_key == _base.normalize_key(OTHER_OPERATING_EXPENSE_TITLE):
        repaired: list[list[str]] = []
        index = 0
        while index < len(rows):
            row = rows[index]
            key = _base.normalize_key(row[0]) if row else ""
            if (
                "arastirma ve gelistirme giderleri kredi kartlari" in key
                and index + 1 < len(rows)
                and _base.normalize_key(rows[index + 1][0]) == "giderleri"
            ):
                repaired.append(["Araştırma ve Geliştirme Giderleri", *row[1:]])
                repaired.append(
                    [
                        "Kredi Kartları ve Bankacılık Hizmetlerine İlişkin Promosyon Uygulamaları Giderleri",
                        *rows[index + 1][1:],
                    ]
                )
                index += 2
                continue
            repaired.append(row)
            index += 1
        rows = repaired
    return rows


def extract_restructured_npl_rows(page_texts: list[str]) -> list[list[str]]:
    rows = _plain_rows_after_title(page_texts, RESTRUCTURED_NPL_TITLE, ("onceki donem", "k zarar"), max_pages=1)
    wanted = []
    for row in rows:
        key = _base.normalize_key(row[0])
        if "faiz tahakkuk" in key or "karsilik tutari" in key:
            wanted.append(row)
    return wanted or rows


def extract_enpara_other_operating_expense_rows(page_texts, page_words):
    rows = _ORIGINAL_EXTRACT_OTHER_EXPENSE_WORDS(page_texts, page_words)
    repaired: list[list[str]] = []
    index = 0
    while index < len(rows):
        row = rows[index]
        key = _base.normalize_key(row[0]) if row else ""
        if (
            "arastirma ve gelistirme giderleri kredi kartlari" in key
            and index + 1 < len(rows)
            and _base.normalize_key(rows[index + 1][0]) == "giderleri"
        ):
            repaired.append(["Araştırma ve Geliştirme Giderleri", *row[1:]])
            repaired.append(
                [
                    "Kredi Kartları ve Bankacılık Hizmetlerine İlişkin Promosyon Uygulamaları Giderleri",
                    *rows[index + 1][1:],
                ]
            )
            index += 2
            continue
        repaired.append(row)
        index += 1
    return repaired


def extract_stage_provision_rows(page_texts: list[str]) -> list[list[str]]:
    targets = [
        ("12 aylik beklenen zarar karsiligi", "12 Aylik Beklenen Zarar Karsiligi"),
        ("kredi riskinde onemli artis", "Kredi Riskinde Onemli Artis"),
    ]
    found: dict[str, list[str]] = {}
    for start_page in _base.title_pages(page_texts, STAGE_PROVISION_TITLE):
        text = page_texts[start_page]
        if "Cari Dönem" not in text:
            continue
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


def extract_deposit_interest_rows(page_texts: list[str]) -> list[list[str]]:
    rows = _ORIGINAL_EXTRACT_DEPOSIT_INTEREST_ROWS(page_texts)
    output = []
    for row in rows:
        if row and _base.normalize_key(row[0]) == "turk parasi 7 gun ihbarli mevduat" and len(row) < 9:
            row = [row[0], "-", "-", "-", "-", "-", "-", "-", "-"]
        output.append(row)
    return output


def extract_liquidity_coverage_rows(page_texts: list[str]) -> list[list[str]]:
    for start_page in _base.title_pages(page_texts, LIQUIDITY_COVERAGE_RATIO_TITLE):
        for text in page_texts[start_page : min(len(page_texts), start_page + 3)]:
            for raw_line in text.splitlines():
                line = _base.normalize_text(raw_line)
                key = _base.normalize_key(line)
                if "likidite karsilama orani" not in key or "oran" not in key:
                    continue
                values = [
                    token
                    for token in line.split()
                    if re.fullmatch(r"[0-9]+(?:\.[0-9]{3})*,[0-9]+", token)
                ]
                if len(values) >= 2:
                    return [
                        [
                            "Likidite Karsilama Orani (%)",
                            values[-2].replace(".", "").replace(",", "."),
                            values[-1].replace(".", "").replace(",", "."),
                        ]
                    ]
    return []


def _patch_enpara_overrides() -> None:
    _base.extract_raw_rows = extract_raw_rows
    _base.extract_restructured_npl_rows = extract_restructured_npl_rows
    _base.extract_stage_provision_rows = extract_stage_provision_rows
    _base.extract_deposit_interest_rows = extract_deposit_interest_rows
    _base.extract_liquidity_coverage_rows = extract_liquidity_coverage_rows
    _base.extract_other_operating_expense_rows_from_words = extract_enpara_other_operating_expense_rows


_patch_base()
_patch_enpara_overrides()

for _name in dir(_base):
    if not _name.startswith("__") and _name not in globals():
        globals()[_name] = getattr(_base, _name)


def main() -> None:
    _patch_base()
    _patch_enpara_overrides()
    _base.main()


if __name__ == "__main__":
    main()
