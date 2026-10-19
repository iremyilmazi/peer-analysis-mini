from __future__ import annotations

import re
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from teb_bankasi.peerpdf_scripts import teb_dipnot as _base

_ORIGINAL_EXTRACT_RAW_ROWS = _base.extract_raw_rows


CONSUMER_CREDIT_TITLE = (
    "Tüketici kredileri, bireysel kredi kartları, personel kredileri "
    "ve personel kredi kartlarına ilişkin bilgiler"
)
INSTALLMENT_COMMERCIAL_TITLE = "Taksitli ticari krediler ve kurumsal kredi kartlarına ilişkin bilgiler"
DEPOSIT_MATURITY_TITLE = "Mevduatın/toplanan fonların vade yapısına ilişkin bilgiler"
DEPOSIT_TOTAL_BY_MATURITY_TITLE = "Mevduatın/toplanan fonların vade yapısına ilişkin toplam bilgileri"
DEPOSIT_INTEREST_TITLE = "Mevduata Ödenen Faizin Vade Yapısına Göre Gösterimi"
RESTRUCTURED_NPL_TITLE = "Donuk alacaklar ve yeniden yapılandırılan kredilere ilişkin bilgiler"
NPL_MOVEMENT_TITLE = "Toplam donuk alacak hareketlerine ilişkin bilgiler"
BANKS_INTEREST_INCOME_TITLE = "Bankalardan alınan faiz gelirlerine ilişkin bilgiler"
SECURITIES_INTEREST_INCOME_TITLE = "Menkul Değerlerden Alınan Faizlere İlişkin Bilgiler"
EQUITY_ITEMS_TITLE = "Özkaynak Kalemlerine İlişkin Bilgiler"
EXPECTED_LOSS_PROVISIONS_TITLE = "Beklenen zarar karşılıkları ve diğer karşılık giderlerine ilişkin açıklamalar:"
STAGE_PROVISION_TITLE = (
    "Standart nitelikli ve yakın izlemedeki krediler ile yeniden yapılandırılan "
    "yakın izlemedeki kredilere ilişkin bilgiler"
)
LIQUIDITY_COVERAGE_RATIO_TITLE = "Likidite karşılama oranı"
STAGE3_PROVISIONS_TITLE = "Kredilere ilişkin olarak ayrılan özel karşılıklar veya temerrüt (Üçüncü Aşama) karşılıkları:"
OTHER_OPERATING_INCOME_TITLE = "Diğer faaliyet gelirlerine ilişkin açıklamalar:"
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
                if key == "onceki donem" or key.startswith(("onceki donem 31", "previous period 31")):
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


def _combine_wrapped_rows(lines: list[str], stop_keys: tuple[str, ...]) -> list[list[str]]:
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
            continue
        if rows and _base.normalize_key(rows[-1][0]).startswith("toplam"):
            break
        pending_label = " ".join(part for part in [pending_label, line] if part).strip()
    return rows


def extract_raw_rows(
    page_texts: list[str],
    page_tables: list[list[list[list[str | None]]]],
    title: str,
    include_empty_amount_rows: bool = False,
) -> list[list[str]]:
    title_key = _base.normalize_key(title)
    if title_key == _base.normalize_key(DEPOSIT_MATURITY_TITLE):
        rows = _current_text_rows(page_texts, DEPOSIT_MATURITY_TITLE, max_pages=2)
        if rows:
            return rows
    if title_key == _base.normalize_key(NPL_MOVEMENT_TITLE):
        rows = _current_text_rows(page_texts, NPL_MOVEMENT_TITLE, max_pages=2)
        if rows:
            return rows
    if title_key in {
        _base.normalize_key(SECURITIES_INTEREST_INCOME_TITLE),
        _base.normalize_key(EXPECTED_LOSS_PROVISIONS_TITLE),
        _base.normalize_key(OTHER_OPERATING_INCOME_TITLE),
    }:
        rows = _combine_wrapped_rows(_lines_after_title(page_texts, title, max_pages=1), ())
        if rows:
            return rows
    return _ORIGINAL_EXTRACT_RAW_ROWS(
        page_texts,
        page_tables,
        title,
        include_empty_amount_rows=include_empty_amount_rows,
    )


def extract_restructured_npl_rows(page_texts: list[str]) -> list[list[str]]:
    rows = _current_text_rows(page_texts, RESTRUCTURED_NPL_TITLE, max_pages=2)
    wanted = []
    for row in rows:
        key = _base.normalize_key(row[0])
        if (
            "karsiliklardan onceki brut tutarlar" in key
            or "yeniden yapilandirilan krediler" in key
            or "yeniden yapilandirilan" in key
        ):
            wanted.append(row)
    return wanted or rows


def extract_stage_provision_rows(page_texts: list[str]) -> list[list[str]]:
    targets = [
        ("12 aylik beklenen zarar karsiligi", "12 Aylik Beklenen Zarar Karsiligi"),
        ("kredi riskinde onemli artis", "Kredi Riskinde Onemli Artis"),
    ]
    found: dict[str, list[str]] = {}
    for start_page in _base.title_pages(page_texts, STAGE_PROVISION_TITLE):
        for text in page_texts[start_page : min(len(page_texts), start_page + 3)]:
            for raw_line in text.splitlines():
                parsed = _parse_amount_line(raw_line)
                if not parsed:
                    continue
                label_key = _base.normalize_key(parsed[0])
                for target_key, label in targets:
                    if target_key in label_key and target_key not in found:
                        if len(parsed) >= 5:
                            found[target_key] = [label, parsed[1], parsed[3]]
                        else:
                            found[target_key] = [label, *parsed[1:3]]
            if len(found) == len(targets):
                return [found[target_key] for target_key, _ in targets]
    return [found[target_key] for target_key, _ in targets if target_key in found]


def extract_liquidity_coverage_rows(page_texts: list[str]) -> list[list[str]]:
    for start_page in _base.title_pages(page_texts, LIQUIDITY_COVERAGE_RATIO_TITLE):
        for text in page_texts[start_page : min(len(page_texts), start_page + 2)]:
            for raw_line in text.splitlines():
                line = _base.normalize_text(raw_line)
                if "likidite karsilama orani" in _base.normalize_key(line):
                    values = re.findall(r"%\s*([0-9]+(?:[.,][0-9]+)?)", line)
                    if len(values) >= 2:
                        return [
                            [
                                "Likidite Karsilama Orani (%)",
                                values[0].replace(",", "."),
                                values[1].replace(",", "."),
                            ]
                        ]
                parsed = _parse_amount_line(raw_line)
                if parsed and "likidite karsilama orani" in _base.normalize_key(parsed[0]) and len(parsed) >= 3:
                    return [["Likidite Karsilama Orani (%)", parsed[1], parsed[2]]]
    return []


def _patch_halkbank_overrides() -> None:
    _base.extract_raw_rows = extract_raw_rows
    _base.extract_restructured_npl_rows = extract_restructured_npl_rows
    _base.extract_stage_provision_rows = extract_stage_provision_rows
    _base.extract_liquidity_coverage_rows = extract_liquidity_coverage_rows


_patch_base()
_patch_halkbank_overrides()

for _name in dir(_base):
    if not _name.startswith("__") and _name not in globals():
        globals()[_name] = getattr(_base, _name)


def main() -> None:
    _patch_base()
    _patch_halkbank_overrides()
    _base.main()


if __name__ == "__main__":
    main()
