from __future__ import annotations

import re
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from teb_bankasi.peerpdf_scripts import teb_dipnot as _base

_ORIGINAL_EXTRACT_RAW_ROWS = _base.extract_raw_rows
_ORIGINAL_TITLE_PAGES = _base.title_pages


CONSUMER_CREDIT_TITLE = (
    "Tüketici kredileri, bireysel kredi kartları ve personel kredi kartlarına ilişkin bilgiler"
)
INSTALLMENT_COMMERCIAL_TITLE = "Taksitli ticari krediler ve kurumsal kredi kartlarına ilişkin bilgiler"
DEPOSIT_MATURITY_TITLE = "Mevduatın vade yapısına ilişkin bilgiler"
DEPOSIT_TOTAL_BY_MATURITY_TITLE = "Mevduatın vade yapısına ilişkin toplam bilgileri"
DEPOSIT_INTEREST_TITLE = "Mevduata ödenen faizin vade yapısına göre gösterimi"
RESTRUCTURED_NPL_TITLE = (
    "Donuk alacaklara ve Banka tarafından yeniden yapılandırılan kredilere ilişkin bilgiler"
)
NPL_MOVEMENT_TITLE = "Toplam donuk alacak hareketlerine ilişkin bilgiler"
BANKS_INTEREST_INCOME_TITLE = "Bankalardan alınan faiz gelirlerine ilişkin bilgiler"
SECURITIES_INTEREST_INCOME_TITLE = "Menkul değerlerden alınan faizlere ilişkin bilgiler"
EQUITY_ITEMS_TITLE = "Özkaynak kalemlerine ilişkin bilgiler"
EXPECTED_LOSS_PROVISIONS_TITLE = "Beklenen kredi zararı karşılıkları ve diğer karşılıklar"
EXPECTED_LOSS_PROVISIONS_ALIASES = (
    EXPECTED_LOSS_PROVISIONS_TITLE,
    "Beklenen zarar karşılıkları ve diğer karşılık giderleri",
    "Beklenen zarar karşılıkları ve diğer karşılık giderlerine ilişkin açıklamalar",
    "Beklenen kredi zararı karşılıkları",
)
STAGE_PROVISION_TITLE = (
    "Standart nitelikli ve yakın izlemedeki krediler ile yeniden yapılandırılan "
    "yakın izlemedeki kredilere ilişkin bilgiler"
)
STAGE_PROVISION_HEADERS = ["Kalem", "Standart Nitelikli Krediler", "Yakın İzlemedeki Krediler"]
LIQUIDITY_COVERAGE_RATIO_TITLE = "LİKİDİTE KARŞILAMA ORANI"
STAGE3_PROVISIONS_TITLE = "Kredilere ilişkin olarak ayrılan özel karşılıklar veya temerrüt (Üçüncü Aşama) karşılıklar"
OTHER_OPERATING_INCOME_TITLE = "Diğer faaliyet gelirlerine ilişkin açıklamalar"
CASH_AND_TCMB_TITLE = "Nakit değerler ve T.C. Merkez Bankası’na ilişkin bilgiler"
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
    _base.STAGE_PROVISION_HEADERS = STAGE_PROVISION_HEADERS
    _base.STAGE_PROVISION_TARGET_LABELS = [
        "12 aylık beklenen zarar karşılığı",
        "Kredi riskinde önemli artış",
    ]


def _is_title_line(line: str, title: str) -> bool:
    if _base.normalize_key(title) == _base.normalize_key(EXPECTED_LOSS_PROVISIONS_TITLE):
        return any(
            _base.normalize_key(alias) in _base.normalize_key(line)
            or " ".join(_base.normalize_key(alias).split()[:5]) in _base.normalize_key(line)
            for alias in EXPECTED_LOSS_PROVISIONS_ALIASES
        )
    title_key = _base.normalize_key(title)
    line_key = _base.normalize_key(line)
    title_prefix = " ".join(title_key.split()[:5])
    return title_key in line_key or (title_prefix and title_prefix in line_key)


def title_pages(page_texts: list[str], title: str) -> list[int]:
    if _base.normalize_key(title) != _base.normalize_key(EXPECTED_LOSS_PROVISIONS_TITLE):
        return _ORIGINAL_TITLE_PAGES(page_texts, title)
    pages: set[int] = set()
    for alias in EXPECTED_LOSS_PROVISIONS_ALIASES:
        pages.update(_ORIGINAL_TITLE_PAGES(page_texts, alias))
    detailed_pages = [
        page
        for page in sorted(pages)
        if "12 aylik beklenen zarar" in _base.normalize_key(page_texts[page])
        and "temerrut" in _base.normalize_key(page_texts[page])
    ]
    return detailed_pages or sorted(pages)


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
    parts = _base.normalize_text(line).replace("--", "-").split()
    values: list[str] = []
    while parts and _base.SIGNED_AMOUNT_TOKEN_RE.match(parts[-1]):
        if re.fullmatch(r"\(\d+\)", parts[-1]):
            break
        values.insert(0, parts.pop())
    if len(values) < 2 or not parts:
        return None
    return [" ".join(parts), *values]


def _parse_variable_amount_line(line: str, min_values: int = 1) -> list[str] | None:
    parts = _base.normalize_text(line).replace("--", "-").split()
    values: list[str] = []
    while parts and _base.SIGNED_AMOUNT_TOKEN_RE.match(parts[-1]):
        if re.fullmatch(r"\(\d+\)", parts[-1]):
            break
        values.insert(0, parts.pop())
    if len(values) < min_values or not parts:
        return None
    return [" ".join(parts), *values]


def _is_amount_only_line(line: str) -> bool:
    parts = _base.normalize_text(line).replace("--", "-").split()
    return bool(parts) and all(_base.SIGNED_AMOUNT_TOKEN_RE.match(part) for part in parts)


def _clean_label(label: str, title: str | None = None) -> str:
    cleaned = _base.normalize_text(label)
    if title is not None:
        title_key = _base.normalize_key(title)
        cleaned_key = _base.normalize_key(cleaned)
        if title_key in cleaned_key and " bilgiler " in cleaned:
            cleaned = cleaned.split(" bilgiler ", 1)[1].strip()
    while ":" in cleaned:
        prefix, tail = cleaned.split(":", 1)
        prefix_key = _base.normalize_key(prefix)
        if (title is not None and _base.normalize_key(title) in prefix_key) or "iliskin bilgiler" in prefix_key:
            cleaned = tail.strip()
            continue
        break
    cleaned = re.sub(r"\s*:?\s*31\s+(?:Mart|Aralık)\s+20\d{2}\s*$", "", cleaned).strip()
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


def _cash_and_tcmb_rows(page_texts: list[str]) -> list[list[str]]:
    rows: list[list[str]] = []
    for line in _lines_after_title(page_texts, CASH_AND_TCMB_TITLE, max_pages=1):
        key = _base.normalize_key(line)
        if key.startswith("2 t c merkez bankasi hesabina iliskin bilgiler"):
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
        if not parsed or len(parsed) < 3:
            continue
        label = _clean_label(parsed[0], CASH_AND_TCMB_TITLE)
        if _base.normalize_key(label) == "diger":
            label = "Diğer"
        rows.append([label, parsed[1], parsed[2]])
        if _base.normalize_key(label).startswith("toplam"):
            break
    return rows


def _expected_loss_rows(page_texts: list[str]) -> list[list[str]]:
    rows: list[list[str]] = []
    pending_label = ""
    pending_amounts: list[str] | None = None
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

        parts = line.replace("--", "-").split()
        line_values: list[str] = []
        while parts and _base.SIGNED_AMOUNT_TOKEN_RE.match(parts[-1]):
            line_values.insert(0, parts.pop())
        label_part = " ".join(parts).strip()

        if pending_amounts is not None:
            if label_part and not line_values:
                rows.append([_clean_label(f"{pending_label} {label_part}".strip()), *pending_amounts])
                pending_label = ""
                pending_amounts = None
                continue
            rows.append([_clean_label(pending_label), *pending_amounts])
            pending_label = ""
            pending_amounts = None

        if line_values:
            if _base.normalize_key(label_part).startswith("toplam"):
                pending_label = ""
            label = _clean_label(" ".join(part for part in [pending_label, label_part] if part).strip())
            if not label and pending_label:
                label = _clean_label(pending_label)
            if label and _base.normalize_key(label).endswith("deger dusus") and len(line_values) >= 2:
                pending_label = label
                pending_amounts = line_values
                continue
            if label:
                rows.append([label, *line_values])
                pending_label = ""
                if _base.normalize_key(label).startswith("toplam"):
                    break
            continue

        pending_label = " ".join(part for part in [pending_label, label_part or line] if part).strip()

    if pending_amounts is not None:
        rows.append([_clean_label(pending_label), *pending_amounts])
    return rows


def _npl_movement_rows(page_texts: list[str]) -> list[list[str]]:
    rows: list[list[str]] = []
    for line in _lines_after_title(page_texts, NPL_MOVEMENT_TITLE, max_pages=1, include_title_tail=True):
        key = _base.normalize_key(line)
        if rows and key.startswith(("iii yabanci", "iii donuk", "iv donuk", "banka yonetim")):
            break
        if (
            not line
            or key in {"iii grup iv grup v grup"}
            or key.startswith(("tahsil imkani", "sinirli krediler", "krediler krediler"))
            or line.startswith("Yetkili ")
        ):
            continue
        parsed = _parse_amount_line(line)
        if parsed:
            rows.append([_clean_label(parsed[0], NPL_MOVEMENT_TITLE), *parsed[1:]])
    return rows


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


def _ensure_fx_deposit_total(rows: list[list[str]]) -> list[list[str]]:
    """Add DenizBank's omitted FX-deposit total from its two resident rows.

    Some report layouts (notably 2024 Q1 and Q2) print only the domestic and
    foreign resident components under the FX-deposit heading.  Other periods
    also print their aggregate as ``Döviz Tevdiat Hesabı (*)``.  Keeping that
    aggregate in the normalized schema prevents a legitimate layout change
    from becoming three blank workbook cells.
    """
    if any(
        _base.normalize_key(row[0]).startswith("doviz tevdiat hesabi")
        for row in rows
        if row
    ):
        return rows

    components: dict[str, list[str]] = {}
    for row in rows:
        if not row:
            continue
        key = _base.normalize_key(row[0])
        if key.startswith("yurt icinde yer"):
            components["domestic"] = row
        elif key.startswith("yurt disinda yer"):
            components["foreign"] = row

    if set(components) != {"domestic", "foreign"}:
        return rows

    domestic = components["domestic"]
    foreign = components["foreign"]
    width = max(len(domestic), len(foreign), 4)
    aggregate = ["Döviz Tevdiat Hesabı (*)"]
    for column in range(1, width):
        values = [
            row[column] if column < len(row) else "-"
            for row in (domestic, foreign)
        ]
        aggregate.append(_base.format_amount(_base.sum_amounts(values)))

    insert_at = next(
        (
            index + 1
            for index, row in enumerate(rows)
            if row and _base.normalize_key(row[0]).startswith("tasarruf mevduati")
        ),
        0,
    )
    return [*rows[:insert_at], aggregate, *rows[insert_at:]]


def extract_raw_rows(
    page_texts: list[str],
    page_tables: list[list[list[list[str | None]]]],
    title: str,
    include_empty_amount_rows: bool = False,
) -> list[list[str]]:
    title_key = _base.normalize_key(title)
    if title_key == _base.normalize_key(CONSUMER_CREDIT_TITLE):
        rows = _plain_current_rows(page_texts, title, ("1 7 4",), max_pages=1)
        if rows:
            return rows
    if title_key == _base.normalize_key(INSTALLMENT_COMMERCIAL_TITLE):
        rows = _plain_current_rows(page_texts, title, ("yetkili",), max_pages=1)
        if rows:
            return rows
    if title_key == _base.normalize_key(DEPOSIT_MATURITY_TITLE):
        rows = _plain_current_rows(page_texts, title, ("onceki donem",), max_pages=1)
        if rows:
            return _ensure_fx_deposit_total(rows)
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
    if title_key == _base.normalize_key(CASH_AND_TCMB_TITLE):
        rows = _cash_and_tcmb_rows(page_texts)
        if rows:
            return rows
    if title_key == _base.normalize_key(BANKS_INTEREST_INCOME_TITLE):
        rows = _combine_wrapped_rows(_lines_after_title(page_texts, title, max_pages=1, include_title_tail=True))
        if rows:
            for row in rows:
                row[0] = _clean_label(row[0], BANKS_INTEREST_INCOME_TITLE)
            return rows
    if title_key == _base.normalize_key(LOAN_INTEREST_INCOME_TITLE):
        rows = _combine_wrapped_rows(_lines_after_title(page_texts, title, max_pages=1, include_title_tail=True))
        if rows:
            for row in rows:
                row[0] = _clean_label(row[0], LOAN_INTEREST_INCOME_TITLE)
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
    rows: list[list[str]] = []
    pending_values: list[str] | None = None
    for line in _lines_after_title(page_texts, RESTRUCTURED_NPL_TITLE, max_pages=1):
        key = _base.normalize_key(line)
        if "onceki donem" in key:
            break
        parsed = _parse_amount_line(line)
        if parsed:
            if key.startswith("cari donem"):
                pending_values = parsed[1:]
                continue
            rows.append([_clean_label(parsed[0], RESTRUCTURED_NPL_TITLE), *parsed[1:]])
            continue
        if pending_values and "karsiliklardan onceki brut tutarlar" in key:
            rows.append(["Karşılıklardan Önceki Brüt Tutarlar", *pending_values])
            pending_values = None
    return rows


def extract_stage_provision_rows(page_texts: list[str]) -> list[list[str]]:
    output: list[list[str]] = []
    for start_page in _base.title_pages(page_texts, STAGE_PROVISION_TITLE):
        in_current = False
        for raw_line in page_texts[start_page].splitlines():
            key_line = _base.normalize_key(raw_line)
            if "cari donem" in key_line:
                in_current = True
                continue
            if "onceki donem" in key_line and output:
                break
            if not in_current:
                continue
            parsed = _parse_amount_line(raw_line)
            if not parsed:
                continue
            key = _base.normalize_key(parsed[0])
            if "12 aylik beklenen zarar karsiligi" in key:
                output.append(["12 aylık beklenen zarar karşılığı", parsed[1], parsed[2] if len(parsed) > 2 else ""])
            elif "kredi riskinde onemli artis" in key:
                output.append(["Kredi riskinde önemli artış", parsed[1], parsed[2] if len(parsed) > 2 else ""])
        if output:
            break
    return output


def extract_deposit_interest_rows(page_texts: list[str]) -> list[list[str]]:
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
        parsed = _parse_amount_line(line)
        if not parsed:
            continue
        label = _clean_label(parsed[0], DEPOSIT_INTEREST_TITLE)
        if section and not _base.normalize_key(label).startswith(("toplam", "genel toplam")):
            label = f"{section} - {label}"
        values = parsed[1:]
        if len(values) >= 9:
            values = values[:8]
        rows.append([label, *values])
        if _base.normalize_key(label) == "genel toplam":
            break
    return rows


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


def extract_securities_interest_rows(page_texts: list[str]) -> list[list[str]]:
    return _rows_with_amount_continuation(page_texts, SECURITIES_INTEREST_INCOME_TITLE, ("a 4 istirak",))


def extract_used_loans_interest_rows(page_texts: list[str]) -> list[list[str]]:
    rows: list[list[str]] = []
    blank_labels = {
        "t c merkez bankasina",
        "yurt disi merkez ve subelere",
    }
    for line in _lines_after_title(page_texts, USED_LOANS_INTEREST_EXPENSE_TITLE, max_pages=1):
        key = _base.normalize_key(line)
        if key.startswith("b 2 istirak") or "istirak ve bagli ortakliklara verilen faizler" in key:
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
            rows.append([_clean_label(parsed[0], USED_LOANS_INTEREST_EXPENSE_TITLE), *parsed[1:]])
        elif key in blank_labels:
            rows.append([line, "", ""])
    return rows


def extract_equity_rows(page_texts: list[str]) -> list[list[str]]:
    targets = {
        "cekirdek sermaye toplami": ("Çekirdek Sermaye Toplamı", None),
        "katki sermaye toplami": ("Katkı Sermaye Toplamı", None),
        "toplam risk agirlikli tutarlar": ("Toplam Risk Ağırlıklı Tutarlar", None),
    }
    total_equity_label = "Toplam Özkaynak (Ana sermaye ve katkı sermaye toplamı)"
    total_equity_value: str | None = None
    for start_page in _base.title_pages(page_texts, EQUITY_ITEMS_TITLE):
        found_any = False
        for text in page_texts[start_page : min(len(page_texts), start_page + 4)]:
            previous_parsed: list[str] | None = None
            for raw_line in text.splitlines():
                line = _base.normalize_text(raw_line)
                key = _base.normalize_key(line)
                parsed = _parse_amount_line(line)
                if not parsed:
                    continue
                if "toplam risk agirlikli tutarlar" in key:
                    if (
                        previous_parsed
                        and "toplam ozkaynak ana sermaye ve katki sermaye toplami"
                        in _base.normalize_key(previous_parsed[0])
                    ):
                        total_equity_value = previous_parsed[1]
                        found_any = True
                for target_key, (label, value) in list(targets.items()):
                    if value is None and target_key in key:
                        targets[target_key] = (label, parsed[1])
                        found_any = True
                previous_parsed = parsed
        if found_any and total_equity_value is not None and all(value is not None for _, value in targets.values()):
            break

    rows: list[list[str]] = []
    missing: list[str] = []
    ordered_targets = [
        targets["cekirdek sermaye toplami"],
        targets["katki sermaye toplami"],
        (total_equity_label, total_equity_value),
        targets["toplam risk agirlikli tutarlar"],
    ]
    for label, value in ordered_targets:
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
                if not key.startswith("23 likidite karsilama orani"):
                    continue
                values = [token for token in line.split() if re.fullmatch(r"[0-9]+(?:\.[0-9]{3})*,[0-9]+", token)]
                if len(values) >= 2:
                    return [["Likidite Karşılama Oranı (%)", values[0], values[1]]]
    return []


def to_transposed_total(rows: list[list[str]], fallback_headers: list[str]) -> list[list[str]]:
    total_row = next((row for row in rows if row and _base.normalize_key(row[0]).startswith(("toplam", "genel toplam"))), rows[-1])
    headers = [
        "Vadesiz",
        "7 Gün İhbarlı",
        "1 Aya Kadar",
        "1-3 Ay",
        "3-6 Ay",
        "6 Ay-1 Yıl",
        "1 Yıl ve Üstü",
        "Birikimli Mevduat",
        "Toplam",
    ][
        : len(total_row) - 1
    ]
    return [[header, "", "", _base.clean_cell(value)] for header, value in zip(headers, total_row[1:])]


def to_transposed_general_total(rows: list[list[str]], fallback_headers: list[str]) -> list[list[str]]:
    total_row = next((row for row in rows if row and _base.normalize_key(row[0]) == "genel toplam"), None)
    if total_row is None:
        total_row = next((row for row in reversed(rows) if row and _base.normalize_key(row[0]).endswith("toplam")), rows[-1])
    headers = ["Vadesiz Mevduat", "1 Aya Kadar", "3 Aya Kadar", "6 Aya Kadar", "1 Yıla Kadar", "1 Yıldan Uzun", "Birikimli Mevduat", "Toplam"][
        : len(total_row) - 1
    ]
    return [[header, "", "", _base.clean_cell(value)] for header, value in zip(headers, total_row[1:])]


def _patch_qnb_overrides() -> None:
    _base.title_pages = title_pages
    _base.extract_raw_rows = extract_raw_rows
    _base.extract_restructured_npl_rows = extract_restructured_npl_rows
    _base.extract_stage_provision_rows = extract_stage_provision_rows
    _base.extract_deposit_interest_rows = extract_deposit_interest_rows
    _base.extract_equity_rows = extract_equity_rows
    _base.extract_liquidity_coverage_rows = extract_liquidity_coverage_rows
    _base.to_transposed_total = to_transposed_total
    _base.to_transposed_general_total = to_transposed_general_total


_patch_base()
_patch_qnb_overrides()

for _name in dir(_base):
    if not _name.startswith("__") and _name not in globals():
        globals()[_name] = getattr(_base, _name)


def main() -> None:
    _patch_base()
    _patch_qnb_overrides()
    _base.main()


if __name__ == "__main__":
    main()
