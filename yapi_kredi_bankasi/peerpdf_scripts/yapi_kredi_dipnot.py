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
DEPOSIT_INTEREST_TITLE = "Mevduata ödenen faizin vade yapısına göre gösterimi"
RESTRUCTURED_NPL_TITLE = "Donuk alacaklardan yeniden yapılandırılan kredilere ilişkin bilgiler"
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
STAGE3_PROVISIONS_TITLE = "Ayrılan temerrüt (üçüncü aşama) karşılıklarına ilişkin bilgiler"
OTHER_OPERATING_INCOME_TITLE = "Diğer faaliyet gelirlerine ilişkin açıklamalar"
CASH_AND_TCMB_TITLE = "Nakit değerler ve TCMB hesabına ilişkin bilgiler"
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
    title_key = _base.normalize_key(title)
    line_key = _base.normalize_key(line)
    title_prefix = " ".join(title_key.split()[:5])
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


def _expected_loss_rows(page_texts: list[str]) -> list[list[str]]:
    rows: list[list[str]] = []
    pending_label = ""
    pending_amounts: list[str] | None = None
    for line in _lines_after_title(page_texts, EXPECTED_LOSS_PROVISIONS_TITLE, max_pages=1):
        key = _base.normalize_key(line)
        if key.startswith(("4 5 diger", "4 6 diger", "diger faaliyet")):
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
    parent_label = (
        "İştirakler, Bağlı Ortaklıklar ve Birlikte Kontrol Edilen "
        "Ortaklıklar Değer Düşüş Karşılıkları"
    )
    parent_prefix = _base.normalize_key(
        "İştirakler, Bağlı Ortaklıklar ve Birlikte Kontrol Edilen "
        "Ortaklıklar Değer Düşüş"
    )
    for row in rows:
        label_key = _base.normalize_key(row[0])
        if label_key.startswith(parent_prefix):
            # Some 2020 PDFs wrap "Karşılıkları" and the following
            # "İştirakler" sub-row onto the same extracted text line. Keep
            # the parent row canonical and let the normal sub-row retain its
            # own position instead of creating "Karşılıkları İştirakler".
            row[0] = parent_label
    return rows


def _npl_movement_rows(page_texts: list[str]) -> list[list[str]]:
    lines: list[str] = []
    for line in _lines_after_title(page_texts, NPL_MOVEMENT_TITLE, max_pages=1, include_title_tail=True):
        key = _base.normalize_key(line)
        if key.startswith(("1 7 8 3", "yabanci para", "banka yonetim")):
            break
        if (
            not line
            or key in {"iii grup iv grup v grup"}
            or key.startswith(("tahsil imkani", "sinirli krediler", "krediler krediler"))
            or line.startswith("Yetkili ")
        ):
            continue
        lines.append(line)
    return _combine_wrapped_rows(lines, skip_period_markers=False)


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
    if title_key in {
        _base.normalize_key(BANKS_INTEREST_INCOME_TITLE),
        _base.normalize_key(SECURITIES_INTEREST_INCOME_TITLE),
        _base.normalize_key(CASH_AND_TCMB_TITLE),
        _base.normalize_key(USED_LOANS_INTEREST_EXPENSE_TITLE),
        _base.normalize_key(LOAN_INTEREST_INCOME_TITLE),
    }:
        rows = _combine_wrapped_rows(_lines_after_title(page_texts, title, max_pages=1, include_title_tail=True))
        if rows:
            return rows
    return _ORIGINAL_EXTRACT_RAW_ROWS(
        page_texts,
        page_tables,
        title,
        include_empty_amount_rows=include_empty_amount_rows,
    )


def extract_restructured_npl_rows(page_texts: list[str]) -> list[list[str]]:
    return _plain_current_rows(page_texts, RESTRUCTURED_NPL_TITLE, ("onceki donem", "1 7 8 2"), max_pages=1)


def extract_stage_provision_rows(page_texts: list[str]) -> list[list[str]]:
    # Annual reports can place the two-row provision table on the page after
    # the disclosure title, so scan every page and match the exact row labels.
    candidate_pages = list(range(len(page_texts)))
    for start_page in candidate_pages:
        output: list[list[str]] = []
        for raw_line in page_texts[start_page].splitlines():
            parsed = _parse_amount_line(raw_line)
            # Yapı Kredi'nin bazı 2023 PDF'lerinde tablo başlığı metin
            # katmanında ayrışıyor. Sayfada iki kredi sınıfı başlığını zorunlu
            # tutmak, aynı ifadeleri içeren karşılık gideri tablosunu dışlar.
            if not parsed:
                continue
            key = _base.normalize_key(parsed[0])
            if key == "12 aylik beklenen zarar karsiligi":
                output.append(["12 aylık beklenen zarar karşılığı", parsed[1], parsed[2] if len(parsed) > 2 else ""])
            elif key == "kredi riskinde onemli artis":
                output.append(["Kredi riskinde önemli artış", parsed[1], parsed[2] if len(parsed) > 2 else ""])
        if len(output) == 2:
            return output
    return []


def extract_deposit_interest_rows(page_texts: list[str]) -> list[list[str]]:
    rows: list[list[str]] = []
    section = ""
    pending_label = ""
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
        if key == "7 gun ihbarli":
            pending_label = line
            continue
        parsed = _parse_amount_line(line)
        if not parsed:
            continue
        label = _clean_label(parsed[0], DEPOSIT_INTEREST_TITLE)
        if pending_label and _base.normalize_key(label) == "mevduat":
            label = f"{pending_label} {label}"
        pending_label = ""
        if section and not _base.normalize_key(label).startswith(("toplam", "genel toplam")):
            label = f"{section} - {label}"
        values = parsed[1:]
        if len(values) >= 9:
            values = values[:8]
        rows.append([label, *values])
        if _base.normalize_key(label) == "genel toplam":
            break
    return rows


def extract_equity_rows(page_texts: list[str]) -> list[list[str]]:
    targets = {
        "cekirdek sermaye toplami": ("Çekirdek Sermaye Toplamı", None),
        "katki sermaye toplami": ("Katkı Sermaye Toplamı", None),
        "toplam ozkaynak ana sermaye ve katki sermaye toplami": (
            "Toplam Özkaynak (Ana sermaye ve katkı sermaye toplamı)",
            None,
        ),
        "toplam risk agirlikli tutarlar": ("Toplam Risk Ağırlıklı Tutarlar", None),
    }
    for start_page in _base.title_pages(page_texts, EQUITY_ITEMS_TITLE):
        for text in page_texts[start_page : min(len(page_texts), start_page + 4)]:
            for raw_line in text.splitlines():
                line = _base.normalize_text(raw_line)
                key = _base.normalize_key(line)
                parsed = _parse_amount_line(line)
                if not parsed:
                    continue
                for target_key, (label, value) in list(targets.items()):
                    if value is None and target_key in key:
                        targets[target_key] = (label, parsed[1])
        break

    rows: list[list[str]] = []
    missing: list[str] = []
    for label, value in targets.values():
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
    headers = ["Vadesiz", "1 Aya Kadar", "1-3 Ay", "3-6 Ay", "6 Ay-1 Yıl", "1 Yıl ve Üstü", "Birikimli Mevduat", "Toplam"][
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


def _patch_yapi_kredi_overrides() -> None:
    _base.extract_raw_rows = extract_raw_rows
    _base.extract_restructured_npl_rows = extract_restructured_npl_rows
    _base.extract_stage_provision_rows = extract_stage_provision_rows
    _base.extract_deposit_interest_rows = extract_deposit_interest_rows
    _base.extract_equity_rows = extract_equity_rows
    _base.extract_liquidity_coverage_rows = extract_liquidity_coverage_rows
    _base.to_transposed_total = to_transposed_total
    _base.to_transposed_general_total = to_transposed_general_total


_patch_base()
_patch_yapi_kredi_overrides()

for _name in dir(_base):
    if not _name.startswith("__") and _name not in globals():
        globals()[_name] = getattr(_base, _name)


def main() -> None:
    _patch_base()
    _patch_yapi_kredi_overrides()
    _base.main()


if __name__ == "__main__":
    main()
