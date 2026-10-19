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
_ORIGINAL_EXTRACT_FX_RISK_ROWS = _base.extract_fx_risk_rows


CONSUMER_CREDIT_TITLE = (
    "Tüketici kredileri, bireysel kredi kartları, personel kredileri "
    "ve personel kredi kartlarına ilişkin bilgiler"
)
INSTALLMENT_COMMERCIAL_TITLE = "Taksitli ticari krediler ve kurumsal kredi kartlarına ilişkin bilgiler"
DEPOSIT_MATURITY_TITLE = "Mevduatın vade yapısına ilişkin bilgiler"
DEPOSIT_TOTAL_BY_MATURITY_TITLE = "Mevduatın vade yapısına ilişkin toplam bilgileri"
DEPOSIT_INTEREST_TITLE = "Mevduata Ödenen Faizin Vade Yapısına Göre Gösterimi"
RESTRUCTURED_NPL_TITLE = "Donuk alacaklara ve yeniden yapılandırılan kredilere ilişkin bilgiler"
NPL_MOVEMENT_TITLE = "Toplam donuk alacak hareketlerine ilişkin bilgiler"
BANKS_INTEREST_INCOME_TITLE = "Bankalardan alınan faiz gelirlerine ilişkin bilgiler"
SECURITIES_INTEREST_INCOME_TITLE = "Menkul değerlerden alınan faiz gelirlerine ilişkin bilgiler"
EQUITY_ITEMS_TITLE = "Özkaynak kalemlerine ilişkin bilgiler"
EXPECTED_LOSS_PROVISIONS_TITLE = "Bankaların kredi ve diğer alacaklara ilişkin değer düşüş karşılıkları"
STAGE_PROVISION_TITLE = (
    "Standart nitelikli ve yakın izlemedeki krediler ile yeniden yapılandırılan "
    "yakın izlemedeki kredilere ilişkin bilgiler"
)
STAGE_PROVISION_HEADERS = ["Kalem", "Standart Nitelikli Krediler", "Yakın İzlemedeki Krediler"]
LIQUIDITY_COVERAGE_RATIO_TITLE = "LİKİDİTE KARŞILAMA ORANI"
STAGE3_PROVISIONS_TITLE = "Kredilere ilişkin olarak ayrılan üçüncü aşama karşılıkları"
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
    _base.STAGE_PROVISION_HEADERS = STAGE_PROVISION_HEADERS
    _base.STAGE_PROVISION_TARGET_LABELS = [
        "12 Aylık Beklenen Zarar Karşılığı",
        "Kredi Riskinde Önemli Artış",
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
                            title_key = _base.normalize_key(title)
                            line_key = _base.normalize_key(line)
                            if title_key != line_key:
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
    cleaned = cleaned.replace("İştirakler. Bağlı", "İştirakler, Bağlı")
    return cleaned


def to_transposed_total(rows: list[list[str]], fallback_headers: list[str]) -> list[list[str]]:
    total_row = next((row for row in rows if row and _base.normalize_key(row[0]).startswith(("toplam", "genel toplam"))), rows[-1])
    headers = [
        "Vadesiz",
        "1 Aya Kadar",
        "1-3 Ay",
        "3-6 Ay",
        "6 Ay-1 Yıl",
        "1 Yıl ve Üstü",
        "Birikimli Mevduat",
        "Toplam",
    ][: len(total_row) - 1]
    return [[header, "", "", _base.clean_cell(value)] for header, value in zip(headers, total_row[1:])]


def to_transposed_general_total(rows: list[list[str]], fallback_headers: list[str]) -> list[list[str]]:
    total_row = next((row for row in rows if row and _base.normalize_key(row[0]) == "genel toplam"), None)
    if total_row is None:
        total_row = next((row for row in reversed(rows) if row and _base.normalize_key(row[0]).endswith("toplam")), rows[-1])
    headers = [
        "Vadesiz Mevduat",
        "1 Aya Kadar",
        "3 Aya Kadar",
        "6 Aya Kadar",
        "1 Yıla Kadar",
        "1 Yıldan Uzun",
        "Toplam",
    ][: len(total_row) - 1]
    return [[header, "", "", _base.clean_cell(value)] for header, value in zip(headers, total_row[1:])]


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


def _plain_current_rows(
    page_texts: list[str],
    title: str,
    stop_keys: tuple[str, ...] = (),
    max_pages: int = 1,
) -> list[list[str]]:
    rows: list[list[str]] = []
    in_current = False
    for line in _lines_after_title(page_texts, title, max_pages=max_pages):
        key = _base.normalize_key(line)
        if rows and any(key.startswith(stop_key) for stop_key in stop_keys):
            break
        if "cari donem" in key:
            in_current = True
            continue
        if "onceki donem" in key:
            if rows:
                break
            in_current = False
            continue
        if not in_current:
            continue
        if key in {
            "kisa vadeli orta ve uzun vadeli toplam",
            "iii grup iv grup v grup",
            "tp yp tp yp",
        }:
            continue
        parsed = _parse_amount_line(line)
        if parsed:
            rows.append(parsed)
    return rows


def _rows_after_marker(
    page_texts: list[str],
    title: str,
    start_key: str,
    stop_keys: tuple[str, ...],
    max_pages: int = 1,
    require_current: bool = False,
) -> list[list[str]]:
    lines = _lines_after_title(page_texts, title, max_pages=max_pages, include_title_tail=True)
    rows: list[list[str]] = []
    started = not start_key
    in_current = not require_current
    for line in lines:
        key = _base.normalize_key(line)
        if not started:
            if key.startswith(start_key):
                started = True
            continue
        if rows and any(key.startswith(stop_key) for stop_key in stop_keys):
            break
        if "cari donem" in key:
            in_current = True
            continue
        if "onceki donem" in key:
            if rows:
                break
            in_current = False
            continue
        if require_current and not in_current:
            continue
        if (
            not line
            or key.startswith("31 mart")
            or key in {"iii grup iv grup v grup", "tp yp tp yp"}
            or key.startswith("krediler krediler krediler")
        ):
            continue
        parsed = _parse_amount_line(line)
        if parsed:
            rows.append([_clean_label(parsed[0], title), *parsed[1:]])
    return rows


def _expected_loss_rows(page_texts: list[str]) -> list[list[str]]:
    for start_page in _base.title_pages(page_texts, EXPECTED_LOSS_PROVISIONS_TITLE):
        text = page_texts[start_page]
        if "beklenen kredi zarar karsiliklari" not in _base.normalize_key(text):
            continue
        lines = []
        started = False
        for raw_line in text.splitlines():
            line = _base.normalize_text(raw_line)
            key = _base.normalize_key(line)
            if not started:
                if key.startswith("beklenen zarar karsiliklari"):
                    started = True
                continue
            if key.startswith("f diger faaliyet") or key.startswith("diger faaliyet gider"):
                break
            if key.startswith("31 mart") or key.startswith("beklenen zarar karsiliklari"):
                continue
            lines.append(line)
        rows = _combine_wrapped_rows(lines)
        if rows:
            return rows
    return []


def _npl_movement_rows(page_texts: list[str]) -> list[list[str]]:
    lines = []
    for line in _lines_after_title(page_texts, NPL_MOVEMENT_TITLE, max_pages=1, include_title_tail=True):
        key = _base.normalize_key(line)
        if key.startswith(("iii yabanci", "iii donuk", "iv donuk", "v donuk")):
            break
        if (
            not line
            or key in {"iii grup iv grup v grup", "krediler krediler krediler"}
            or key.startswith("tahsil imkani sinirli tahsili supheli")
            or line.startswith("(*)")
            or line.startswith("(**)")
        ):
            continue
        lines.append(re.sub(r":?\s*31\s+Aralık\s+20\d{2}", "", line).strip())
    return _combine_wrapped_rows(lines, skip_period_markers=False)


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
        rows = _plain_current_rows(page_texts, title, ("onceki donem",), max_pages=2)
        if rows:
            return rows
    if title_key == _base.normalize_key(INSTALLMENT_COMMERCIAL_TITLE):
        rows = _plain_current_rows(page_texts, title, ("onceki donem",), max_pages=1)
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
    if title_key == _base.normalize_key(STAGE3_PROVISIONS_TITLE):
        rows = _rows_after_marker(page_texts, title, "", ("8 donuk",), max_pages=1, require_current=True)
        if rows:
            allowed_labels = {
                "tahsil imkani sinirli krediler icin ayrilanlar",
                "tahsili supheli krediler icin ayrilanlar",
                "zarar niteligindeki krediler icin ayrilanlar",
                "toplam",
            }
            return [row for row in rows if row and _base.normalize_key(row[0]) in allowed_labels]
    return _ORIGINAL_EXTRACT_RAW_ROWS(
        page_texts,
        page_tables,
        title,
        include_empty_amount_rows=include_empty_amount_rows,
    )


def extract_restructured_npl_rows(page_texts: list[str]) -> list[list[str]]:
    return _rows_after_marker(
        page_texts,
        RESTRUCTURED_NPL_TITLE,
        "",
        ("onceki donem", "ii toplam donuk"),
        max_pages=1,
        require_current=True,
    )


def extract_stage_ageing_rows(page_texts: list[str]) -> list[list[str]]:
    """Extract Akbank's current-period ageing analysis into Total slots."""
    target_labels = {
        "30 60 gun gecikmis": "30-60 gün gecikmiş",
        "60 90 gun gecikmis": "60-90 gün gecikmiş",
        "90 gun uzeri gecikmis": "90 gün üzeri gecikmiş",
        "toplam": "Toplam",
    }
    for text in page_texts:
        text_key = _base.normalize_key(text)
        if not all(
            token in text_key
            for token in (
                "30 60 gun gecikmis",
                "60 90 gun gecikmis",
                "90 gun uzeri gecikmis",
            )
        ):
            continue
        output: list[list[str]] = []
        in_ageing_section = False
        for raw_line in text.splitlines():
            line_key = _base.normalize_key(raw_line)
            if line_key.startswith("30 60 gun gecikmis"):
                in_ageing_section = True
            if not in_ageing_section:
                continue
            parsed = _parse_amount_line(raw_line)
            if not parsed:
                continue
            label_key = _base.normalize_key(parsed[0])
            canonical_label = target_labels.get(label_key)
            if canonical_label is None:
                continue
            # The first amount is Cari Donem; the second is Onceki Donem.
            output.append([canonical_label, "", "", parsed[1]])
            if label_key == "toplam":
                return output
        if output:
            return output
    return []


def extract_stage_provision_rows(page_texts: list[str]) -> list[list[str]]:
    targets = (
        ("12 aylik beklenen zarar karsiligi", "12 Aylık Beklenen Zarar Karşılığı"),
        ("kredi riskinde onemli artis", "Kredi Riskinde Önemli Artış"),
    )
    for text in page_texts:
        text_key = _base.normalize_key(text)
        if not all(target in text_key for target, _label in targets):
            continue
        rows: list[list[str]] = []
        for raw_line in text.splitlines():
            parsed = _parse_amount_line(raw_line)
            if not parsed:
                continue
            label_key = _base.normalize_key(parsed[0])
            canonical = next((label for target, label in targets if target in label_key), None)
            if canonical is not None and len(parsed) >= 3:
                rows.append([canonical, parsed[1], parsed[2]])
        if len(rows) >= 2:
            return rows[:2]
    return []


def extract_deposit_interest_rows(page_texts: list[str]) -> list[list[str]]:
    return _ORIGINAL_EXTRACT_DEPOSIT_INTEREST_ROWS(page_texts)


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


def extract_fx_risk_rows(page_texts: list[str]) -> list[list[str]]:
    """Akbank PDF metninde bölünebilen 'aritmetik' ifadesini tolere eder."""
    try:
        return _ORIGINAL_EXTRACT_FX_RISK_ROWS(page_texts)
    except RuntimeError:
        lines = [
            _base.normalize_text(line)
            for text in page_texts
            for line in text.splitlines()
            if _base.normalize_text(line)
        ]
        current = None
        average = None
        for line in lines:
            key = _base.normalize_key(line)
            values = _base._fx_rate_tokens(line)
            if current is None and "bilanco degerleme kuru" in key and values:
                current = values[0]
        for index, line in enumerate(lines):
            key = _base.normalize_key(line)
            if "otuz gunluk" not in key or "ortalama" not in key:
                continue
            for nearby in lines[index + 1:index + 7]:
                nearby_key = _base.normalize_key(nearby)
                values = _base._fx_rate_tokens(nearby)
                if values and nearby_key.startswith(("usd", "abd dolari")):
                    average = values[0]
                    break
            if average is not None:
                break
        if current is None or average is None:
            raise RuntimeError(
                f"Akbank USD kur satirlari eksik: cari={current!r}, ortalama={average!r}"
            )
        return [
            ["USD Cari Döviz Alış Kuru", "", "", current],
            ["USD Aritmetik Ortalama", "", "", average],
        ]


def _patch_akbank_overrides() -> None:
    _base.extract_raw_rows = extract_raw_rows
    _base.extract_restructured_npl_rows = extract_restructured_npl_rows
    _base.extract_stage_provision_rows = extract_stage_provision_rows
    _base.extract_deposit_interest_rows = extract_deposit_interest_rows
    _base.extract_liquidity_coverage_rows = extract_liquidity_coverage_rows
    _base.extract_fx_risk_rows = extract_fx_risk_rows
    _base.to_transposed_total = to_transposed_total
    _base.to_transposed_general_total = to_transposed_general_total


_patch_base()
_patch_akbank_overrides()


def extract_all_tables(pdf_path: Path):
    tables = _base.extract_all_tables(Path(pdf_path))
    with _base.pdfplumber.open(str(pdf_path)) as pdf:
        page_texts = [page.extract_text(x_tolerance=1, y_tolerance=3) or "" for page in pdf.pages]
    ageing_rows = extract_stage_ageing_rows(page_texts)
    if ageing_rows:
        tables.append(
            (
                f"{STAGE_PROVISION_TITLE} - gecikme bilgileri",
                ["Kalem", "", "", "Toplam"],
                ageing_rows,
            )
        )
    return tables

for _name in dir(_base):
    if not _name.startswith("__") and _name not in globals():
        globals()[_name] = getattr(_base, _name)


def main() -> None:
    _patch_base()
    _patch_akbank_overrides()
    _base.main()


if __name__ == "__main__":
    main()
