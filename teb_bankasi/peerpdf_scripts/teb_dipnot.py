from __future__ import annotations

import argparse
import re
import sys
import unicodedata
from pathlib import Path


def add_bundled_pdf_dependencies() -> None:
    """Compatibility hook; install dependencies from requirements.txt."""
    return None


add_bundled_pdf_dependencies()

import pdfplumber


CONSUMER_CREDIT_TITLE = (
    "Tüketici kredileri, bireysel kredi kartları, personel kredileri "
    "ve personel kredi kartlarına ilişkin bilgiler"
)
INSTALLMENT_COMMERCIAL_TITLE = "Taksitli ticari krediler ve kurumsal kredi kartlarına ilişkin bilgiler"
DEPOSIT_MATURITY_TITLE = "Mevduatın vade yapısına ilişkin bilgiler"
DEPOSIT_TOTAL_BY_MATURITY_TITLE = "Mevduatın vade yapısına ilişkin toplam bilgileri"
DEPOSIT_INTEREST_TITLE = "Mevduata ödenen faizin vade yapısına göre gösterimi"
RESTRUCTURED_NPL_TITLE = "Donuk alacaklara ve yeniden yapılandırılan kredilere ilişkin bilgiler"
NPL_MOVEMENT_TITLE = "Toplam donuk alacak hareketlerine ilişkin bilgiler"
BANKS_INTEREST_INCOME_TITLE = "Bankalardan alınan faiz gelirlerine ilişkin bilgiler"
SECURITIES_INTEREST_INCOME_TITLE = "Menkul değerlerden alınan faizlere ilişkin bilgiler"
EQUITY_ITEMS_TITLE = "Özkaynak kalemlerine ilişkin bilgiler"
EXPECTED_LOSS_PROVISIONS_TITLE = "Beklenen zarar karşılıkları ve diğer karşılık giderleri"
STAGE_PROVISION_TITLE = (
    "Standart Nitelikli ve Yakın İzlemedeki krediler ile yeniden yapılandırılan "
    "Yakın İzlemedeki kredilere ilişkin bilgiler"
)
LIQUIDITY_COVERAGE_RATIO_TITLE = "Likidite karşılama oranı"
STAGE3_PROVISIONS_TITLE = (
    "Kredilere ilişkin olarak ayrılan özel karşılıklar veya temerrüt "
    "(Üçüncü Aşama) karşılıkları"
)
OTHER_OPERATING_INCOME_TITLE = "Diğer faaliyet gelirlerine ilişkin açıklamalar"
OTHER_OPERATING_EXPENSE_TITLE = "Diğer faaliyet giderlerine ilişkin bilgiler"
FX_RISK_TITLE = "Kur riskine ilişkin açıklamalar"
CASH_AND_TCMB_TITLE = "Nakit değerler ve TCMB’ye ilişkin bilgiler"
USED_LOANS_INTEREST_EXPENSE_TITLE = "Kullanılan kredilere verilen faizlere ilişkin bilgiler"
LOAN_INTEREST_INCOME_TITLE = "Kredilerden alınan faiz gelirlerine ilişkin bilgiler"

CREDIT_HEADERS = ["Kalem", "Kısa Vadeli", "Orta ve Uzun Vadeli", "Toplam"]
DEPOSIT_HEADERS = ["Kalem", "Vadesiz", "Vadeli", "Toplam"]
DEPOSIT_TOTAL_HEADERS = ["Vade", "", "", "Toplam"]
NPL_HEADERS = ["Kalem", "III. Grup", "IV. Grup", "V. Grup"]
INTEREST_HEADERS = ["Kalem", "TP", "YP"]
TOTAL_HEADERS = ["Kalem", "", "", "Toplam"]
OTHER_OPERATING_EXPENSE_HEADERS = ["Kalem", "", "", "Toplam"]
FX_RISK_HEADERS = ["Kalem", "", "", "Kur"]
LIQUIDITY_HEADERS = ["Kalem", "TP+YP", "YP"]
STAGE_PROVISION_HEADERS = [
    "Kalem",
    "Cari Donem - Standart Nitelikli Krediler",
    "Cari Donem - Yakin Izlemedeki Krediler",
    "Onceki Donem - Standart Nitelikli Krediler",
    "Onceki Donem - Yakin Izlemedeki Krediler",
]
STAGE_PROVISION_TARGET_LABELS = [
    "12 Aylık Beklenen Zarar Karşılığı",
    "Kredi Riskinde Önemli Artış",
]

NUMBER_RE = re.compile(r"^-?\d+(?:[.,]\d{3})*(?:[.,]\d+)?$")
AMOUNT_TOKEN_RE = re.compile(r"^(?:-|\d+(?:,\d{3})*(?:\.\d+)?|\d+(?:\.\d{3})*(?:,\d+)?)$")
SIGNED_AMOUNT_TOKEN_RE = re.compile(r"^(?:-|\(?\d+(?:,\d{3})*(?:\.\d+)?\)?|\(?\d+(?:\.\d{3})*(?:,\d+)?\)?)$")


def _expense_amount_token(value: str) -> bool:
    cleaned = normalize_text(value).rstrip(".,;")
    return bool(
        re.fullmatch(
            r"(?:[-–—]+|\(?\d{1,3}(?:[.,]\d{3})*(?:[.,]\d+)?\)?)",
            cleaned,
        )
    )


def _expense_key(value: str) -> str:
    return (
        normalize_key(value)
        .replace("o nceki", "onceki")
        .replace("t oplam", "toplam")
    )


def _is_complete_expense_label(value: str) -> bool:
    key = _expense_key(value)
    exact_fragments = (
        "kidem tazminati karsiligi",
        "banka sosyal yardim sandigi varlik aciklari karsiligi",
        "banka emekli ve yardim sandigi varlik aciklari karsiligi",
        "maddi duran varlik deger dusus gider",
        "maddi olmayan duran varlik deger dusus gider",
        "serefiye deger dusus gider",
        "ozkaynak yontemi uygulanan ortaklik paylari deger dusus gider",
        "elden cikarilacak kiymetler deger dusus gider",
        "elden cikarilacak menkul kiymetler deger dusus gider",
        "elden cikarilacak kiymetler amortisman gider",
        "elden cikarilacak menkul kiymetler amortisman gider",
    )
    return any(fragment in key for fragment in exact_fragments) or (
        "satis amacli" in key and "deger dusus gider" in key
    )


def _expense_rows_after_header(lines: list[str], start: int) -> list[list[str]]:
    rows: list[list[str]] = []
    pending: list[str] = []
    for raw_line in lines[start : start + 80]:
        line = normalize_text(raw_line)
        key = _expense_key(line)
        if not line:
            continue
        if rows and (
            line.startswith(("(*)", "(**)", "(***)"))
            or re.match(r"^(?:[0-9]+[.)]|[ivx]+[.)])\s+", key)
        ):
            break
        parts = line.split()
        if len(parts) >= 2 and _expense_amount_token(parts[-1]) and _expense_amount_token(parts[-2]):
            inline_label = " ".join(parts[:-2])
            if not inline_label and not pending:
                continue
            label = normalize_text(" ".join([*pending, inline_label]))
            pending = []
            if not label or _expense_key(label) in {"cari donem", "onceki donem"}:
                continue
            rows.append([label, "", "", parts[-2].rstrip(".,;")])
            if _expense_key(label).startswith("toplam"):
                break
            continue
        if re.fullmatch(r"\d{1,2}\s+\S+\s+20\d{2}", line) or re.fullmatch(
            r"\d{1,2}[./]\d{1,2}[./]20\d{2}", line
        ):
            continue
        if "cari donem" in key or "onceki donem" in key:
            pending = []
            continue
        pending.append(line)
        if len(" ".join(pending)) > 300:
            pending = [line]
    return rows


def extract_other_operating_expense_rows(page_texts: list[str]) -> list[list[str]]:
    """Extract current-period rows from cross-bank other-operating-expense layouts."""
    candidates: list[list[list[str]]] = []
    for page_index, page_text in enumerate(page_texts):
        if "diger faaliyet gider" not in _expense_key(page_text):
            continue
        lines = [normalize_text(line) for line in page_text.splitlines() if normalize_text(line)]
        for index, line in enumerate(lines):
            key = _expense_key(line)
            if "cari donem" not in key:
                continue
            if "onceki donem" in key:
                start = index + 1
            else:
                previous_index = next(
                    (
                        probe
                        for probe in range(index + 1, min(len(lines), index + 6))
                        if "onceki donem" in _expense_key(lines[probe])
                    ),
                    None,
                )
                if previous_index is None:
                    continue
                start = previous_index + 1
            rows = _expense_rows_after_header(lines, start)
            if rows:
                candidates.append(rows)

    if not candidates:
        raise RuntimeError(f"Hedef tablo satirlari bulunamadi: {OTHER_OPERATING_EXPENSE_TITLE}")

    typical = (
        "kidem tazminati",
        "maddi duran varlik amortisman",
        "diger isletme gider",
        "aktiflerin satisindan dogan zarar",
    )

    def score(rows: list[list[str]]) -> tuple[int, int]:
        keys = [_expense_key(row[0]) for row in rows]
        points = len(rows)
        points += 100 if any(key.startswith("toplam") for key in keys) else 0
        points += 20 * sum(any(token in key for key in keys) for token in typical)
        return points, len(rows)

    best = max(candidates, key=score)
    if score(best)[0] < 25:
        raise RuntimeError(f"Hedef gider tablosu guvenilir sekilde ayrisitirilamadi: {OTHER_OPERATING_EXPENSE_TITLE}")
    return best


def extract_other_operating_expense_rows_from_words(
    page_texts: list[str],
    page_words: list[list[dict[str, object]]],
) -> list[list[str]]:
    """Use positioned PDF words to preserve blank cells and wrapped labels."""
    candidates: list[list[list[str]]] = []
    for page_index, words in enumerate(page_words):
        if not words or "diger faaliyet gider" not in _expense_key(page_texts[page_index]):
            continue
        groups: list[list[dict[str, object]]] = []
        for word in sorted(words, key=lambda item: (float(item["top"]), float(item["x0"]))):
            group = next(
                (
                    existing
                    for existing in reversed(groups[-4:])
                    if abs(float(existing[0]["top"]) - float(word["top"])) <= 2.4
                ),
                None,
            )
            if group is None:
                group = []
                groups.append(group)
            group.append(word)

        title_tops = []
        for group in groups:
            text = " ".join(str(word["text"]) for word in sorted(group, key=lambda item: float(item["x0"])))
            if "diger faaliyet gider" in _expense_key(text):
                title_tops.append(float(group[0]["top"]))
        if not title_tops:
            continue

        for header_index, group in enumerate(groups):
            ordered = sorted(group, key=lambda item: float(item["x0"]))
            tokens = [str(word["text"]) for word in ordered]
            header_key = _expense_key(" ".join(tokens))
            if "cari donem" not in header_key or "onceki donem" not in header_key:
                continue
            header_top = float(group[0]["top"])
            if not any(0 <= header_top - top <= 220 for top in title_tops):
                continue
            current_words = [word for word in ordered if "cari" in _expense_key(str(word["text"]))]
            previous_words = [word for word in ordered if "onceki" in _expense_key(str(word["text"]))]
            if not current_words or not previous_words:
                continue
            current_x = float(current_words[0]["x0"])
            previous_x = float(previous_words[0]["x0"])
            if previous_x <= current_x:
                continue
            label_limit = current_x - 18
            midpoint = previous_x - 5
            pending: list[str] = []
            rows: list[list[str]] = []
            row_needs_suffix = False
            for row_group in groups[header_index + 1 :]:
                row_top = float(row_group[0]["top"])
                if row_top - header_top > 360:
                    break
                row_words = sorted(row_group, key=lambda item: float(item["x0"]))
                label = normalize_text(
                    " ".join(str(word["text"]) for word in row_words if float(word["x0"]) < label_limit)
                )
                current = normalize_text(
                    " ".join(
                        str(word["text"])
                        for word in row_words
                        if label_limit <= float(word["x0"]) < midpoint
                    )
                )
                previous = normalize_text(
                    " ".join(str(word["text"]) for word in row_words if float(word["x0"]) >= midpoint)
                )
                if not current and not previous:
                    if label:
                        if row_needs_suffix and rows:
                            rows[-1][0] = normalize_text(f"{rows[-1][0]} {label}")
                            row_needs_suffix = False
                        elif _is_complete_expense_label(label):
                            rows.append([label, "", "", "-"])
                        else:
                            pending.append(label)
                    continue
                if not (
                    (not current or _expense_amount_token(current))
                    and (not previous or _expense_amount_token(previous))
                    and (current or previous)
                ):
                    continue
                full_label = normalize_text(" ".join([*pending, label]))
                pending = []
                if not full_label:
                    continue
                rows.append([full_label, "", "", current or "-"])
                row_needs_suffix = not _is_complete_expense_label(full_label)
                if _expense_key(full_label).startswith("toplam"):
                    break
            if rows:
                candidates.append(rows)

    if not candidates:
        return extract_other_operating_expense_rows(page_texts)

    def score(rows: list[list[str]]) -> tuple[int, int]:
        keys = [_expense_key(row[0]) for row in rows]
        typical = (
            "kidem tazminati",
            "maddi duran varlik amortisman",
            "diger isletme gider",
            "aktiflerin satisindan dogan zarar",
        )
        points = len(rows)
        points += 100 if any(key.startswith("toplam") for key in keys) else 0
        points += 20 * sum(any(token in key for key in keys) for token in typical)
        return points, len(rows)

    best = max(candidates, key=score)
    return best if score(best)[0] >= 25 else extract_other_operating_expense_rows(page_texts)


def normalize_text(value: str | None) -> str:
    if value is None:
        return ""
    return " ".join(str(value).replace("\n", " ").split())


def normalize_key(value: str | None) -> str:
    text = normalize_text(value).casefold()
    text = text.replace("ı", "i").replace("İ", "i")
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.replace("’", "'").replace("`", "'")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def clean_cell(value: str | None) -> str:
    cleaned = normalize_text(value)
    return cleaned if cleaned else "-"


def parse_amount_text(value: str | None) -> int | None:
    cleaned = normalize_text(value)
    if cleaned in {"", "-"}:
        return None
    digits = "".join(char for char in cleaned if char.isdigit())
    return int(digits) if digits else None


def format_amount(value: int | None) -> str:
    if value is None:
        return "-"
    return f"{value:,}"


def sum_amounts(values: list[str]) -> int | None:
    parsed = [parse_amount_text(value) for value in values]
    present = [value for value in parsed if value is not None]
    if not present:
        return None
    return sum(present)


def has_numeric_value(row: list[str]) -> bool:
    for value in row[1:]:
        cleaned = normalize_text(value).replace(" ", "")
        if cleaned == "-":
            continue
        if NUMBER_RE.match(cleaned):
            return True
    return False


def is_current_marker(row_text: str) -> bool:
    key = normalize_key(row_text)
    return "cari donem" in key or "current period" in key


def is_previous_marker(row_text: str) -> bool:
    key = normalize_key(row_text)
    if key.startswith("onceki donem sonu bakiyesi"):
        return False
    return "onceki donem" in key or "previous period" in key


def is_noise_row(row: list[str]) -> bool:
    first = normalize_key(row[0] if row else "")
    if not first:
        return True
    if first in {"kalem", "cari donem", "onceki donem"}:
        return True
    if first.isdigit():
        return True
    return False


def normalize_row(raw_row: list[str | None], keep_trailing_dashes: bool = False) -> list[str]:
    row = [clean_cell(cell) for cell in raw_row]
    while not keep_trailing_dashes and row and row[-1] == "-":
        row.pop()
    return row


def has_amount_cells(row: list[str]) -> bool:
    return len(row) >= 4 and all(AMOUNT_TOKEN_RE.match(clean_cell(value)) for value in row[1:4])


def parse_current_rows(raw_table: list[list[str | None]], include_empty_amount_rows: bool = False) -> list[list[str]]:
    rows: list[list[str]] = []
    in_current = False
    saw_marker = False
    for raw_row in raw_table:
        row = normalize_row(raw_row, keep_trailing_dashes=include_empty_amount_rows)
        if not row:
            continue
        row_text = " ".join(row)
        if is_previous_marker(row_text):
            if in_current or saw_marker:
                break
            continue
        if is_current_marker(row_text):
            in_current = True
            saw_marker = True
            continue
        if saw_marker and not in_current:
            continue
        if is_noise_row(row):
            continue
        if include_empty_amount_rows:
            if not has_amount_cells(row):
                continue
        elif not has_numeric_value(row):
            continue
        rows.append(row)
    if rows:
        return rows

    for raw_row in raw_table:
        row = normalize_row(raw_row, keep_trailing_dashes=include_empty_amount_rows)
        if not row or is_noise_row(row):
            continue
        if include_empty_amount_rows:
            if has_amount_cells(row):
                rows.append(row)
        elif has_numeric_value(row):
            rows.append(row)
    return rows


def parse_text_line(line: str) -> list[str] | None:
    cleaned = normalize_text(line)
    if not cleaned or is_current_marker(cleaned) or is_previous_marker(cleaned):
        return None
    if cleaned.startswith("(") or cleaned.isdigit():
        return None
    parts = cleaned.split()
    values: list[str] = []
    while parts and AMOUNT_TOKEN_RE.match(parts[-1]):
        values.insert(0, parts.pop())
    if len(values) < 2 or not parts:
        return None
    label = " ".join(parts)
    if normalize_key(label) in {"kisa vadeli", "orta ve uzun vadeli toplam", "tp yp"}:
        return None
    return [label, *values]


def parse_text_rows_for_title(page_texts: list[str], title_pages: list[int], title: str) -> list[list[str]]:
    title_key = normalize_key(title)
    title_prefix = " ".join(title_key.split()[:5])
    rows: list[list[str]] = []
    started = False
    for start_page in title_pages:
        for text in page_texts[start_page : min(len(page_texts), start_page + 2)]:
            for raw_line in text.splitlines():
                line = normalize_text(raw_line)
                line_key = normalize_key(line)
                if not started:
                    if title_key in line_key or (title_prefix and title_prefix in line_key):
                        started = True
                    continue
                if rows and re.match(r"^[a-zçğıöşü]\)", line.strip(), flags=re.IGNORECASE):
                    return rows
                parsed = parse_text_line(line)
                if parsed:
                    rows.append(parsed)
                    if normalize_key(parsed[0]).startswith("toplam"):
                        return rows
            if rows:
                return rows
    return rows


def parse_two_amount_text_line(line: str) -> list[str] | None:
    cleaned = normalize_text(line)
    if not cleaned or is_current_marker(cleaned) or is_previous_marker(cleaned):
        return None
    if cleaned.startswith("(") or cleaned.isdigit():
        return None
    parts = cleaned.split()
    values: list[str] = []
    while parts and SIGNED_AMOUNT_TOKEN_RE.match(parts[-1]):
        values.insert(0, parts.pop())
    if len(values) < 2 or not parts:
        return None
    return [" ".join(parts), *values[:2]]


def extract_expected_loss_provision_rows(page_texts: list[str]) -> list[list[str]]:
    pages = title_pages(page_texts, EXPECTED_LOSS_PROVISIONS_TITLE)
    if not pages:
        return []
    title_key = normalize_key(EXPECTED_LOSS_PROVISIONS_TITLE)
    rows: list[list[str]] = []
    started = False
    pending_label = ""
    for start_page in pages:
        for text in page_texts[start_page : min(len(page_texts), start_page + 2)]:
            for raw_line in text.splitlines():
                line = normalize_text(raw_line)
                key = normalize_key(line)
                if not started:
                    if title_key in key:
                        started = True
                    continue
                if rows and key.startswith("6 diger faaliyet"):
                    return rows
                if (
                    not line
                    or key in {"cari donem onceki donem"}
                    or key.startswith(("cari donem", "onceki donem"))
                ):
                    continue
                parsed = parse_two_amount_text_line(line)
                if parsed:
                    label = " ".join(part for part in [pending_label, parsed[0]] if part).strip()
                    rows.append([label, *parsed[1:3]])
                    pending_label = ""
                    if normalize_key(label).startswith("toplam"):
                        return rows
                    continue
                pending_label = " ".join(part for part in [pending_label, line] if part).strip()
        if rows:
            return rows
    return rows


def title_pages(page_texts: list[str], title: str) -> list[int]:
    target = normalize_key(title)
    exact = [idx for idx, text in enumerate(page_texts) if target in normalize_key(text)]
    if exact:
        return exact
    # Bank PDFs frequently insert qualifiers such as "özel karşılıklar veya"
    # into an otherwise unchanged regulatory heading.  The first five words
    # identify our target headings reliably and let old periods resolve without
    # weakening exact matching for normal periods.
    prefix = " ".join(target.split()[:4])
    return [idx for idx, text in enumerate(page_texts) if prefix and prefix in normalize_key(text)]


def extract_raw_rows(
    page_texts: list[str],
    page_tables: list[list[list[list[str | None]]]],
    title: str,
    include_empty_amount_rows: bool = False,
) -> list[list[str]]:
    if normalize_key(title) == normalize_key(OTHER_OPERATING_INCOME_TITLE):
        return []
    if normalize_key(title) == normalize_key(OTHER_OPERATING_EXPENSE_TITLE):
        return extract_other_operating_expense_rows(page_texts)
    if normalize_key(title) == normalize_key(EXPECTED_LOSS_PROVISIONS_TITLE):
        rows = extract_expected_loss_provision_rows(page_texts)
        if rows:
            return rows
    pages = title_pages(page_texts, title)
    if not pages:
        raise RuntimeError(f"Hedef baslik PDF icinde bulunamadi: {title}")
    best_rows: list[list[str]] = []
    best_score = -1
    for start_page in pages:
        for tables in page_tables[start_page : min(len(page_tables), start_page + 3)]:
            for table in tables:
                rows = parse_current_rows(table, include_empty_amount_rows=include_empty_amount_rows)
                score = len(rows) * 10 + sum(len(row) for row in rows)
                if rows and score > best_score:
                    best_rows = rows
                    best_score = score
    if not best_rows:
        best_rows = parse_text_rows_for_title(page_texts, pages, title)
    if not best_rows:
        raise RuntimeError(f"Hedef tablo bulundu ancak satir ayrisitirilamadi: {title}")
    if normalize_key(title) == normalize_key(SECURITIES_INTEREST_INCOME_TITLE):
        for row in best_rows:
            if row and normalize_key(row[0]) == "varliklar":
                row[0] = "Gerçeğe Uygun Değer Farkı Diğer Kapsamlı Gelire Yansıtılan Finansal Varlıklar"
    return best_rows


def trim_row(row: list[str], width: int) -> list[str]:
    output = [clean_cell(value) for value in row[:width]]
    while len(output) < width:
        output.append("")
    return output


def to_total_row(row: list[str]) -> list[str]:
    if len(row) >= 4:
        return [row[0], "", "", row[3]]
    if len(row) >= 2:
        return [row[0], "", "", row[1]]
    return [row[0], "", "", ""]


def to_interest_row(row: list[str]) -> list[str]:
    return trim_row(row, 3)


def to_deposit_summary(rows: list[list[str]]) -> list[list[str]]:
    output: list[list[str]] = []
    for row in rows:
        label = clean_cell(row[0])
        vadesiz = clean_cell(row[1] if len(row) > 1 else "")
        toplam = clean_cell(row[-1] if len(row) > 1 else "")
        toplam_amount = parse_amount_text(toplam)
        vadesiz_amount = parse_amount_text(vadesiz)
        vadeli = None if toplam_amount is None else toplam_amount - (vadesiz_amount or 0)
        output.append([label, vadesiz, format_amount(vadeli), toplam])
    return output


def to_transposed_total(rows: list[list[str]], fallback_headers: list[str]) -> list[list[str]]:
    total_row = next((row for row in rows if row and normalize_key(row[0]).startswith(("toplam", "genel toplam"))), rows[-1])
    headers = fallback_headers[: len(total_row) - 1]
    return [[header, "", "", clean_cell(value)] for header, value in zip(headers, total_row[1:])]


def to_transposed_general_total(rows: list[list[str]], fallback_headers: list[str]) -> list[list[str]]:
    total_row = next((row for row in rows if row and normalize_key(row[0]) == "genel toplam"), None)
    if total_row is None:
        total_row = next((row for row in reversed(rows) if row and normalize_key(row[0]).endswith("toplam")), rows[-1])
    headers = fallback_headers[: len(total_row) - 1]
    return [[header, "", "", clean_cell(value)] for header, value in zip(headers, total_row[1:])]


def liquidity_rows(rows: list[list[str]]) -> list[list[str]]:
    for row in rows:
        key = normalize_key(row[0])
        if "likidite karsilama orani" in key:
            return [trim_row(row, 3)]
    raise RuntimeError("Hedef likidite satiri bulunamadi: Likidite Karşılama Oranı (%)")


def extract_deposit_interest_rows(page_texts: list[str]) -> list[list[str]]:
    pages = title_pages(page_texts, DEPOSIT_INTEREST_TITLE)
    if not pages:
        raise RuntimeError(f"Hedef baslik PDF icinde bulunamadi: {DEPOSIT_INTEREST_TITLE}")

    rows: list[list[str]] = []
    section = ""
    started = False
    for page_index in pages:
        for raw_line in page_texts[page_index].splitlines():
            line = normalize_text(raw_line)
            key = normalize_key(line)
            if normalize_key(DEPOSIT_INTEREST_TITLE) in key:
                started = True
                continue
            if not started:
                continue
            if key.startswith("onceki donem"):
                break
            if key == "turk parasi":
                section = "Türk Parası"
                continue
            if key == "yabanci para":
                section = "Yabancı Para"
                continue
            parsed = parse_text_line(line)
            if not parsed:
                continue
            label_key = normalize_key(parsed[0])
            if label_key == "genel toplam":
                rows.append(["Genel Toplam", *parsed[1:]])
                return rows
            if section and label_key not in {"hesap adi", "vadeli mevduat"}:
                rows.append([f"{section} - {parsed[0]}", *parsed[1:]])
        if rows:
            break

    if not rows:
        raise RuntimeError(f"Hedef cari donem satirlari bulunamadi: {DEPOSIT_INTEREST_TITLE}")
    return rows


def extract_liquidity_coverage_rows(page_texts: list[str]) -> list[list[str]]:
    pages = title_pages(page_texts, LIQUIDITY_COVERAGE_RATIO_TITLE)
    if not pages:
        raise RuntimeError(f"Hedef baslik PDF icinde bulunamadi: {LIQUIDITY_COVERAGE_RATIO_TITLE}")

    target_key = normalize_key("Likidite Karşılama Oranı (%)")
    for page_index in pages:
        for raw_line in page_texts[page_index].splitlines():
            parsed = parse_text_line(raw_line)
            if not parsed:
                continue
            if normalize_key(parsed[0]) == target_key:
                return [trim_row(["Likidite Karşılama Oranı (%)", *parsed[1:]], 3)]

    raise RuntimeError("Hedef likidite satiri bulunamadi: Likidite Karşılama Oranı (%)")


def extract_restructured_npl_rows(page_texts: list[str]) -> list[list[str]]:
    pages = title_pages(page_texts, RESTRUCTURED_NPL_TITLE)
    if not pages:
        raise RuntimeError(f"Hedef baslik PDF icinde bulunamadi: {RESTRUCTURED_NPL_TITLE}")
    rows: list[list[str]] = []
    for page_index in pages:
        in_section = False
        in_current = False
        for raw_line in page_texts[page_index].splitlines():
            line = normalize_text(raw_line)
            key = normalize_key(line)
            if normalize_key(RESTRUCTURED_NPL_TITLE) in key:
                in_section = True
                continue
            if not in_section:
                continue
            if key == "cari donem":
                in_current = True
                continue
            if key == "onceki donem":
                break
            if not in_current:
                continue
            parsed = parse_text_line(line)
            if parsed:
                rows.append(trim_row(parsed, 4))
        if rows:
            return rows
    raise RuntimeError(f"Hedef cari donem satirlari bulunamadi: {RESTRUCTURED_NPL_TITLE}")


def extract_stage_provision_rows(page_texts: list[str]) -> list[list[str]]:
    pages = title_pages(page_texts, STAGE_PROVISION_TITLE)
    if not pages:
        raise RuntimeError(f"Hedef baslik PDF icinde bulunamadi: {STAGE_PROVISION_TITLE}")

    target_keys = {normalize_key(label): label for label in STAGE_PROVISION_TARGET_LABELS}
    found: dict[str, list[str]] = {}

    for page_index in pages:
        for raw_line in page_texts[page_index].splitlines():
            parsed = parse_text_line(raw_line)
            if not parsed:
                continue
            label_key = normalize_key(parsed[0])
            for target_key, target_label in target_keys.items():
                if label_key == target_key or label_key.startswith(target_key):
                    values = [clean_cell(value) for value in parsed[1:5]]
                    while len(values) < 4:
                        values.append("")
                    found[target_label] = [target_label, *values]
        if all(label in found for label in STAGE_PROVISION_TARGET_LABELS):
            return [found[label] for label in STAGE_PROVISION_TARGET_LABELS]

    missing = [label for label in STAGE_PROVISION_TARGET_LABELS if label not in found]
    raise RuntimeError(f"Hedef karsilik satirlari bulunamadi: {', '.join(missing)}")


def first_amount_from_line(line: str, label: str) -> str | None:
    normalized_line = normalize_text(line)
    if not normalize_key(normalized_line).startswith(normalize_key(label)):
        return None
    tail = normalized_line[len(label) :]
    amounts = [part for part in tail.split() if AMOUNT_TOKEN_RE.match(part)]
    return amounts[0] if amounts else None


def extract_equity_rows(page_texts: list[str]) -> list[list[str]]:
    targets = {
        "Çekirdek Sermaye Toplamı": None,
        "Katkı Sermaye Toplamı": None,
    }
    ozkaynak_total = None
    risk_weighted_total = None

    for text in page_texts:
        for line in text.splitlines():
            for label, current_value in list(targets.items()):
                if current_value is not None:
                    continue
                amount = first_amount_from_line(line, label)
                if amount is not None:
                    targets[label] = amount

        lines = [normalize_text(line) for line in text.splitlines()]
        in_ozkaynak_block = False
        for line in lines:
            key = normalize_key(line)
            if key == "ozkaynak":
                in_ozkaynak_block = True
                continue
            if in_ozkaynak_block and key.startswith("sermaye yeterliligi oranlari"):
                in_ozkaynak_block = False
                continue
            if not in_ozkaynak_block:
                continue
            if ozkaynak_total is None:
                ozkaynak_total = first_amount_from_line(
                    line,
                    "Toplam Özkaynak (Ana sermaye ve katkı sermaye toplamı)",
                )
            if risk_weighted_total is None:
                risk_weighted_total = first_amount_from_line(line, "Toplam Risk Ağırlıklı Tutarlar")

    rows = []
    for label in ["Çekirdek Sermaye Toplamı", "Katkı Sermaye Toplamı"]:
        value = targets[label]
        if value is None:
            raise RuntimeError(f"Hedef özkaynak kalemi bulunamadi: {label}")
        rows.append([label, "", "", value])
    if risk_weighted_total is None:
        raise RuntimeError("Hedef özkaynak kalemi bulunamadi: Toplam Risk Ağırlıklı Tutarlar")
    if ozkaynak_total is None:
        raise RuntimeError("Hedef özkaynak kalemi bulunamadi: Toplam Özkaynak")
    rows.append(["Toplam Risk Ağırlıklı Tutarlar", "", "", risk_weighted_total])
    rows.append(["Toplam Özkaynak", "", "", ozkaynak_total])
    return rows


_FX_RATE_RE = re.compile(r"(?<![\d.,])(\d{1,3}[.,]\d{2,7})(?!\d)")


def _fx_rate_tokens(text: str) -> list[str]:
    without_dates = re.sub(r"\b\d{1,2}[./]\d{1,2}[./]20\d{2}\b", " ", text)
    return _FX_RATE_RE.findall(without_dates)


def _usd_value_from_values(values: list[str], header_key: str) -> str | None:
    if not values:
        return None
    usd_pos = min((pos for token in ("usd", "abd dolari") if (pos := header_key.find(token)) >= 0), default=-1)
    eur_pos = min((pos for token in ("eur", "euro", "avro") if (pos := header_key.find(token)) >= 0), default=-1)
    if usd_pos >= 0 and eur_pos >= 0 and len(values) >= 2:
        return values[0] if usd_pos < eur_pos else values[1]
    return values[0]


def extract_fx_risk_rows(page_texts: list[str]) -> list[list[str]]:
    """Extract the two comparable USD rates despite bank/year layout changes."""
    lines: list[str] = []
    for page_text in page_texts:
        page_key = normalize_key(page_text)
        if "kur riskine iliskin" not in page_key and "cari doviz alis kuru" not in page_key:
            continue
        page_lines = [normalize_text(line) for line in page_text.splitlines() if normalize_text(line)]
        # Skip table-of-contents pages; they carry the title but no rate data.
        if not any(_fx_rate_tokens(line) for line in page_lines):
            continue
        lines.extend(page_lines)

    current: str | None = None
    average: str | None = None
    last_currency_header = ""
    seen_average_phrase = False
    previous_key = ""
    for index, line in enumerate(lines):
        key = normalize_key(line)
        if "aritmetik" in key:
            seen_average_phrase = True
        if ("usd" in key or "abd dolari" in key) and any(token in key for token in ("eur", "euro", "avro")):
            last_currency_header = key

        values = _fx_rate_tokens(line)
        if current is None and not seen_average_phrase and key.startswith(("usd ", "abd dolari ")) and len(values) >= 2:
            # Horizontal layout: dates are columns and USD/EUR are rows.
            current = values[-1]
        if current is None and values and "bilanco tarihindeki" in previous_key and "degerleme kuru" in key:
            current = _usd_value_from_values(values, last_currency_header)
        if current is None and values and "bilanco" in key and (
            "degerleme kuru" in key or "cari doviz alis kuru" in key
        ):
            current = _usd_value_from_values(values, last_currency_header)

        if "aritmetik" in key:
            candidate_header = last_currency_header
            for nearby in lines[index:min(len(lines), index + 6)]:
                nearby_key = normalize_key(nearby)
                if ("usd" in nearby_key or "abd dolari" in nearby_key) and any(
                    token in nearby_key for token in ("eur", "euro", "avro")
                ):
                    candidate_header = nearby_key
                nearby_values = _fx_rate_tokens(nearby)
                if not nearby_values:
                    continue
                if "usd" in nearby_key or "abd dolari" in nearby_key:
                    average = nearby_values[0]
                    break
                if candidate_header:
                    average = _usd_value_from_values(nearby_values, candidate_header)
                    break
        previous_key = key

    # Some reports show the current rate as dated rows instead of a labelled
    # balance-sheet valuation row. Pick the latest date, independent of the
    # visual/text extraction order of those rows.
    if current is None:
        months = {
            "ocak": 1, "subat": 2, "mart": 3, "nisan": 4, "mayis": 5, "haziran": 6,
            "temmuz": 7, "agustos": 8, "eylul": 9, "ekim": 10, "kasim": 11, "aralik": 12,
        }
        dated: list[tuple[tuple[int, int, int], str]] = []
        last_currency_header = ""
        for line in lines:
            key = normalize_key(line)
            if ("usd" in key or "abd dolari" in key) and any(token in key for token in ("eur", "euro", "avro")):
                last_currency_header = key
                continue
            values = _fx_rate_tokens(line)
            if not last_currency_header or len(values) < 2:
                continue
            numeric_date = re.search(r"\b(\d{1,2})[./](\d{1,2})[./](20\d{2})\b", line)
            named_date = re.search(r"\b(\d{1,2})\s+([a-z]+)\s+(20\d{2})\b", key)
            if numeric_date:
                day, month, year = map(int, numeric_date.groups())
            elif named_date and named_date.group(2) in months:
                day, month, year = int(named_date.group(1)), months[named_date.group(2)], int(named_date.group(3))
            else:
                continue
            dated.append(((year, month, day), _usd_value_from_values(values, last_currency_header)))
        if dated:
            current = max(dated, key=lambda item: item[0])[1]

    # Final fallback for layouts whose current row is not explicitly dated.
    if current is None:
        last_currency_header = ""
        rows_after_header = 0
        for line in lines:
            key = normalize_key(line)
            if ("usd" in key or "abd dolari" in key) and any(token in key for token in ("eur", "euro", "avro")):
                last_currency_header = key
                rows_after_header = 0
                continue
            if last_currency_header:
                rows_after_header += 1
            values = _fx_rate_tokens(line)
            if last_currency_header and rows_after_header <= 10 and len(values) >= 2 and "aritmetik ortalama" not in key:
                current = _usd_value_from_values(values, last_currency_header)
                break

    if current is None or average is None:
        raise RuntimeError(f"USD kur satirlari eksik: cari={current!r}, ortalama={average!r}")
    return [
        ["USD Cari Döviz Alış Kuru", "", "", current],
        ["USD Aritmetik Ortalama", "", "", average],
    ]


def extract_all_tables(pdf_path: Path) -> list[tuple[str, list[str], list[list[str]]]]:
    with pdfplumber.open(str(pdf_path)) as pdf:
        page_texts = [page.extract_text(x_tolerance=1, y_tolerance=3) or "" for page in pdf.pages]
        page_tables = [page.extract_tables() for page in pdf.pages]
        page_words = [
            page.extract_words(keep_blank_chars=False)
            if "diger faaliyet gider" in _expense_key(text)
            else []
            for page, text in zip(pdf.pages, page_texts)
        ]

    specs = [
        (CONSUMER_CREDIT_TITLE, CREDIT_HEADERS, lambda rows: [trim_row(row, 4) for row in rows]),
        (INSTALLMENT_COMMERCIAL_TITLE, CREDIT_HEADERS, lambda rows: [trim_row(row, 4) for row in rows]),
        (DEPOSIT_MATURITY_TITLE, DEPOSIT_HEADERS, to_deposit_summary),
        (
            DEPOSIT_TOTAL_BY_MATURITY_TITLE,
            DEPOSIT_TOTAL_HEADERS,
            lambda rows: to_transposed_total(
                rows,
                ["Vadesiz", "7 Gün İhbarlı", "1 Aya Kadar", "1-3 Ay", "3-6 Ay", "6 Ay-1 Yıl", "1 Yıl ve Üstü", "Birikimli Mevduat", "Toplam"],
            ),
        ),
        (RESTRUCTURED_NPL_TITLE, NPL_HEADERS, lambda rows: rows),
        (NPL_MOVEMENT_TITLE, NPL_HEADERS, lambda rows: [trim_row(row, 4) for row in rows]),
        (BANKS_INTEREST_INCOME_TITLE, INTEREST_HEADERS, lambda rows: [to_interest_row(row) for row in rows]),
        (SECURITIES_INTEREST_INCOME_TITLE, INTEREST_HEADERS, lambda rows: [to_interest_row(row) for row in rows]),
        (EQUITY_ITEMS_TITLE, TOTAL_HEADERS, lambda rows: rows),
        (EXPECTED_LOSS_PROVISIONS_TITLE, TOTAL_HEADERS, lambda rows: [to_total_row(row) for row in rows]),
        (STAGE_PROVISION_TITLE, STAGE_PROVISION_HEADERS, lambda rows: rows),
        (LIQUIDITY_COVERAGE_RATIO_TITLE, LIQUIDITY_HEADERS, liquidity_rows),
        (STAGE3_PROVISIONS_TITLE, TOTAL_HEADERS, lambda rows: [to_total_row(row) for row in rows]),
        (OTHER_OPERATING_INCOME_TITLE, TOTAL_HEADERS, lambda rows: [to_total_row(row) for row in rows]),
        (OTHER_OPERATING_EXPENSE_TITLE, OTHER_OPERATING_EXPENSE_HEADERS, lambda rows: rows),
        (FX_RISK_TITLE, FX_RISK_HEADERS, lambda rows: rows),
        (CASH_AND_TCMB_TITLE, INTEREST_HEADERS, lambda rows: [to_interest_row(row) for row in rows]),
        (USED_LOANS_INTEREST_EXPENSE_TITLE, INTEREST_HEADERS, lambda rows: [to_interest_row(row) for row in rows]),
        (LOAN_INTEREST_INCOME_TITLE, INTEREST_HEADERS, lambda rows: [to_interest_row(row) for row in rows]),
        (DEPOSIT_INTEREST_TITLE, ["Kalem", "Vadesiz", "Vadeli", "Toplam"], to_deposit_summary),
        (
            f"{DEPOSIT_INTEREST_TITLE} - toplam vade kirilimi",
            DEPOSIT_TOTAL_HEADERS,
            lambda rows: to_transposed_general_total(
                rows,
                ["Vadesiz Mevduat", "1 Aya Kadar", "3 Aya Kadar", "6 Aya Kadar", "1 Yıla Kadar", "1 Yıldan Uzun", "Birikimli Mevduat", "Toplam"],
            ),
        ),
    ]

    raw_cache: dict[str, list[list[str]]] = {}
    tables: list[tuple[str, list[str], list[list[str]]]] = []
    for title, headers, transformer in specs:
        if title == DEPOSIT_TOTAL_BY_MATURITY_TITLE:
            source_title = DEPOSIT_MATURITY_TITLE
        elif title.startswith(f"{DEPOSIT_INTEREST_TITLE} - "):
            source_title = DEPOSIT_INTEREST_TITLE
        else:
            source_title = title
        try:
            if title == EQUITY_ITEMS_TITLE:
                tables.append((title, headers, extract_equity_rows(page_texts)))
                continue
            if title == RESTRUCTURED_NPL_TITLE:
                tables.append((title, headers, extract_restructured_npl_rows(page_texts)))
                continue
            if title == STAGE_PROVISION_TITLE:
                tables.append((title, headers, extract_stage_provision_rows(page_texts)))
                continue
            if title == LIQUIDITY_COVERAGE_RATIO_TITLE:
                tables.append((title, headers, extract_liquidity_coverage_rows(page_texts)))
                continue
            if title == FX_RISK_TITLE:
                tables.append((title, headers, extract_fx_risk_rows(page_texts)))
                continue
            if title == OTHER_OPERATING_EXPENSE_TITLE:
                tables.append(
                    (
                        title,
                        headers,
                        extract_other_operating_expense_rows_from_words(page_texts, page_words),
                    )
                )
                continue
            if source_title == DEPOSIT_INTEREST_TITLE:
                if source_title not in raw_cache:
                    raw_cache[source_title] = extract_deposit_interest_rows(page_texts)
                tables.append((title, headers, transformer(raw_cache[source_title])))
                continue
            if source_title not in raw_cache:
                raw_cache[source_title] = extract_raw_rows(
                    page_texts,
                    page_tables,
                    source_title,
                    include_empty_amount_rows=source_title == CONSUMER_CREDIT_TITLE,
                )
            tables.append((title, headers, transformer(raw_cache[source_title])))
        except Exception as exc:
            print(f"WARN: dipnot alinamadi: {title} -> {type(exc).__name__}: {exc}", flush=True)
    return tables


def main() -> None:
    parser = argparse.ArgumentParser(description="TEB PDF dipnot tablolarini Garanti mantigiyla ayrisitirir.")
    parser.add_argument("pdf", type=Path, help="Kaynak PDF yolu.")
    args = parser.parse_args()

    tables = extract_all_tables(args.pdf)
    for title, headers, rows in tables:
        print(f"{title}: {len(rows)} satir, {len(headers)} kolon")


if __name__ == "__main__":
    main()
