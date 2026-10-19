from __future__ import annotations

import argparse
import re
import unicodedata
from pathlib import Path

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


TABLE_TITLE = (
    "Tüketici kredileri, bireysel kredi kartları, personel kredileri "
    "ve personel kredi kartlarına ilişkin bilgiler"
)
INSTALLMENT_COMMERCIAL_TITLE = (
    "Taksitli ticari krediler ve kurumsal kredi kartlarına ilişkin bilgiler"
)
DEPOSIT_MATURITY_TITLE = "Mevduatın/ toplanan fonların vade yapısına ilişkin bilgiler"
DEPOSIT_TOTAL_BY_MATURITY_TITLE = (
    "Mevduatın/ toplanan fonların vade yapısına ilişkin toplam bilgileri"
)
NPL_MOVEMENT_TITLE = "Toplam donuk alacak hareketlerine ilişkin bilgiler"
RESTRUCTURED_NPL_TITLE = (
    "Donuk alacaklardan yeniden yapılandırılan kredilere ilişkin bilgiler"
)
LOAN_INTEREST_INCOME_TITLE = "Kredilerden alınan faiz gelirlerine ilişkin bilgiler"
SECURITIES_INTEREST_INCOME_TITLE = (
    "Menkul değerlerden alınan faiz gelirlerine ilişkin bilgiler"
)
EQUITY_ITEMS_TITLE = "Özkaynak kalemlerine ilişkin bilgiler"
EXPECTED_LOSS_PROVISIONS_TITLE = (
    "Beklenen zarar karşılıkları ve diğer karşılık giderleri"
)
STAGE3_PROVISIONS_TITLE = "Kredilere ilişkin olarak ayrılan 3. aşama karşılıklar"
OTHER_OPERATING_INCOME_TITLE = "Diğer faaliyet gelirlerine ilişkin açıklamalar"
OTHER_OPERATING_EXPENSE_TITLE = "Diğer faaliyet giderlerine ilişkin bilgiler"
FX_RISK_TITLE = "Kur riskine ilişkin açıklamalar"
CASH_AND_TCMB_TITLE = "Nakit değerler ve TCMB’ye ilişkin bilgiler"
USED_LOANS_INTEREST_EXPENSE_TITLE = (
    "Kullanılan kredilere verilen faizlere ilişkin bilgiler"
)
BANKS_INTEREST_INCOME_TITLE = "Bankalardan alınan faiz gelirlerine ilişkin bilgiler"
LIQUIDITY_COVERAGE_RATIO_TITLE = "Likidite karşılama oranı"

CREDIT_TABLE_TITLES = [TABLE_TITLE, INSTALLMENT_COMMERCIAL_TITLE]
CREDIT_HEADERS = ["Kalem", "Kısa Vadeli", "Orta ve Uzun Vadeli", "Toplam"]
DEPOSIT_HEADERS = ["Kalem", "Vadesiz", "Vadeli", "Toplam"]
DEPOSIT_TOTAL_HEADERS = ["Vade", "", "", "Toplam"]
NPL_MOVEMENT_HEADERS = ["Kalem", "III. Grup", "IV. Grup", "V. Grup"]
RESTRUCTURED_NPL_HEADERS = ["Kalem", "III. Grup", "IV. Grup", "V. Grup"]
LOAN_INTEREST_INCOME_HEADERS = ["Kalem", "TP", "YP"]
SECURITIES_INTEREST_INCOME_HEADERS = ["Kalem", "TP", "YP"]
EQUITY_ITEMS_HEADERS = ["Kalem", "", "", "Toplam"]
EXPECTED_LOSS_PROVISIONS_HEADERS = ["Kalem", "", "", "Toplam"]
STAGE3_PROVISIONS_HEADERS = ["Kalem", "", "", "Toplam"]
OTHER_OPERATING_INCOME_HEADERS = ["Kalem", "", "", "Toplam"]
OTHER_OPERATING_EXPENSE_HEADERS = ["Kalem", "", "", "Toplam"]
FX_RISK_HEADERS = ["Kalem", "", "", "Kur"]
CASH_AND_TCMB_HEADERS = ["Kalem", "TP", "YP"]
USED_LOANS_INTEREST_EXPENSE_HEADERS = ["Kalem", "TP", "YP"]
BANKS_INTEREST_INCOME_HEADERS = ["Kalem", "TP", "YP"]
LIQUIDITY_COVERAGE_RATIO_HEADERS = ["Kalem", "TP+YP", "YP"]
EQUITY_TARGET_LABELS = [
    "Çekirdek Sermaye Toplamı",
    "Katkı Sermaye Toplamı",
    "Toplam Özkaynak (Ana sermaye ve katkı sermaye toplamı)",
    "Toplam Risk Ağırlıklı Tutarlar",
]

NUMBER_TOKEN = r"(?:-|\d+(?:,\d{3})*(?:\.\d+)?)"
ROW_RE = re.compile(
    rf"^(?P<label>.+?)\s+(?P<short>{NUMBER_TOKEN})\s+"
    rf"(?P<medium>{NUMBER_TOKEN})\s+(?P<total>{NUMBER_TOKEN})$"
)


def normalize_text(value: str) -> str:
    return " ".join(value.split())


def _title_key(value: str) -> str:
    value = unicodedata.normalize("NFKD", normalize_text(value).casefold().replace("ı", "i"))
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    return " ".join(re.sub(r"[^a-z0-9]+", " ", value).split())


def title_in_page(target: str, page_text: str) -> bool:
    target_key, page_key = _title_key(target), _title_key(page_text)
    if target_key in page_key:
        return True
    tokens = target_key.split()
    if len(tokens) < 3:
        return False
    # Long titles must also match as a phrase. Token-set matching alone can
    # select unrelated segment-report pages containing the same generic words.
    if len(tokens) >= 6 and " ".join(tokens[:6]) not in page_key:
        return False
    page_tokens = set(page_key.split())
    matched = sum(token in page_tokens for token in tokens)
    return all(token in page_tokens for token in tokens[:3]) and matched / len(tokens) >= 0.50


def find_default_pdf() -> Path:
    pdfs = sorted(Path(".").glob("*.pdf"))
    if not pdfs:
        raise FileNotFoundError("Bu klasörde PDF bulunamadı.")
    if len(pdfs) > 1:
        raise ValueError(
            "Klasörde birden fazla PDF var. Lütfen PDF yolunu argüman olarak verin."
        )
    return pdfs[0]


def extract_table_lines(pdf_path: Path, table_title: str) -> list[str]:
    target = normalize_text(table_title)
    lines: list[str] = []
    started = False

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            normalized_page = normalize_text(text)

            if not started:
                if not title_in_page(target, normalized_page):
                    continue
                started = True

            for raw_line in text.splitlines():
                line = normalize_text(raw_line)
                if not line:
                    continue

                if not lines:
                    if line.startswith("Cari Dönem"):
                        lines.append(line)
                    continue

                if line.isdigit():
                    continue

                lines.append(line)
                if line.startswith("Toplam "):
                    return lines

    if not lines:
        raise RuntimeError(f"Hedef tablo PDF içinde bulunamadı: {table_title}")
    return lines


def parse_rows(lines: list[str]) -> list[list[str]]:
    rows: list[list[str]] = []
    in_current_period = False

    for line in lines:
        if line.startswith("Cari Dönem"):
            in_current_period = True
            continue
        if line.startswith("Önceki Dönem"):
            break
        if not in_current_period:
            continue

        match = ROW_RE.match(line)
        if not match:
            continue

        row = [
            match.group("label"),
            match.group("short"),
            match.group("medium"),
            match.group("total"),
        ]
        rows.append(row)
        if normalize_text(row[0]).startswith("Toplam"):
            break

    if not rows:
        raise RuntimeError("Tablo bulundu ancak satırlar ayrıştırılamadı.")
    return rows


def parse_amount(value: str | None) -> int | None:
    if value is None:
        return None
    cleaned = normalize_text(value)
    if cleaned == "-":
        return None
    return int(cleaned.replace(",", ""))


def format_amount(value: int | None) -> str:
    if value is None:
        return "-"
    return f"{value:,}"


def clean_amount_text(value: str | None) -> str:
    if value is None:
        return "-"
    cleaned = normalize_text(value.replace("\n", " "))
    if cleaned == "-":
        return "-"
    digits = "".join(char for char in cleaned if char.isdigit())
    if not digits:
        return "-"
    return f"{int(digits):,}"


def to_excel_number(value):
    if not isinstance(value, str):
        return value
    cleaned = normalize_text(value)
    if re.fullmatch(r"\d+(?:,\d{3})*", cleaned):
        return int(cleaned.replace(",", ""))
    if re.fullmatch(r"\d+\.\d+", cleaned):
        return float(cleaned)
    return value


def apply_numeric_format(cell) -> None:
    if isinstance(cell.value, int):
        cell.number_format = "#,##0"
    elif isinstance(cell.value, float):
        cell.number_format = "0.00"


def extract_deposit_maturity_rows(pdf_path: Path) -> list[list[str]]:
    target = normalize_text(DEPOSIT_MATURITY_TITLE)

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            if not title_in_page(target, text):
                continue

            for table in page.extract_tables():
                if not table or not table[0]:
                    continue
                if normalize_text(table[0][0] or "") != "Cari Dönem":
                    continue

                rows: list[list[str]] = []
                for raw_row in table[1:]:
                    if not raw_row or not raw_row[0]:
                        continue

                    label = normalize_text(raw_row[0].replace("\n", " "))
                    vadesiz = normalize_text(raw_row[1] or "-")
                    toplam = normalize_text(raw_row[-1] or "-")
                    vadesiz_amount = parse_amount(vadesiz)
                    toplam_amount = parse_amount(toplam)
                    vadeli = (
                        None
                        if vadesiz_amount is None or toplam_amount is None
                        else toplam_amount - vadesiz_amount
                    )
                    rows.append([label, vadesiz, format_amount(vadeli), toplam])

                    if label.startswith("Toplam"):
                        return rows

    raise RuntimeError(f"Hedef tablo PDF içinde bulunamadı: {DEPOSIT_MATURITY_TITLE}")


def extract_deposit_total_by_maturity_rows(pdf_path: Path) -> list[list[str]]:
    target = normalize_text(DEPOSIT_MATURITY_TITLE)

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            if not title_in_page(target, text):
                continue

            for table in page.extract_tables():
                if not table or not table[0]:
                    continue
                if normalize_text(table[0][0] or "") != "Cari Dönem":
                    continue

                maturity_headers = [
                    normalize_text((header or "").replace("\n", " "))
                    for header in table[0][1:]
                ]
                for raw_row in table[1:]:
                    if not raw_row or not raw_row[0]:
                        continue
                    label = normalize_text(raw_row[0].replace("\n", " "))
                    if not label.startswith("Toplam"):
                        continue

                    return [
                        [header, "", "", normalize_text(value or "-")]
                        for header, value in zip(maturity_headers, raw_row[1:])
                    ]

    raise RuntimeError(
        f"Hedef toplam satırı PDF içinde bulunamadı: {DEPOSIT_MATURITY_TITLE}"
    )


def extract_npl_movement_rows(pdf_path: Path) -> list[list[str]]:
    target = normalize_text(NPL_MOVEMENT_TITLE)

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            if not title_in_page(target, text):
                continue

            for table in page.extract_tables():
                if not table or not table[0]:
                    continue
                header = [normalize_text(cell or "") for cell in table[0]]
                if header != ["Cari Dönem", "III. Grup", "IV. Grup", "V. Grup"]:
                    continue

                rows: list[list[str]] = []
                for raw_row in table[1:]:
                    if not raw_row or not raw_row[0]:
                        continue
                    label = normalize_text(raw_row[0].replace("\n", " "))
                    if label == "Önceki Dönem":
                        continue
                    rows.append(
                        [
                            label,
                            clean_amount_text(raw_row[1]),
                            clean_amount_text(raw_row[2]),
                            clean_amount_text(raw_row[3]),
                        ]
                    )

                if rows:
                    return rows

    raise RuntimeError(f"Hedef tablo PDF içinde bulunamadı: {NPL_MOVEMENT_TITLE}")


def extract_restructured_npl_rows(pdf_path: Path) -> list[list[str]]:
    target = normalize_text(RESTRUCTURED_NPL_TITLE)

    text_row_re = re.compile(
        rf"^(?P<label>.+?)\s+(?P<a>{NUMBER_TOKEN})\s+(?P<b>{NUMBER_TOKEN})\s+(?P<c>{NUMBER_TOKEN})$"
    )
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            compact = _title_key(text).replace(" ", "")
            if "donukalacak" not in compact or "yenidenyapilandirilan" not in compact:
                continue
            rows: list[list[str]] = []
            active = False
            for raw_line in text.splitlines():
                line = normalize_text(raw_line)
                line_key = _title_key(line).replace(" ", "")
                if line_key == "caridonem":
                    active = True
                    continue
                if active and line_key.startswith("oncekidonem"):
                    break
                if not active:
                    continue
                match = text_row_re.match(line)
                if not match:
                    continue
                label_key = _title_key(match.group("label")).replace(" ", "")
                if "karsiliklardanoncekibruttutarl" in label_key or "yenidenyapilandirilankrediler" in label_key:
                    rows.append([match.group("label"), match.group("a"), match.group("b"), match.group("c")])
            if len(rows) >= 2:
                return [["Cari Dönem", "", "", ""], *rows[:2]]

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            if not title_in_page(target, text):
                continue

            for table in page.extract_tables():
                if not table or not table[0]:
                    continue
                header = [normalize_text(cell or "") for cell in table[0]]
                if header != ["", "III. Grup", "IV. Grup", "V. Grup"]:
                    continue

                rows: list[list[str]] = []
                in_current_period = False
                for raw_row in table[1:]:
                    if not raw_row or not raw_row[0]:
                        continue
                    label = normalize_text(raw_row[0].replace("\n", " "))
                    if label == "Cari Dönem":
                        in_current_period = True
                        rows.append([label, "", "", ""])
                        continue
                    if label == "Önceki Dönem":
                        break
                    if not in_current_period:
                        continue
                    rows.append(
                        [
                            label,
                            clean_amount_text(raw_row[1]),
                            clean_amount_text(raw_row[2]),
                            clean_amount_text(raw_row[3]),
                        ]
                    )

                if rows:
                    return rows

    raise RuntimeError(f"Hedef tablo PDF içinde bulunamadı: {RESTRUCTURED_NPL_TITLE}")


def extract_loan_interest_income_rows(pdf_path: Path) -> list[list[str]]:
    target = normalize_text(LOAN_INTEREST_INCOME_TITLE)

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            if not title_in_page(target, text):
                continue

            for table in page.extract_tables():
                if not table or len(table) < 3:
                    continue
                header = [normalize_text(cell or "") for cell in table[2]]
                if header[:3] != ["", "TP", "YP"]:
                    continue

                rows: list[list[str]] = []
                pending_label = ""
                pending_tp = ""
                pending_yp = ""
                for raw_row in table[3:]:
                    if not raw_row:
                        continue

                    label = normalize_text((raw_row[0] or "").replace("\n", " "))
                    tp = clean_amount_text(raw_row[1])
                    yp = clean_amount_text(raw_row[2])

                    if label and pending_label and raw_row[1] is None and raw_row[2] is None:
                        rows.append([f"{pending_label} {label}", pending_tp, pending_yp])
                        pending_label = ""
                        pending_tp = ""
                        pending_yp = ""
                        continue

                    if label and tp == "-" and yp == "-":
                        pending_label = label
                        pending_tp = ""
                        pending_yp = ""
                        continue

                    if not label and pending_label:
                        pending_tp = tp
                        pending_yp = yp
                        continue

                    pending_label = ""
                    pending_tp = ""
                    pending_yp = ""
                    if not label:
                        continue
                    rows.append([label, tp, yp])

                if rows:
                    return rows

    raise RuntimeError(f"Hedef tablo PDF içinde bulunamadı: {LOAN_INTEREST_INCOME_TITLE}")


def extract_securities_interest_income_rows(pdf_path: Path) -> list[list[str]]:
    target = normalize_text(SECURITIES_INTEREST_INCOME_TITLE)
    expected_label = "Gerçeğe Uygun Değer Farkı Kar veya Zarara"

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            if not title_in_page(target, text):
                continue

            for table in page.extract_tables():
                if not table:
                    continue

                header_idx = None
                for idx, raw_row in enumerate(table):
                    row = [normalize_text(cell or "") for cell in raw_row]
                    if len(row) >= 3 and row[1] == "TP" and row[2] == "YP":
                        header_idx = idx
                        break
                if header_idx is None:
                    continue

                labels = [
                    normalize_text((row[0] or "").replace("\n", " "))
                    for row in table[header_idx + 1 :]
                    if row and row[0]
                ]
                if not any(label.startswith(expected_label) for label in labels):
                    continue

                rows: list[list[str]] = []
                for raw_row in table[header_idx + 1 :]:
                    if not raw_row or not raw_row[0]:
                        continue
                    label = normalize_text(raw_row[0].replace("\n", " "))
                    rows.append(
                        [label, clean_amount_text(raw_row[1]), clean_amount_text(raw_row[2])]
                    )

                if rows:
                    return rows

    raise RuntimeError(
        f"Hedef tablo PDF içinde bulunamadı: {SECURITIES_INTEREST_INCOME_TITLE}"
    )


def extract_equity_item_rows(pdf_path: Path) -> list[list[str]]:
    targets = {normalize_text(label): label for label in EQUITY_TARGET_LABELS}
    rows_by_label: dict[str, list[str]] = {}

    compact_targets = [
        ("cekirdeksermayetoplam", EQUITY_TARGET_LABELS[0]),
        ("katkisermayetoplam", EQUITY_TARGET_LABELS[1]),
        ("toplamozkaynak", EQUITY_TARGET_LABELS[2]),
        ("toplamriskagirliklitutar", EQUITY_TARGET_LABELS[3]),
    ]
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            if page_number > 65:
                break
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            for raw_line in text.splitlines():
                line = normalize_text(raw_line)
                compact = _title_key(line).replace(" ", "")
                amounts = re.findall(r"\d+(?:,\d{3})*", line)
                if not amounts:
                    continue
                for needle, label in compact_targets:
                    # "Toplam Özkaynak (Ana Sermaye ve Katkı Sermaye Toplamı)"
                    # satırı da metin içinde "Katkı Sermaye Toplamı" ifadesini
                    # taşır. Alt dize eşleşmesi katkı sermayeyi toplam özkaynakla
                    # ezdiği için yalnızca gerçek satır başlangıcını kabul et.
                    if compact.startswith(needle):
                        rows_by_label[label] = [label, "", "", amounts[0]]
            if page_number >= 50 and all(label in rows_by_label for label in EQUITY_TARGET_LABELS):
                return [rows_by_label[label] for label in EQUITY_TARGET_LABELS]

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            if page_number > 60:
                break
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            if not any(label in text for label in EQUITY_TARGET_LABELS):
                continue

            previous_line = ""
            for raw_line in text.splitlines():
                line = normalize_text(raw_line)
                for normalized_label, label in targets.items():
                    if not line.startswith(normalized_label):
                        continue
                    amounts = re.findall(NUMBER_TOKEN, line[len(normalized_label) :])
                    if not amounts and label == "Katkı Sermaye Toplamı":
                        amounts = re.findall(NUMBER_TOKEN, previous_line)
                    if not amounts:
                        continue
                    rows_by_label[label] = [label, "", "", clean_amount_text(amounts[0])]
                previous_line = line

            if all(label in rows_by_label for label in EQUITY_TARGET_LABELS):
                return [rows_by_label[label] for label in EQUITY_TARGET_LABELS]

    missing = [label for label in EQUITY_TARGET_LABELS if label not in rows_by_label]
    raise RuntimeError(f"Hedef özkaynak kalemleri bulunamadı: {', '.join(missing)}")


def extract_expected_loss_provision_rows(pdf_path: Path) -> list[list[str]]:
    target = normalize_text(EXPECTED_LOSS_PROVISIONS_TITLE)

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            if not title_in_page(target, text):
                continue

            for table in page.extract_tables():
                if not table or not table[0]:
                    continue
                header = [normalize_text(cell or "") for cell in table[0]]
                if header != ["", "Cari Dönem", "Önceki Dönem"]:
                    continue

                labels = [
                    normalize_text((row[0] or "").replace("\n", " "))
                    for row in table[1:]
                    if row and row[0]
                ]
                if "Beklenen Kredi Zararı Karşılıkları" not in labels:
                    continue

                rows: list[list[str]] = []
                for raw_row in table[1:]:
                    if not raw_row or not raw_row[0]:
                        continue
                    label = normalize_text(raw_row[0].replace("\n", " "))
                    rows.append([label, "", "", clean_amount_text(raw_row[1])])

                if rows:
                    return rows

    raise RuntimeError(
        f"Hedef tablo PDF içinde bulunamadı: {EXPECTED_LOSS_PROVISIONS_TITLE}"
    )


def extract_stage3_provision_rows(pdf_path: Path) -> list[list[str]]:
    target = normalize_text(STAGE3_PROVISIONS_TITLE)
    expected_labels = {
        "Tahsil İmkanı Sınırlı Krediler İçin Ayrılanlar",
        "Tahsili Şüpheli Krediler İçin Ayrılanlar",
        "Zarar Niteliğindeki Krediler İçin Ayrılanlar",
    }

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            if not title_in_page(target, text):
                continue

            for table in page.extract_tables():
                if not table or not table[0]:
                    continue
                header = [normalize_text(cell or "") for cell in table[0]]
                if header != ["", "Cari Dönem", "Önceki Dönem"]:
                    continue

                labels = {
                    normalize_text((row[0] or "").replace("\n", " "))
                    for row in table[1:]
                    if row and row[0]
                }
                if not expected_labels.issubset(labels):
                    continue

                rows: list[list[str]] = []
                for raw_row in table[1:]:
                    if not raw_row or not raw_row[0]:
                        continue
                    label = normalize_text(raw_row[0].replace("\n", " "))
                    rows.append([label, "", "", clean_amount_text(raw_row[1])])

                if rows:
                    return rows

    raise RuntimeError(f"Hedef tablo PDF içinde bulunamadı: {STAGE3_PROVISIONS_TITLE}")


def _canonicalize_other_operating_income_rows(rows: list[list[str]]) -> list[list[str]]:
    seen_asset_sale = False
    canonical: list[list[str]] = []
    for row in rows:
        updated = list(row)
        key = _title_key(updated[0])
        if key.startswith("aktiflerin ") and "satisindan elde edilen gelirler" in key:
            updated[0] = "Aktiflerin Satışından Elde Edilen Gelirler"
            seen_asset_sale = True
        elif key == "diger" and not seen_asset_sale:
            # This is not the table-level "Diğer" line. In Garanti PDFs it
            # is the final child of "Geçmiş Yıllara Ait Gider İptali".
            # Keep the hierarchy explicit in the integrated flat table so
            # users do not add both the subtotal and its children.
            updated[0] = "Geçmiş Yıllara Ait Gider İptali - Diğer"
        elif key == "diger":
            updated[0] = "Diğer"
        canonical.append(updated)
    return canonical


def extract_other_operating_income_rows(pdf_path: Path) -> list[list[str]]:
    target = normalize_text(OTHER_OPERATING_INCOME_TITLE)
    two_period_row_re = re.compile(
        rf"^(?P<label>.+?)\s+(?P<current>{NUMBER_TOKEN})\s+(?P<previous>{NUMBER_TOKEN})$"
    )
    first_label_key = _title_key("Geçmiş Yıllara Ait Gider İptali")

    # In several Garanti reports the visible table is stored as positioned
    # text rather than PDF table objects. Parse those two-period rows directly
    # and keep the current-period amount used by the integrated workbook.
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            if not title_in_page(target, text):
                continue

            rows: list[list[str]] = []
            found_title = False
            found_first_row = False
            for raw_line in text.splitlines():
                line = normalize_text(raw_line)
                if not found_title:
                    if _title_key(target) in _title_key(line):
                        found_title = True
                    continue

                match = two_period_row_re.match(line)
                if not match:
                    continue
                label = normalize_text(match.group("label"))
                label_key = _title_key(label)
                if not found_first_row:
                    if label_key != first_label_key:
                        continue
                    found_first_row = True

                rows.append(
                    [label, "", "", clean_amount_text(match.group("current"))]
                )
                if label_key == "toplam":
                    return _canonicalize_other_operating_income_rows(rows)

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            if not title_in_page(target, text):
                continue

            for table in page.extract_tables():
                if not table or not table[0]:
                    continue
                header = [normalize_text(cell or "") for cell in table[0]]
                if header != ["", "Cari Dönem", "Önceki Dönem"]:
                    continue

                labels = [
                    normalize_text((row[0] or "").replace("\n", " "))
                    for row in table[1:]
                    if row and row[0]
                ]
                if "Geçmiş Yıllara Ait Gider İptali" not in labels:
                    continue

                rows: list[list[str]] = []
                for raw_row in table[1:]:
                    if not raw_row or not raw_row[0]:
                        continue
                    label = normalize_text(raw_row[0].replace("\n", " "))
                    rows.append([label, "", "", clean_amount_text(raw_row[1])])

                if rows:
                    return _canonicalize_other_operating_income_rows(rows)

    raise RuntimeError(f"Hedef tablo PDF içinde bulunamadı: {OTHER_OPERATING_INCOME_TITLE}")


def extract_fx_risk_rows(pdf_path: Path) -> list[list[str]]:
    # Keep the cross-bank row semantics identical while letting Garanti retain
    # its independent PDF opening/extraction pipeline.
    from teb_bankasi.peerpdf_scripts.teb_dipnot import extract_fx_risk_rows as extract_from_text

    with pdfplumber.open(str(pdf_path)) as pdf:
        page_texts = [page.extract_text(x_tolerance=1, y_tolerance=3) or "" for page in pdf.pages]
    return extract_from_text(page_texts)


def extract_other_operating_expense_rows(pdf_path: Path) -> list[list[str]]:
    from teb_bankasi.peerpdf_scripts.teb_dipnot import (
        extract_other_operating_expense_rows as extract_from_text,
        extract_other_operating_expense_rows_from_words as extract_from_words,
    )

    with pdfplumber.open(str(pdf_path)) as pdf:
        page_texts = [page.extract_text(x_tolerance=1, y_tolerance=3) or "" for page in pdf.pages]
        page_words = [
            page.extract_words(keep_blank_chars=False)
            if "diger faaliyet gider" in _title_key(text)
            else []
            for page, text in zip(pdf.pages, page_texts)
        ]
    rows = extract_from_words(page_texts, page_words)

    # Garanti'nin bazı dönemlerinde "Diğer (**)" ile "Toplam" aynı düşey
    # blokta yer alıyor. Konum bazlı ayrıştırıcı bu iki etiketi birleştirip
    # toplam tutarını düşürebiliyor. Toplam satırı ayrı ve sayısal değilse,
    # satır sınırlarını doğru koruyan metin ayrıştırıcısına geri dön.
    total_rows = [
        row for row in rows
        if row and _title_key(row[0]).startswith("toplam")
    ]
    if len(total_rows) != 1 or not clean_amount_text(total_rows[0][-1]):
        rows = extract_from_text(page_texts)
    return rows


def extract_cash_and_tcmb_rows(pdf_path: Path) -> list[list[str]]:
    target = normalize_text(CASH_AND_TCMB_TITLE)
    expected_labels = {"Kasa/Efektif", "TCMB", "Diğer", "Toplam"}

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            if not title_in_page(target, text):
                continue

            for table in page.extract_tables():
                if not table or len(table) < 3:
                    continue
                top_header = [normalize_text(cell or "") for cell in table[0]]
                sub_header = [normalize_text(cell or "") for cell in table[1]]
                if top_header[:3] != ["", "Cari Dönem", ""]:
                    continue
                if sub_header[:3] != ["", "TP", "YP"]:
                    continue

                labels = {
                    normalize_text((row[0] or "").replace("\n", " "))
                    for row in table[2:]
                    if row and row[0]
                }
                if not expected_labels.issubset(labels):
                    continue

                rows: list[list[str]] = []
                for raw_row in table[2:]:
                    if not raw_row or not raw_row[0]:
                        continue
                    label = normalize_text(raw_row[0].replace("\n", " "))
                    rows.append(
                        [
                            label,
                            clean_amount_text(raw_row[1]),
                            clean_amount_text(raw_row[2]),
                        ]
                    )

                if rows:
                    return rows

    raise RuntimeError(f"Hedef tablo PDF içinde bulunamadı: {CASH_AND_TCMB_TITLE}")


def extract_used_loans_interest_expense_rows(pdf_path: Path) -> list[list[str]]:
    target = normalize_text(USED_LOANS_INTEREST_EXPENSE_TITLE)
    expected_labels = {"Bankalara", "Diğer Kuruluşlara", "Toplam"}

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            if not title_in_page(target, text):
                continue

            for table in page.extract_tables():
                if not table:
                    continue

                header_idx = None
                for idx, raw_row in enumerate(table):
                    row = [normalize_text(cell or "") for cell in raw_row]
                    if len(row) >= 3 and row[1] == "TP" and row[2] == "YP":
                        header_idx = idx
                        break
                if header_idx is None:
                    continue

                rows: list[list[str]] = []
                for raw_row in table[header_idx + 1 :]:
                    if not raw_row or not raw_row[0]:
                        continue
                    label = normalize_text(raw_row[0].replace("\n", " "))
                    rows.append(
                        [
                            label,
                            clean_amount_text(raw_row[1]),
                            clean_amount_text(raw_row[2]),
                        ]
                    )
                    if _title_key(label).startswith("toplam"):
                        break

                # The interest-income and interest-expense tables are on the same
                # PDF page and both use TP/YP headers. Accept only the table that
                # contains the expense-specific "Bankalara" and
                # "Diger Kuruluslara" rows; otherwise the first TP/YP table on the
                # page is incorrectly returned.
                row_keys = {
                    _title_key(row[0]).replace(" ", "")
                    for row in rows
                    if row and row[0]
                }
                has_banks_total = "bankalara" in row_keys
                has_other_entities = any(
                    key.startswith("digerkuruluslara") for key in row_keys
                )
                has_total = "toplam" in row_keys
                if rows and has_banks_total and has_other_entities and has_total:
                    return rows

    raise RuntimeError(
        f"Hedef tablo PDF içinde bulunamadı: {USED_LOANS_INTEREST_EXPENSE_TITLE}"
    )


def extract_banks_interest_income_rows(pdf_path: Path) -> list[list[str]]:
    target = normalize_text(BANKS_INTEREST_INCOME_TITLE)
    expected_labels = {"T.C. Merkez Bankasından", "Yurt İçi Bankalardan", "Toplam"}

    expected_compact = {
        "tcmerkezbankasindan", "yurticibankalardan", "yurtdisibankalardan",
        "yurtdisimerkezvesubelerden", "toplam",
    }
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            compact_page = _title_key(text).replace(" ", "")
            if "bankalardanalinanfaizgelirlerine" not in compact_page:
                continue
            rows: list[list[str]] = []
            for raw_line in text.splitlines():
                line = normalize_text(raw_line)
                parts = line.split()
                values: list[str] = []
                while parts and re.fullmatch(NUMBER_TOKEN, parts[-1]):
                    values.insert(0, parts.pop())
                label = " ".join(parts)
                if _title_key(label).replace(" ", "") not in expected_compact or len(values) < 4:
                    continue
                rows.append([label, values[0], values[1]])
            if len(rows) == 5:
                return rows

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            if not title_in_page(target, text):
                continue

            for table in page.extract_tables():
                if not table:
                    continue

                header_idx = None
                for idx, raw_row in enumerate(table):
                    row = [normalize_text(cell or "") for cell in raw_row]
                    if len(row) >= 3 and row[1] == "TP" and row[2] == "YP":
                        header_idx = idx
                        break
                if header_idx is None:
                    continue

                labels = {
                    normalize_text((row[0] or "").replace("\n", " "))
                    for row in table[header_idx + 1 :]
                    if row and row[0]
                }
                if not expected_labels.issubset(labels):
                    continue

                rows: list[list[str]] = []
                for raw_row in table[header_idx + 1 :]:
                    if not raw_row or not raw_row[0]:
                        continue
                    label = normalize_text(raw_row[0].replace("\n", " "))
                    rows.append(
                        [
                            label,
                            clean_amount_text(raw_row[1]),
                            clean_amount_text(raw_row[2]),
                        ]
                    )

                if rows:
                    return rows

            # Older reports expose this table as positioned text rather than an
            # extractable grid. The first two of four values are current TP/YP.
            text_rows: list[list[str]] = []
            started = False
            for raw_line in text.splitlines():
                line = normalize_text(raw_line)
                if not started:
                    if title_in_page(target, line):
                        started = True
                    continue
                if text_rows and "menkul degerlerden alinan" in _title_key(line):
                    break
                parts = line.split()
                values: list[str] = []
                while parts and re.fullmatch(NUMBER_TOKEN, parts[-1]):
                    values.insert(0, parts.pop())
                label = " ".join(parts)
                if len(values) >= 4 and label:
                    text_rows.append([label, values[0], values[1]])
                    if _title_key(label).startswith("toplam"):
                        return text_rows

    raise RuntimeError(f"Hedef tablo PDF içinde bulunamadı: {BANKS_INTEREST_INCOME_TITLE}")


def extract_liquidity_coverage_ratio_rows(pdf_path: Path) -> list[list[str]]:
    target = normalize_text(LIQUIDITY_COVERAGE_RATIO_TITLE).lower()

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            if not title_in_page(target, text):
                continue

            for table in page.extract_tables():
                if not table or not table[0]:
                    continue
                header = [normalize_text(cell or "") for cell in table[0]]
                if not any("Dikkate Alınma Oranı" in cell for cell in header):
                    continue

                for raw_row in table:
                    if not raw_row or len(raw_row) < 6:
                        continue
                    label = normalize_text((raw_row[1] or "").replace("\n", " "))
                    if label == "LİKİDİTE KARŞILAMA ORANI (%)":
                        return [[label, normalize_text(raw_row[4] or "-"), normalize_text(raw_row[5] or "-")]]

    raise RuntimeError(f"Hedef tablo PDF içinde bulunamadı: {LIQUIDITY_COVERAGE_RATIO_TITLE}")


def write_excel(
    tables: list[tuple[str, list[str], list[list[str]]]], output_path: Path
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "garanti"

    header_fill = PatternFill("solid", fgColor="D9EAF7")

    for table_title, headers, rows in tables:
        title_row = 1 if ws.max_row == 1 and ws["A1"].value is None else ws.max_row + 1
        ws.cell(row=title_row, column=1, value=table_title)
        ws.cell(row=title_row, column=1).font = Font(bold=True, size=12)
        ws.cell(row=title_row, column=1).alignment = Alignment(wrap_text=True)

        ws.append(headers)
        for cell in ws[title_row + 1]:
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        first_data_row = ws.max_row + 1
        for row in rows:
            ws.append(
                [
                    row[0],
                    *[to_excel_number(value) for value in row[1:]],
                ]
            )

        for row in ws.iter_rows(min_row=first_data_row, min_col=2, max_col=4):
            for cell in row:
                apply_numeric_format(cell)
                cell.alignment = Alignment(horizontal="right")

        for row in ws.iter_rows(min_row=first_data_row, min_col=1, max_col=1):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")

    widths = {
        "A": 48,
        "B": 18,
        "C": 22,
        "D": 18,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    for col_idx in range(1, 5):
        col = get_column_letter(col_idx)
        for cell in ws[col]:
            apply_numeric_format(cell)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "PDF içindeki cari dönem kredi tablolarını Excel'e aktarır. "
            "Sayılar metin olarak yazılır; ayraçlar korunur."
        )
    )
    parser.add_argument("pdf", nargs="?", type=Path, help="Kaynak PDF yolu")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("garanti_dipnot.xlsx"),
        help="Oluşturulacak Excel dosyası",
    )
    args = parser.parse_args()

    pdf_path = args.pdf or find_default_pdf()
    tables = []
    for title in CREDIT_TABLE_TITLES:
        lines = extract_table_lines(pdf_path, title)
        rows = parse_rows(lines)
        tables.append((title, CREDIT_HEADERS, rows))

    deposit_rows = extract_deposit_maturity_rows(pdf_path)
    tables.append((DEPOSIT_MATURITY_TITLE, DEPOSIT_HEADERS, deposit_rows))

    deposit_total_rows = extract_deposit_total_by_maturity_rows(pdf_path)
    tables.append(
        (DEPOSIT_TOTAL_BY_MATURITY_TITLE, DEPOSIT_TOTAL_HEADERS, deposit_total_rows)
    )

    restructured_npl_rows = extract_restructured_npl_rows(pdf_path)
    tables.append(
        (RESTRUCTURED_NPL_TITLE, RESTRUCTURED_NPL_HEADERS, restructured_npl_rows)
    )

    npl_movement_rows = extract_npl_movement_rows(pdf_path)
    tables.append((NPL_MOVEMENT_TITLE, NPL_MOVEMENT_HEADERS, npl_movement_rows))

    loan_interest_income_rows = extract_loan_interest_income_rows(pdf_path)
    tables.append(
        (
            LOAN_INTEREST_INCOME_TITLE,
            LOAN_INTEREST_INCOME_HEADERS,
            loan_interest_income_rows,
        )
    )

    securities_interest_income_rows = extract_securities_interest_income_rows(pdf_path)
    tables.append(
        (
            SECURITIES_INTEREST_INCOME_TITLE,
            SECURITIES_INTEREST_INCOME_HEADERS,
            securities_interest_income_rows,
        )
    )

    equity_item_rows = extract_equity_item_rows(pdf_path)
    tables.append((EQUITY_ITEMS_TITLE, EQUITY_ITEMS_HEADERS, equity_item_rows))

    expected_loss_provision_rows = extract_expected_loss_provision_rows(pdf_path)
    tables.append(
        (
            EXPECTED_LOSS_PROVISIONS_TITLE,
            EXPECTED_LOSS_PROVISIONS_HEADERS,
            expected_loss_provision_rows,
        )
    )

    stage3_provision_rows = extract_stage3_provision_rows(pdf_path)
    tables.append((STAGE3_PROVISIONS_TITLE, STAGE3_PROVISIONS_HEADERS, stage3_provision_rows))

    other_operating_income_rows = extract_other_operating_income_rows(pdf_path)
    tables.append(
        (
            OTHER_OPERATING_INCOME_TITLE,
            OTHER_OPERATING_INCOME_HEADERS,
            other_operating_income_rows,
        )
    )

    other_operating_expense_rows = extract_other_operating_expense_rows(pdf_path)
    tables.append(
        (
            OTHER_OPERATING_EXPENSE_TITLE,
            OTHER_OPERATING_EXPENSE_HEADERS,
            other_operating_expense_rows,
        )
    )

    fx_risk_rows = extract_fx_risk_rows(pdf_path)
    tables.append((FX_RISK_TITLE, FX_RISK_HEADERS, fx_risk_rows))

    cash_and_tcmb_rows = extract_cash_and_tcmb_rows(pdf_path)
    tables.append((CASH_AND_TCMB_TITLE, CASH_AND_TCMB_HEADERS, cash_and_tcmb_rows))

    used_loans_interest_expense_rows = extract_used_loans_interest_expense_rows(pdf_path)
    tables.append(
        (
            USED_LOANS_INTEREST_EXPENSE_TITLE,
            USED_LOANS_INTEREST_EXPENSE_HEADERS,
            used_loans_interest_expense_rows,
        )
    )

    banks_interest_income_rows = extract_banks_interest_income_rows(pdf_path)
    tables.append(
        (
            BANKS_INTEREST_INCOME_TITLE,
            BANKS_INTEREST_INCOME_HEADERS,
            banks_interest_income_rows,
        )
    )

    liquidity_coverage_ratio_rows = extract_liquidity_coverage_ratio_rows(pdf_path)
    tables.append(
        (
            LIQUIDITY_COVERAGE_RATIO_TITLE,
            LIQUIDITY_COVERAGE_RATIO_HEADERS,
            liquidity_coverage_ratio_rows,
        )
    )

    write_excel(tables, args.output)
    total_rows = sum(len(rows) for _, _, rows in tables)
    print(f"{len(tables)} tablo, {total_rows} satır yazıldı: {args.output}")


if __name__ == "__main__":
    main()
