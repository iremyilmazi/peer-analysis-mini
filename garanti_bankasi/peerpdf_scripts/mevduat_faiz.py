from __future__ import annotations

import argparse
import re
import unicodedata
from pathlib import Path

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


TABLE_TITLE = "Mevduata ödenen faizin vade yapısına göre gösterimi"
MATURITY_TABLE_TITLE = "Mevduatın/ toplanan fonların vade yapısına ilişkin bilgiler"
OUTPUT_FILE = "mevduat_faiz.xlsx"
SHEET_NAME = "garanti"

HEADERS = [
    "{report_date}",
    "Vadesiz Mevduat",
    "1 Aya Kadar",
    "3 Aya Kadar",
    "6 Aya Kadar",
    "1 Yıla Kadar",
    "1 Yıldan Uzun",
    "Birikimli Mevduat",
    "Toplam",
]

MATURITY_HEADERS = [
    "Kalem",
    "Vadesiz",
    "7 Gün İhbarlı",
    "1 Aya Kadar",
    "1-3 Ay",
    "3-6 Ay",
    "6 Ay-1 Yıl",
    "1 Yıl ve Üstü",
    "Birikimli Mevduat",
    "Toplam",
]
MATURITY_START_COL = 11


def normalize_text(value: str) -> str:
    return " ".join(value.split())


def title_in_page(target: str, text: str) -> bool:
    def key(value: str) -> str:
        value = unicodedata.normalize("NFKD", normalize_text(value).casefold().replace("ı", "i"))
        value = "".join(ch for ch in value if not unicodedata.combining(ch))
        return " ".join(re.sub(r"[^a-z0-9]+", " ", value).split())
    target_tokens = key(target).split(); page_tokens = set(key(text).split())
    return key(target) in key(text) or (
        len(target_tokens) >= 3 and all(t in page_tokens for t in target_tokens[:3])
        and sum(t in page_tokens for t in target_tokens) / len(target_tokens) >= 0.50
    )


def find_default_pdf() -> Path:
    pdfs = sorted(Path(".").glob("*.pdf"))
    if not pdfs:
        raise FileNotFoundError("Bu klasörde PDF bulunamadı.")
    if len(pdfs) > 1:
        raise ValueError(
            "Klasörde birden fazla PDF var. Lütfen PDF yolunu argüman olarak verin."
        )
    return pdfs[0]


def extract_report_date(pdf_path: Path) -> str:
    match = re.search(r"(\d{2})\s+(\d{2})\s+(\d{4})", pdf_path.stem)
    if not match:
        return ""
    day, month, year = match.groups()
    return f"{day}.{month}.{year}"


def clean_cell(value: str | None) -> str:
    if value is None:
        return ""
    return normalize_text(value.replace("\n", " "))


def to_excel_number(value):
    if not isinstance(value, str):
        return value
    cleaned = normalize_text(value)
    if re.fullmatch(r"\d+(?:,\d{3})*", cleaned):
        return int(cleaned.replace(",", ""))
    return value


def extract_tables(pdf_path: Path) -> list[tuple[str, list[list[str]]]]:
    target = normalize_text(TABLE_TITLE)

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            if not title_in_page(target, text):
                continue

            extracted: list[tuple[str, list[list[str]]]] = []
            for table in page.extract_tables():
                if not table or not table[0]:
                    continue

                period = clean_cell(table[0][0])
                if period != "Cari Dönem":
                    continue

                rows: list[list[str]] = []
                for raw_row in table[2:]:
                    if not raw_row or not raw_row[0]:
                        continue

                    row = [clean_cell(cell) or "-" for cell in raw_row]
                    rows.append(row)

                if rows:
                    extracted.append((period, rows))

            if extracted:
                return extracted

    raise RuntimeError(f"Hedef tablo PDF içinde bulunamadı: {TABLE_TITLE}")


def extract_maturity_rows(pdf_path: Path) -> list[list[str]]:
    target = normalize_text(MATURITY_TABLE_TITLE)

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            if not title_in_page(target, text):
                continue

            for table in page.extract_tables():
                if not table or not table[0]:
                    continue
                if clean_cell(table[0][0]) != "Cari Dönem":
                    continue

                rows: list[list[str]] = []
                for raw_row in table[1:]:
                    if not raw_row or not raw_row[0]:
                        continue

                    row = [clean_cell(cell) or "-" for cell in raw_row]
                    rows.append(row)

                    if row[0].startswith("Toplam"):
                        return rows

    raise RuntimeError(f"Hedef tablo PDF içinde bulunamadı: {MATURITY_TABLE_TITLE}")


def write_excel(
    tables: list[tuple[str, list[list[str]]]],
    maturity_rows: list[list[str]],
    output_path: Path,
    report_date: str,
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME

    header_fill = PatternFill("solid", fgColor="D9EAF7")

    for period, rows in tables:
        title_row = 1 if ws.max_row == 1 and ws["A1"].value is None else ws.max_row + 1
        ws.cell(title_row, 1, f"{TABLE_TITLE} - {period}")
        ws.cell(title_row, 1).font = Font(bold=True, size=12)
        ws.cell(title_row, 1).alignment = Alignment(wrap_text=True)

        headers = [report_date if header == "{report_date}" else header for header in HEADERS]
        ws.append(headers)
        for cell in ws[title_row + 1]:
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        first_data_row = ws.max_row + 1
        for row in rows:
            ws.append([row[0], *[to_excel_number(value) for value in row[1:]]])

        for row in ws.iter_rows(
            min_row=first_data_row, min_col=2, max_col=len(HEADERS)
        ):
            for cell in row:
                if isinstance(cell.value, int):
                    cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right")

        for row in ws.iter_rows(min_row=first_data_row, min_col=1, max_col=1):
            for cell in row:
                cell.number_format = "@"
                cell.alignment = Alignment(wrap_text=True, vertical="top")

    maturity_title_row = 1
    maturity_header_row = 2
    maturity_first_data_row = 3
    ws.cell(maturity_title_row, MATURITY_START_COL, MATURITY_TABLE_TITLE)
    ws.cell(maturity_title_row, MATURITY_START_COL).font = Font(bold=True, size=12)
    ws.cell(maturity_title_row, MATURITY_START_COL).alignment = Alignment(wrap_text=True)

    for offset, header in enumerate(MATURITY_HEADERS):
        cell = ws.cell(maturity_header_row, MATURITY_START_COL + offset, header)
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_idx, row in enumerate(maturity_rows, start=maturity_first_data_row):
        ws.cell(row_idx, MATURITY_START_COL, row[0])
        ws.cell(row_idx, MATURITY_START_COL).alignment = Alignment(
            wrap_text=True, vertical="top"
        )
        for offset, value in enumerate(row[1:], start=1):
            cell = ws.cell(row_idx, MATURITY_START_COL + offset, to_excel_number(value))
            if isinstance(cell.value, int):
                cell.number_format = "#,##0"
            cell.alignment = Alignment(horizontal="right")

    widths = {
        "A": 28,
        "B": 18,
        "C": 18,
        "D": 18,
        "E": 18,
        "F": 18,
        "G": 18,
        "H": 20,
        "I": 18,
        "K": 32,
        "L": 18,
        "M": 18,
        "N": 18,
        "O": 18,
        "P": 18,
        "Q": 18,
        "R": 18,
        "S": 20,
        "T": 18,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    for col_idx in range(1, MATURITY_START_COL + len(MATURITY_HEADERS)):
        col = get_column_letter(col_idx)
        for cell in ws[col]:
            if isinstance(cell.value, int):
                cell.number_format = "#,##0"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "PDF içindeki mevduata ödenen faizin vade yapısı tablosunu Excel'e "
            "aktarır. Sayılar metin olarak yazılır; ayraçlar korunur."
        )
    )
    parser.add_argument("pdf", nargs="?", type=Path, help="Kaynak PDF yolu")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path(OUTPUT_FILE),
        help="Oluşturulacak Excel dosyası",
    )
    args = parser.parse_args()

    pdf_path = args.pdf or find_default_pdf()
    tables = extract_tables(pdf_path)
    maturity_rows = extract_maturity_rows(pdf_path)
    report_date = extract_report_date(pdf_path)
    write_excel(tables, maturity_rows, args.output, report_date)
    row_count = sum(len(rows) for _, rows in tables)
    print(
        f"{len(tables)} dönem, {row_count} faiz satırı ve "
        f"{len(maturity_rows)} vade satırı yazıldı: {args.output}"
    )


if __name__ == "__main__":
    main()
