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
DEPOSIT_MATURITY_TITLE = "Mevduatın vade yapısına ilişkin bilgiler"
DEPOSIT_TOTAL_BY_MATURITY_TITLE = "Mevduatın vade yapısına ilişkin toplam bilgileri"
DEPOSIT_INTEREST_TITLE = "Mevduata ödenen faizin vade yapısına göre gösterimi"
RESTRUCTURED_NPL_TITLE = (
    "Donuk alacaklardan bankaca yeniden yapılandırılan ya da yeni bir itfa planına bağlanan "
    "krediler ve diğer alacaklara ilişkin bilgiler"
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
STAGE_PROVISION_HEADERS = ["Kalem", "Standart Nitelikli Krediler", "Yakın İzlemedeki Krediler"]
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


def _deposit_maturity_rows(
    page_texts: list[str],
    page_tables: list[list[list[list[str | None]]]] | None = None,
) -> list[list[str]]:
    """Normalize QNB's wrapped/abbreviated deposit labels to one stable schema."""
    schema = [
        "Tasarruf Mevduatı",
        "Döviz Tevdiat Hesabı",
        "Yurt İçinde Yer. K.",
        "Yurt Dışında Yer. K.",
        "Resmi Kur. Mevd.",
        "Tic. Kur. Mevd.",
        "Diğ. Kur. Mevd.",
        "Kıymetli Maden DH",
        "Bankalararası Mevduat",
        "T.C. Merkez B.",
        "Yurt İçi Bankalar",
        "Yurt Dışı Bankalar",
        "Katılım Bankaları",
        "Diğer",
        "Toplam",
    ]

    def canonical(label: str) -> str | None:
        key = _base.normalize_key(label)
        if key.startswith("tasarruf mevduati"):
            return schema[0]
        if key.startswith("doviz tevdiat") or key == "hesabi":
            return schema[1]
        if key.startswith("yurt icinde yer"):
            return schema[2]
        if key.startswith("yurt disinda yer"):
            return schema[3]
        if key.startswith("resmi kur"):
            return schema[4]
        if key.startswith("tic kur"):
            return schema[5]
        if key.startswith(("dig kur", "diger kur")):
            return schema[6]
        if "kiymetli maden" in key or key == "dh":
            return schema[7]
        if (
            "bankalararasi mevduat" in key
            or "bankalar arasi mevdu" in key
            or key == "mevduat"
        ):
            return schema[8]
        if key.startswith("t c merkez"):
            return schema[9]
        if key.startswith("yurt ici bankalar"):
            return schema[10]
        if key.startswith("yurt disi bankalar"):
            return schema[11]
        if key.startswith("katilim bankalari"):
            return schema[12]
        if key == "diger":
            return schema[13]
        if key.startswith("toplam"):
            return schema[14]
        return None

    # The ruled-table extractor reads the first maturity bucket as if it were
    # the total in 2026Q2.  The text layer preserves all nine current-period
    # buckets and the final total, so prefer it and stop before Önceki Dönem.
    extracted = _plain_current_rows(
        page_texts,
        DEPOSIT_MATURITY_TITLE,
        ("onceki donem",),
        max_pages=1,
    )
    if not extracted and page_tables:
        extracted = _ORIGINAL_EXTRACT_RAW_ROWS(
            page_texts,
            page_tables,
            DEPOSIT_MATURITY_TITLE,
        )
    by_label: dict[str, list[str]] = {}
    for row in extracted:
        label = canonical(row[0])
        if label is None:
            continue
        by_label[label] = [label, *row[1:]]
    return [by_label.get(label, [label, "", "", ""]) for label in schema]


def _cash_and_tcmb_rows(page_texts: list[str]) -> list[list[str]]:
    rows: list[list[str]] = []
    lines: list[str] = []
    for candidate_title in (
        CASH_AND_TCMB_TITLE,
        "Nakit değerler ve TCMB’ye ilişkin bilgiler",
        "Nakit değerler ve TCMB'ye ilişkin bilgiler",
        "Nakit değerler ve TCMB ye ilişkin bilgiler",
    ):
        lines = _lines_after_title(page_texts, candidate_title, max_pages=1)
        if lines:
            break
    for line in lines:
        key = _base.normalize_key(line)
        if key.startswith("t c merkez bankasi hesabina iliskin bilgiler"):
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
        label = label.replace("đ", "ğ").replace("Đ", "Ğ")
        if _base.normalize_key(label) == "diger":
            label = "Diğer"
        rows.append([label, parsed[1], parsed[2]])
        if _base.normalize_key(label).startswith("toplam"):
            break
    return rows


def _expected_loss_rows_legacy(page_texts: list[str]) -> list[list[str]]:
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


def _expected_loss_rows(page_texts: list[str]) -> list[list[str]]:
    schema = [
        ("beklenen kredi zarari karsiliklari", "Beklenen Kredi Zararı Karşılıkları"),
        ("12 aylik beklenen zarar karsiligi birinci asama", "12 Aylık Beklenen Zarar Karşılığı (Birinci Aşama)"),
        ("kredi riskinde onemli artis ikinci asama", "Kredi Riskinde Önemli Artış (İkinci Aşama)"),
        ("temerrut ucuncu asama", "Temerrüt (Üçüncü Aşama)"),
        ("menkul degerler deger dusus karsiliklari", "Menkul Değerler Değer Düşüş Karşılıkları"),
        (
            "gercege uygun deger farki kar veya zarara yansitilan finansal varliklar",
            "Gerçeğe Uygun Değer Farkı Kâr veya Zarara Yansıtılan Finansal Varlıklar",
        ),
        (
            "gercege uygun deger farki diger kapsamli gelire yansitilan varliklar",
            "Gerçeğe Uygun Değer Farkı Diğer Kapsamlı Gelire Yansıtılan Varlıklar",
        ),
        (
            "istirakler bagli ortakliklar ve birlikte kontrol edilen ortakliklar deger dusus karsiliklari",
            "İştirakler, Bağlı Ortaklıklar ve Birlikte Kontrol Edilen Ortaklıklar Değer Düşüş Karşılıkları",
        ),
        ("istirakler", "İştirakler"),
        ("bagli ortakliklar", "Bağlı Ortaklıklar"),
        ("birlikte kontrol edilen ortakliklar", "Birlikte Kontrol Edilen Ortaklıklar"),
        ("diger", "Diğer"),
        ("toplam", "Toplam"),
    ]

    start_page = next(
        (
            page_no
            for page_no, text in enumerate(page_texts)
            if any(
                _base.normalize_key(_base.normalize_text(line)).startswith(
                    "beklenen kredi zarari karsiliklari"
                )
                for line in text.splitlines()
            )
        ),
        None,
    )
    if start_page is None:
        return []

    mapped: dict[str, list[str]] = {}
    pending_label = ""
    started = False
    for text in page_texts[start_page : min(len(page_texts), start_page + 2)]:
        for raw_line in text.splitlines():
            line = _base.normalize_text(raw_line)
            key = _base.normalize_key(line)
            if not started:
                if not key.startswith("beklenen kredi zarari karsiliklari"):
                    continue
                started = True
            if (
                not line
                or "cari donem" in key
                or "onceki donem" in key
                or line.startswith("Yetkili ")
            ):
                continue

            parts = line.split()
            amounts: list[str] = []
            while parts and _base.SIGNED_AMOUNT_TOKEN_RE.fullmatch(parts[-1]):
                amounts.insert(0, parts.pop())
            label_part = " ".join(parts).strip()
            if not amounts:
                pending_label = " ".join(
                    part for part in (pending_label, label_part) if part
                ).strip()
                continue

            label = " ".join(
                part for part in (pending_label, label_part) if part
            ).strip()
            pending_label = ""
            label_key = _base.normalize_key(label)
            for canonical_key, canonical_label in schema:
                if label_key == canonical_key:
                    mapped[canonical_key] = [canonical_label, *amounts]
                    break
            if label_key == "toplam":
                break
        if "toplam" in mapped:
            break

    return [
        mapped.get(canonical_key, [canonical_label, "-", "-"])
        for canonical_key, canonical_label in schema
    ]


def _npl_movement_rows(page_texts: list[str]) -> list[list[str]]:
    rows: list[list[str]] = []
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
        parsed = _parse_variable_amount_line(line)
        if parsed:
            rows.append([_clean_label(parsed[0], NPL_MOVEMENT_TITLE), *parsed[1:]])

    schema = [
        "Önceki Dönem Sonu Bakiyesi",
        "Dönem İçinde İntikal (+)",
        "Diğer Donuk Alacak Hesaplarından Giriş (+)",
        "Diğer Donuk Alacak Hesaplarına Çıkış (-)",
        "Dönem İçinde Tahsilat (-)",
        "Kayıttan Düşülen (-)",
        "Satılan / Aktiften Silinen (-)",
        "Kurumsal ve Ticari Krediler",
        "Bireysel Krediler",
        "Kredi Kartları",
        "Diğer",
        "Dönem Sonu Bakiyesi",
        "Karşılık (-)",
        "Bilançodaki Net Bakiyesi",
    ]
    mapped: dict[str, list[str]] = {}
    for row in rows:
        key = _base.normalize_key(row[0])
        label: str | None = None
        if key.startswith("onceki donem sonu"):
            label = schema[0]
        elif key.startswith("donem icinde intikal"):
            label = schema[1]
        elif "hesaplarindan giris" in key:
            label = schema[2]
        elif "hesaplarina cikis" in key:
            label = schema[3]
        elif key.startswith("donem icinde tahsilat"):
            label = schema[4]
        elif key.startswith("kayittan dusulen"):
            label = schema[5]
        elif "satilan" in key or "aktiften silinen" in key:
            label = schema[6]
        elif key.startswith("kurumsal ve ticari"):
            label = schema[7]
        elif key.startswith("bireysel krediler"):
            label = schema[8]
        elif key.startswith("kredi kartlari"):
            label = schema[9]
        elif key == "diger":
            label = schema[10]
        elif key.startswith("donem sonu bakiyesi"):
            label = schema[11]
        elif key.startswith("karsilik"):
            label = schema[12]
        elif key.startswith("bilancodaki net"):
            label = schema[13]
        if label is not None:
            values = row[1:4]
            while len(values) < 3:
                values.append("")
            mapped[label] = [label, *values]
    return [mapped.get(label, [label, "", "", ""]) for label in schema]


def _stage3_rows(page_texts: list[str]) -> list[list[str]]:
    schema = [
        "Tahsil İmkanı Sınırlı Krediler ve Diğer Alacaklar İçin Ayrılanlar",
        "Tahsili Şüpheli Krediler ve Diğer Alacaklar İçin Ayrılanlar",
        "Zarar Niteliğindeki Krediler ve Diğer Alacaklar İçin Ayrılanlar",
        "Toplam",
    ]
    mapped: dict[str, list[str]] = {}
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
        # QNB bu tabloda yalnızca "Toplam" sütununu yayımlıyor. İki tutar
        # bekleyen genel ayrıştırıcı, özellikle eski dönemlerde dört satırın
        # tamamını sessizce dışlıyordu.
        parsed = _parse_variable_amount_line(line, min_values=1)
        if parsed:
            label_key = _base.normalize_key(
                _clean_label(parsed[0], STAGE3_PROVISIONS_TITLE)
            )
            canonical: str | None = None
            if label_key.startswith("tahsil imkani sinirli"):
                canonical = schema[0]
            elif label_key.startswith("tahsili supheli"):
                canonical = schema[1]
            elif label_key.startswith("zarar niteligindeki"):
                canonical = schema[2]
            elif label_key.startswith("toplam"):
                canonical = schema[3]
            if canonical is not None:
                # Tabloda tutarlar "Cari Dönem, Önceki Dönem" sırasındadır.
                # Son sayısal değeri seçmek Q2 kolonuna önceki dönem tutarını
                # yazıyordu; Excel tarihçesi için ilk tutar kullanılmalıdır.
                mapped[canonical] = [canonical, parsed[1]]
            if canonical == schema[3]:
                break
    return [mapped.get(label, [label, ""]) for label in schema]


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
        rows = _deposit_maturity_rows(page_texts, page_tables)
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
    if title_key == _base.normalize_key(CASH_AND_TCMB_TITLE):
        rows = _cash_and_tcmb_rows(page_texts)
        if rows:
            return rows
    if title_key == _base.normalize_key(BANKS_INTEREST_INCOME_TITLE):
        rows = _combine_wrapped_rows(_lines_after_title(page_texts, title, max_pages=1, include_title_tail=True))
        if rows:
            for row in rows:
                row[0] = _clean_label(row[0], title)
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
    return _plain_current_rows(page_texts, RESTRUCTURED_NPL_TITLE, ("onceki donem", "1 7 8 2"), max_pages=1)


def _extract_stage_provision_rows_legacy(page_texts: list[str]) -> list[list[str]]:
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


def extract_stage_provision_rows(page_texts: list[str]) -> list[list[str]]:
    labels = {
        "12 aylik beklenen zarar karsiligi": "12 aylık beklenen zarar karşılığı",
        "kredi riskinde onemli artis": "Kredi riskinde önemli artış",
    }
    target_page = next(
        (
            page_no
            for page_no, text in enumerate(page_texts)
            if "standart nitelikli krediler" in _base.normalize_key(text)
            and "yakin izlemedeki krediler" in _base.normalize_key(text)
            and "12 aylik beklenen zarar karsiligi" in _base.normalize_key(text)
            and "kredi riskinde onemli artis" in _base.normalize_key(text)
        ),
        None,
    )
    if target_page is None:
        return []

    mapped: dict[str, list[str]] = {}
    in_current = False
    for raw_line in page_texts[target_page].splitlines():
        line = _base.normalize_text(raw_line)
        key_line = _base.normalize_key(line)
        if "cari donem" in key_line:
            in_current = True
            continue
        if "onceki donem" in key_line and in_current:
            break
        if not in_current:
            continue
        parsed = _parse_amount_line(line)
        if not parsed:
            continue
        key = _base.normalize_key(parsed[0])
        for target_key, label in labels.items():
            if key.startswith(target_key):
                values = parsed[1:3]
                while len(values) < 2:
                    values.append("-")
                mapped[target_key] = [label, *values]
                break

    return [
        mapped.get(target_key, [label, "-", "-"])
        for target_key, label in labels.items()
    ]


def extract_deposit_interest_rows(page_texts: list[str]) -> list[list[str]]:
    rows: list[list[str]] = []
    section = ""
    pending_label = ""
    pending_values: list[str] | None = None
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
            pending_label = ""
            pending_values = None
            continue
        if pending_values is not None:
            if not _is_amount_only_line(line) and not _parse_amount_line(line):
                label = _clean_label(f"{pending_label} {line}".strip(), DEPOSIT_INTEREST_TITLE)
                rows.append([f"{section} - {label}", *pending_values])
                pending_label = ""
                pending_values = None
                continue
            label = _clean_label(pending_label, DEPOSIT_INTEREST_TITLE)
            rows.append([f"{section} - {label}", *pending_values])
            pending_label = ""
            pending_values = None
        if _is_amount_only_line(line) and pending_label:
            pending_values = line.split()
            continue
        parsed = _parse_amount_line(line)
        if not parsed:
            pending_label = " ".join(part for part in [pending_label, line] if part).strip()
            continue
        label = _clean_label(
            " ".join(part for part in [pending_label, parsed[0]] if part),
            DEPOSIT_INTEREST_TITLE,
        )
        pending_label = ""
        if section and not _base.normalize_key(label).startswith(("toplam", "genel toplam")):
            label = f"{section} - {label}"
        values = parsed[1:]
        if len(values) >= 9:
            if re.fullmatch(r"\d", values[-2]) and re.fullmatch(r"\d{2}\.\d{3}\.\d{3}", values[-1]):
                values = [*values[:-2], f"{values[-2]}{values[-1]}"]
            else:
                values = values[:8]
        rows.append([label, *values])
        if _base.normalize_key(label) == "genel toplam":
            break

    schema = [
        "Türk Parası - Bankalararası Mevduat",
        "Türk Parası - Tasarruf Mevduatı",
        "Türk Parası - Resmi Mevduat",
        "Türk Parası - Ticari Mevduat",
        "Türk Parası - Diğer Mevduat",
        "Türk Parası - 7 Gün İhbarlı Mevduat",
        "Türk Parası - Toplam",
        "Yabancı Para - DTH",
        "Yabancı Para - Bankalararası Mevduat",
        "Yabancı Para - 7 Gün İhbarlı Mevduat",
        "Yabancı Para - Kıymetli Maden",
        "Yabancı Para - Toplam",
        "Genel Toplam",
    ]
    mapped: dict[str, list[str]] = {}
    generic_counts = {"turk": 0, "yabanci": 0}
    total_count = 0
    for row in rows:
        key = _base.normalize_key(row[0])
        label: str | None = None
        if key == "genel toplam":
            label = schema[12]
        elif key == "toplam":
            label = schema[6] if total_count == 0 else schema[11]
            total_count += 1
        elif key.startswith("turk parasi"):
            if "tasarruf" in key:
                label = schema[1]
            elif "resmi" in key:
                label = schema[2]
            elif "ticari" in key:
                label = schema[3]
            elif "diger" in key:
                label = schema[4]
            elif "bankalar" in key:
                label = schema[0]
            elif "7 gun" in key:
                label = schema[5]
            elif key == "turk parasi mevduat":
                label = schema[0] if generic_counts["turk"] == 0 else schema[5]
                generic_counts["turk"] += 1
        elif key.startswith("yabanci para"):
            if key.endswith("dth"):
                label = schema[7]
            elif "kiymetli maden" in key:
                label = schema[10]
            elif "bankalar" in key:
                label = schema[8]
            elif "7 gun" in key:
                label = schema[9]
            elif key == "yabanci para mevduat":
                label = schema[8] if generic_counts["yabanci"] == 0 else schema[9]
                generic_counts["yabanci"] += 1
        if label is not None:
            mapped[label] = [label, *row[1:]]
    return [mapped.get(label, [label, "", "", ""]) for label in schema]


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


def _extract_equity_rows_legacy(page_texts: list[str]) -> list[list[str]]:
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


def extract_equity_rows(page_texts: list[str]) -> list[list[str]]:
    """Extract the four stable regulatory-capital totals across QNB layouts."""
    labels = {
        "core": "Çekirdek Sermaye Toplamı",
        "contribution": "Katkı Sermaye Toplamı",
        "total": "Toplam Özkaynak (Ana sermaye ve katkı sermaye toplamı)",
        "risk": "Toplam Risk Ağırlıklı Tutarlar",
    }

    def amount_tokens(line: str) -> list[str]:
        return [
            token
            for token in line.split()
            if token not in {"(1)", "(2)", "(3)"}
            and _base.SIGNED_AMOUNT_TOKEN_RE.fullmatch(token)
        ]

    start_page = next(
        (
            page_no
            for page_no, text in enumerate(page_texts)
            if any(
                _base.normalize_key(_base.normalize_text(line)).startswith(
                    "cekirdek sermaye toplami"
                )
                for line in text.splitlines()
            )
        ),
        None,
    )
    if start_page is None:
        raise RuntimeError("Hedef özkaynak tablosu bulunamadı")

    values: dict[str, str | None] = {
        "core": None,
        "contribution": None,
        "total": None,
        "risk": None,
    }
    contribution_value_pending = False
    for text in page_texts[start_page : min(len(page_texts), start_page + 6)]:
        for raw_line in text.splitlines():
            line = _base.normalize_text(raw_line)
            key = _base.normalize_key(line)
            amounts = amount_tokens(line)
            if contribution_value_pending and amounts:
                values["contribution"] = amounts[0]
                contribution_value_pending = False
                continue
            if key.startswith("cekirdek sermaye toplami") and amounts:
                values["core"] = amounts[0]
            elif key.startswith("katki sermaye toplami"):
                if amounts:
                    values["contribution"] = amounts[0]
                else:
                    contribution_value_pending = True
            elif key.startswith("toplam ozkaynak") and amounts:
                values["total"] = amounts[0]
            elif key.startswith("toplam risk agirlikli tutarlar") and amounts:
                values["risk"] = amounts[0]
                break
        if values["risk"] is not None:
            break

    missing = [labels[name] for name, value in values.items() if value is None]
    if missing:
        raise RuntimeError("Hedef özkaynak kalemleri bulunamadı: " + ", ".join(missing))
    return [
        [labels[name], "", "", values[name]]
        for name in ("core", "contribution", "total", "risk")
    ]


def _extract_liquidity_coverage_rows_legacy(page_texts: list[str]) -> list[list[str]]:
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


def extract_liquidity_coverage_rows(page_texts: list[str]) -> list[list[str]]:
    value_pattern = re.compile(r"%?\d+(?:\.\d{3})*,\d+")
    for text in page_texts:
        page_key = _base.normalize_key(text)
        if "cari donem" not in page_key:
            continue
        for raw_line in text.splitlines():
            line = _base.normalize_text(raw_line)
            key = _base.normalize_key(line)
            if not (
                key.startswith("23 likidite karsilama orani")
                or key.startswith("likidite karsilama orani")
            ):
                continue
            values = value_pattern.findall(line)
            if len(values) >= 2:
                return [
                    [
                        "Likidite Karşılama Oranı (%)",
                        values[-2].lstrip("%"),
                        values[-1].lstrip("%"),
                    ]
                ]
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
