from __future__ import annotations

import re
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from teb_bankasi.peerpdf_scripts import teb_dipnot as _base

_ORIGINAL_EXTRACT_RAW_ROWS = _base.extract_raw_rows

CREDIT_GROUP_SOURCE_TITLE = (
    "Birinci ve ikinci grup krediler, diğer alacaklar ile yeniden yapılandırılan ya da yeni bir itfa "
    "planına bağlanan krediler ve diğer alacaklara ilişkin bilgiler"
)
RESTRUCTURED_NPL_TITLE = CREDIT_GROUP_SOURCE_TITLE + " - kredi dağılımı"
STAGE_PROVISION_TITLE = CREDIT_GROUP_SOURCE_TITLE + " - beklenen zarar karşılıkları"
SECURITIES_INTEREST_INCOME_TITLE = "Menkul değerlerden alınan faiz gelirlerine ilişkin bilgiler"
LOAN_INTEREST_INCOME_TITLE = "Kredilerden alınan faiz gelirlerine ilişkin bilgiler"
STAGE3_PROVISIONS_TITLE = "Ayrılan özel karşılıklar veya temerrüt (üçüncü aşama) karşılıkları"
CASH_AND_TCMB_TITLE = "Nakit değerler hesabına ilişkin bilgiler"
OTHER_OPERATING_INCOME_TITLE = "Diğer faaliyet gelirlerine ilişkin bilgiler"
NPL_MOVEMENT_TITLE = "Toplam donuk alacak hareketlerine ilişkin bilgiler"

_NPL_ROW_SPECS = [
    (("onceki donem sonu bakiyesi",), "Önceki dönem sonu bakiyesi"),
    (("donem icinde intikal",), "Dönem içinde intikal (+)"),
    (("diger donuk alacak hesaplarina giris",), "Diğer donuk alacak hesaplarına giriş (+)"),
    (("diger donuk alacak hesaplarindan cikis",), "Diğer donuk alacak hesaplarından çıkış (-)"),
    (("donem icinde tahsilat",), "Dönem içinde tahsilat (-)"),
    (("kayittan dusulen", "aktiften silinen"), "Kayıttan düşülen (-)"),
    (("satilan",), "Satılan (-)"),
    (("kurumsal ve ticari krediler",), "Kurumsal ve ticari krediler"),
    (("bireysel krediler",), "Bireysel krediler"),
    (("kredi kartlari",), "Kredi kartları"),
    (("diger",), "Diğer"),
    (("donem sonu bakiyesi",), "Dönem sonu bakiyesi"),
    (("karsilik",), "Karşılık (-)"),
    (("bilancodaki net bakiyesi",), "Bilançodaki net bakiyesi"),
]
_NPL_AMOUNT_RE = re.compile(r"\(\s*\d[\d.,]*\s*\)|(?<!\()-(?!\))|\d[\d.,]*")
_SECURITIES_INTEREST_SPECS = [
    (("gercegeuygundegerfarkikar", "zararayansitilanfinansalvarliklar"), "Gerçeğe Uygun Değer Farkı Kar veya Zarara Yansıtılan Finansal Varlıklar"),
    (("gercegeuygundegerfarkidigerkapsamligelireyansitilanfinansalvarliklar",), "Gerçeğe Uygun Değer Farkı Diğer Kapsamlı Gelire Yansıtılan Finansal Varlıklar"),
    (("itfaedilmismaliyetiuzerindendegerlenenfinansalvarliklar",), "İtfa Edilmiş Maliyeti Üzerinden Değerlenen Finansal Varlıklar"),
    (("toplam",), "Toplam"),
]
_LOAN_INTEREST_SPECS = [
    (("kisa vadeli kredilerden",), "Kısa vadeli kredilerden"),
    (("orta ve uzun vadeli kredilerden",), "Orta ve uzun vadeli kredilerden"),
    (("takipteki alacaklardan alinan faizler",), "Takipteki alacaklardan alınan faizler"),
    (("destekleme fonundan alinan primler",), "Kaynak kul. destekleme fonundan alınan primler"),
]


def _patch_titles() -> None:
    _base.RESTRUCTURED_NPL_TITLE = RESTRUCTURED_NPL_TITLE
    _base.STAGE_PROVISION_TITLE = STAGE_PROVISION_TITLE
    _base.SECURITIES_INTEREST_INCOME_TITLE = SECURITIES_INTEREST_INCOME_TITLE
    _base.STAGE3_PROVISIONS_TITLE = STAGE3_PROVISIONS_TITLE
    _base.CASH_AND_TCMB_TITLE = CASH_AND_TCMB_TITLE
    _base.OTHER_OPERATING_INCOME_TITLE = OTHER_OPERATING_INCOME_TITLE


def extract_ing_credit_group_rows(page_texts: list[str], page_tables: list[list[list[list[str | None]]]]) -> list[list[str]]:
    rows = _ORIGINAL_EXTRACT_RAW_ROWS(page_texts, page_tables, CREDIT_GROUP_SOURCE_TITLE)
    cleaned: list[list[str]] = []
    for row in rows:
        if not row:
            continue
        label_key = _base.normalize_key(row[0])
        if label_key in {"cari donem", "onceki donem"}:
            continue
        cleaned.append(_base.trim_row(row, 4))
    if not cleaned:
        raise RuntimeError("ING birinci ve ikinci grup kredi dağılımı satırları bulunamadı")
    return cleaned


def extract_ing_npl_movement_rows(page_texts: list[str]) -> list[list[str]]:
    for start in _base.title_pages(page_texts, NPL_MOVEMENT_TITLE):
        text = page_texts[start]
        lines = [_base.normalize_text(line) for line in text.splitlines()]
        rows: list[list[str]] = []
        spec_index = 0
        for line in lines:
            key = _base.normalize_key(line)
            if "yabanci para olarak kullandirilan kredilerden kaynaklanan donuk alacaklara" in key:
                break
            if spec_index >= len(_NPL_ROW_SPECS):
                break
            prefixes, canonical = _NPL_ROW_SPECS[spec_index]
            if not any(key.startswith(prefix) for prefix in prefixes):
                continue
            amount_line = re.sub(r"(?<=,)\s+(?=\d)", "", line)
            amounts = _NPL_AMOUNT_RE.findall(amount_line)
            if len(amounts) < 3:
                continue
            rows.append([canonical, *[value.replace(" ", "") for value in amounts[-3:]]])
            spec_index += 1
        if len(rows) == len(_NPL_ROW_SPECS):
            return rows
    raise RuntimeError("ING toplam donuk alacak hareketleri satirlari bulunamadi")


def _compact_key(value: str) -> str:
    return _base.normalize_key(value).replace(" ", "")


def extract_ing_securities_interest_rows(page_texts: list[str]) -> list[list[str]]:
    for text in page_texts:
        compact = _compact_key(text)
        if "itfaedilmismaliyeti" not in compact or "tpyptpyp" not in compact:
            continue
        lines = [_base.normalize_text(line) for line in text.splitlines()]
        rows: list[list[str]] = []
        spec_index = 0
        for index, line in enumerate(lines):
            if spec_index >= len(_SECURITIES_INTEREST_SPECS):
                break
            combined = f"{line} {lines[index + 1] if index + 1 < len(lines) else ''}".strip()
            key = _compact_key(combined)
            required_parts, canonical = _SECURITIES_INTEREST_SPECS[spec_index]
            if spec_index == 3:
                matched = _compact_key(line).startswith("toplam")
                combined = line
            else:
                matched = all(part in key for part in required_parts)
            if not matched:
                continue
            amount_line = re.sub(r"(?<=,)\s+(?=\d)", "", combined)
            amounts = _NPL_AMOUNT_RE.findall(amount_line)
            if len(amounts) < 4:
                continue
            rows.append([canonical, amounts[-4].replace(" ", ""), amounts[-3].replace(" ", "")])
            spec_index += 1
        if len(rows) == len(_SECURITIES_INTEREST_SPECS):
            return rows
    raise RuntimeError("ING menkul deger faiz satirlari bulunamadi")


def extract_ing_loan_interest_rows(page_texts: list[str]) -> list[list[str]]:
    for text in page_texts:
        if "kredilerdenalinanfaizgelirlerineiliskin" not in _compact_key(text):
            continue
        lines = [_base.normalize_text(line) for line in text.splitlines()]
        start = next(i for i, line in enumerate(lines) if "kredilerdenalinanfaizgelirlerineiliskin" in _compact_key(line))
        found: dict[str, list[str]] = {}
        total: list[str] | None = None
        for line in lines[start + 1:]:
            key = _base.normalize_key(line)
            if "bankalardanalinanfaizgelirlerineiliskin" in _compact_key(line):
                break
            amount_line = re.sub(r"(\d[\d,]*,\d{1,2})\s+(\d{1,2})(?=\s|$)", r"\1\2", line)
            amounts = _NPL_AMOUNT_RE.findall(re.sub(r"(?<=,)\s+(?=\d)", "", amount_line))
            if len(amounts) < 2:
                continue
            values = [value.replace(" ", "") for value in amounts[:2]]
            for prefixes, canonical in _LOAN_INTEREST_SPECS:
                if any(prefix in key for prefix in prefixes):
                    found[canonical] = [canonical, *values]
                    break
            if key.startswith("toplam") or key.startswith("kredilerden alinan faizler"):
                total = ["Toplam", *values]
        rows = [found[label] for _prefixes, label in _LOAN_INTEREST_SPECS if label in found]
        if len(rows) == 4 and total is not None:
            return [*rows, total]
    raise RuntimeError("ING kredilerden alinan faiz satirlari bulunamadi")


def extract_ing_stage_provision_rows(page_texts: list[str]) -> list[list[str]]:
    targets = [
        ("12 aylik beklenen zarar", "12 Aylık Beklenen Zarar Karşılığı"),
        ("kredi riskinde onemli artis", "Kredi Riskinde Önemli Artış"),
    ]
    found: dict[str, list[str]] = {}
    pages = _base.title_pages(page_texts, CREDIT_GROUP_SOURCE_TITLE)
    for start in pages:
        for text in page_texts[start:min(len(page_texts), start + 3)]:
            lines = [_base.normalize_text(line) for line in text.splitlines()]
            for index, line in enumerate(lines):
                key = _base.normalize_key(line)
                for target_key, label in targets:
                    if target_key in key and target_key not in found:
                        combined = line
                        parsed = _base.parse_text_line(combined)
                        cursor = index + 1
                        while (not parsed or len(parsed) < 5) and cursor < min(len(lines), index + 4):
                            combined = f"{combined} {lines[cursor]}".strip()
                            parsed = _base.parse_text_line(combined)
                            cursor += 1
                        if not parsed or len(parsed) < 3:
                            continue
                        # Entegre dosyada her rapor için yalnız cari dönem tutulur.
                        values = [_base.clean_cell(v) for v in parsed[1:3]]
                        while len(values) < 2:
                            values.append("")
                        found[target_key] = [label, *values]
        if len(found) == 2:
            break
    if not found:
        raise RuntimeError("ING birinci ve ikinci grup beklenen zarar karşılıkları bulunamadı")
    return [found[key] for key, _label in targets if key in found]


def extract_ing_other_operating_income_rows(page_texts: list[str]) -> list[list[str]]:
    rows: list[list[str]] = []
    for start in _base.title_pages(page_texts, OTHER_OPERATING_INCOME_TITLE):
        started = False
        pending = ""
        for text in page_texts[start:min(len(page_texts), start + 2)]:
            for raw in text.splitlines():
                line = _base.normalize_text(raw)
                key = _base.normalize_key(line)
                if not started:
                    if "diger faaliyet gelirlerine iliskin" in key:
                        started = True
                    continue
                if "beklenen zarar karsiliklari" in key or "diger faaliyet giderlerine" in key:
                    return rows
                if not line or key in {"cari donem onceki donem", "cari donem onceki donem (*)"}:
                    continue
                parsed = _base.parse_text_line(line)
                if parsed and len(parsed) >= 3:
                    label = " ".join(part for part in (pending, parsed[0]) if part).strip()
                    rows.append([label, parsed[1], parsed[2]])
                    pending = ""
                    if _base.normalize_key(label).startswith("toplam"):
                        return rows
                elif not key.startswith(("ing bank", "tutarlar aksi", "iv kar veya zarar")):
                    pending = " ".join(part for part in (pending, line) if part).strip()
        if rows:
            return rows
    raise RuntimeError("ING diğer faaliyet gelirleri satırları bulunamadı")


def extract_raw_rows(page_texts, page_tables, title, include_empty_amount_rows=False):
    if _base.normalize_key(title) == _base.normalize_key(OTHER_OPERATING_INCOME_TITLE):
        return extract_ing_other_operating_income_rows(page_texts)
    if _base.normalize_key(title) == _base.normalize_key(NPL_MOVEMENT_TITLE):
        return extract_ing_npl_movement_rows(page_texts)
    if _base.normalize_key(title) == _base.normalize_key(SECURITIES_INTEREST_INCOME_TITLE):
        return extract_ing_securities_interest_rows(page_texts)
    if _base.normalize_key(title) == _base.normalize_key(LOAN_INTEREST_INCOME_TITLE):
        return extract_ing_loan_interest_rows(page_texts)
    return _ORIGINAL_EXTRACT_RAW_ROWS(page_texts, page_tables, title, include_empty_amount_rows)


def extract_restructured_npl_rows(page_texts: list[str]) -> list[list[str]]:
    specs = [
        ("karsiliklardan onceki brut tutarlar", "Karşılıklardan önceki brüt tutarlar"),
        ("yeniden yapilandirilan krediler", "Yeniden yapılandırılan krediler"),
    ]
    for text in page_texts:
        if "donukalacaklardanbankatarafindanyenidenyapilandirilan" not in _compact_key(text):
            continue
        lines = [_base.normalize_text(line) for line in text.splitlines()]
        start = next(i for i, line in enumerate(lines) if "donukalacaklardanbankatarafindanyenidenyapilandirilan" in _compact_key(line))
        current = next((i for i in range(start + 1, len(lines)) if _base.normalize_key(lines[i]).startswith("cari donem")), None)
        if current is None:
            continue
        found = {}
        for line in lines[current + 1:]:
            key = _base.normalize_key(line)
            if key.startswith("onceki donem") or "toplam donuk alacak hareketlerine" in key:
                break
            amounts = _NPL_AMOUNT_RE.findall(re.sub(r"(?<=,)\s+(?=\d)", "", line))
            if len(amounts) < 3:
                continue
            for prefix, canonical in specs:
                if prefix in key:
                    found[canonical] = [canonical, *[value.replace(" ", "") for value in amounts[:3]]]
        rows = [found[label] for _prefix, label in specs if label in found]
        if len(rows) == 2:
            return rows
    raise RuntimeError("ING yeniden yapılandırılan donuk alacak satırları bulunamadı")


def extract_stage_provision_rows(page_texts: list[str]) -> list[list[str]]:
    return extract_ing_stage_provision_rows(page_texts)


_patch_titles()
_base.extract_restructured_npl_rows = extract_restructured_npl_rows
_base.extract_stage_provision_rows = extract_stage_provision_rows
_base.extract_raw_rows = extract_raw_rows

for _name in dir(_base):
    if not _name.startswith("__") and _name not in globals():
        globals()[_name] = getattr(_base, _name)


def main() -> None:
    _patch_titles()
    _base.extract_restructured_npl_rows = extract_restructured_npl_rows
    _base.extract_stage_provision_rows = extract_stage_provision_rows
    _base.extract_raw_rows = extract_raw_rows
    _base.main()


if __name__ == "__main__":
    main()
