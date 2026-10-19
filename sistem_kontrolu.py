from __future__ import annotations

import importlib
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
MODULES = ("bs4", "lxml", "openpyxl", "pandas", "pdfplumber", "pypdf")
BANKS = (
    ("akbank", "akbank", "akbank"),
    ("denizbank", "denizbank", "denizbank"),
    ("enpara", "enpara_bankasi", "enpara"),
    ("garanti", "garanti_bankasi", "garanti"),
    ("halkbank", "halkbank", "halkbank"),
    ("hsbc", "hsbc", "hsbc"),
    ("ing", "ing_bank", "ing"),
    ("isbank", "is_bankasi", "is_bankasi"),
    ("qnb", "qnb_bank", "qnb"),
    ("teb", "teb_bankasi", "teb"),
    ("vakifbank", "vakifbank", "vakifbank"),
    ("yapikredi", "yapi_kredi_bankasi", "yapi_kredi"),
    ("ziraat", "ziraat_bankasi", "ziraat"),
)
BANK_EXTRA_FILES = {
    "ing": ("ing_excel_guncelle.py", "ing_guncelle.py"),
}


def check_syntax(path: Path) -> None:
    compile(path.read_text(encoding="utf-8"), str(path), "exec")


def main() -> int:
    errors: list[str] = []
    warnings: list[str] = []
    print(f"Python: {sys.version.split()[0]}")
    if sys.version_info < (3, 11):
        errors.append("Python 3.11 veya daha yeni bir surum gerekli")

    for module_name in MODULES:
        try:
            importlib.import_module(module_name)
            print(f"[OK] Kutuphane: {module_name}")
        except Exception as exc:
            errors.append(f"Kutuphane yuklenemedi: {module_name} ({type(exc).__name__})")

    for filename in ("banka_otomatik_guncelle.py", "entegre_excel_olustur.py", "kap_config.json", "panel.py", "requirements.txt", "banka_paneli/index.html", "banka_paneli/app.js"):
        path = ROOT / filename
        if path.is_file():
            print(f"[OK] Ana dosya: {filename}")
        else:
            errors.append(f"Eksik ana dosya: {filename}")

    for key, folder, prefix in BANKS:
        bank_dir = ROOT / folder
        required = (
            bank_dir / f"{prefix}_otomatik_guncelle.py",
            bank_dir / f"{prefix}_kap_history_indir.py",
            bank_dir / f"{prefix}_dipnot_guncelle.py",
        ) + tuple(bank_dir / filename for filename in BANK_EXTRA_FILES.get(key, ()))
        missing = [item.name for item in required if not item.is_file()]
        parser_files = list((bank_dir / "peerpdf_scripts").glob("*.py"))
        if missing or not parser_files:
            errors.append(f"{key}: eksik banka kaynak dosyasi ({', '.join(missing) or 'parser'})")
            continue
        print(f"[OK] Banka kaynaklari: {key}")

    for path in ROOT.rglob("*.py"):
        if any(part in {'.venv','venv','__pycache__','node_modules'} for part in path.relative_to(ROOT).parts):
            continue
        try:
            check_syntax(path)
        except SyntaxError as exc:
            errors.append(f"Sozdizimi hatasi: {path.relative_to(ROOT)}:{exc.lineno}")
    if not errors:
        print("[OK] Python kaynaklarinin sozdizimi")

    workbook = ROOT / "butunlesik_banka_verisi.xlsx"
    if workbook.is_file():
        print("[OK] Girdi ana Excel mevcut")
    else:
        warnings.append("butunlesik_banka_verisi.xlsx yok; guvenli/gercek guncellemeden once bu klasore kopyalayin")

    if errors:
        print("\nSISTEM HAZIR DEGIL")
        for error in errors:
            print(f"[HATA] {error}")
        return 1

    print("\nSISTEM HAZIR")
    for warning in warnings:
        print(f"[UYARI] {warning}")
    print("Arayuz icin: python panel.py. Ilk islemi guvenli deneme modunda yapin.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
