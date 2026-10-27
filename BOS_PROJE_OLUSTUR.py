"""Tek basina calisan, agsiz ve uzerine yazmayan bos proje iskeleti ureticisi."""
from __future__ import annotations

import argparse
from pathlib import Path, PurePosixPath

# Kaynak depo dosya adlarinin sabit anlik goruntusu. Icerik veya indirme yoktur.
FILES = (
    ".gitattributes",
    ".gitignore",
    "BASLAT.cmd",
    "BOS_PROJE_OLUSTUR.py",
    "README.md",
    "akbank/__init__.py",
    "akbank/akbank_dipnot_guncelle.py",
    "akbank/akbank_kap_history_indir.py",
    "akbank/akbank_otomatik_guncelle.py",
    "akbank/peerpdf_scripts/akbank_dipnot.py",
    "banka_otomatik_guncelle.py",
    "banka_paneli/app.js",
    "banka_paneli/index.html",
    "denizbank/__init__.py",
    "denizbank/denizbank_dipnot_guncelle.py",
    "denizbank/denizbank_kap_history_indir.py",
    "denizbank/denizbank_otomatik_guncelle.py",
    "denizbank/peerpdf_scripts/denizbank_dipnot.py",
    "docs/BOS_PROJE_ISKELETI.md",
    "docs/CEO_Handnotes_Tum_Sunum.md",
    "docs/CMD_ILE_KURULUM_VE_BASLATMA.md",
    "docs/DENIZBANK_GIDER_DUZELTMESI.md",
    "docs/EDGE_DRIVER_KURULUMU.md",
    "docs/GITHUB_HAZIRLIK.md",
    "docs/MIMARI_VE_KONTROLLER.md",
    "docs/TEST_DURUMU.md",
    "enpara_bankasi/__init__.py",
    "enpara_bankasi/enpara_dipnot_guncelle.py",
    "enpara_bankasi/enpara_kap_history_indir.py",
    "enpara_bankasi/enpara_otomatik_guncelle.py",
    "enpara_bankasi/peerpdf_scripts/enpara_dipnot.py",
    "entegre_excel_olustur.py",
    "garanti_bankasi/__init__.py",
    "garanti_bankasi/garanti_dipnot_guncelle.py",
    "garanti_bankasi/garanti_kap_history_indir.py",
    "garanti_bankasi/garanti_otomatik_guncelle.py",
    "garanti_bankasi/peerpdf_scripts/birinci_ikinci_asama_krediler.py",
    "garanti_bankasi/peerpdf_scripts/garanti_dipnot.py",
    "garanti_bankasi/peerpdf_scripts/mevduat_faiz.py",
    "halkbank/__init__.py",
    "halkbank/halkbank_dipnot_guncelle.py",
    "halkbank/halkbank_kap_history_indir.py",
    "halkbank/halkbank_otomatik_guncelle.py",
    "halkbank/peerpdf_scripts/halkbank_dipnot.py",
    "hsbc/__init__.py",
    "hsbc/hsbc_dipnot_guncelle.py",
    "hsbc/hsbc_kap_history_indir.py",
    "hsbc/hsbc_otomatik_guncelle.py",
    "hsbc/peerpdf_scripts/hsbc_dipnot.py",
    "ing_bank/__init__.py",
    "ing_bank/ing_dipnot_guncelle.py",
    "ing_bank/ing_excel_guncelle.py",
    "ing_bank/ing_guncelle.py",
    "ing_bank/ing_kap_history_indir.py",
    "ing_bank/ing_otomatik_guncelle.py",
    "ing_bank/peerpdf_scripts/ing_dipnot.py",
    "is_bankasi/__init__.py",
    "is_bankasi/is_bankasi_dipnot_guncelle.py",
    "is_bankasi/is_bankasi_kap_history_indir.py",
    "is_bankasi/is_bankasi_otomatik_guncelle.py",
    "is_bankasi/peerpdf_scripts/is_bankasi_dipnot.py",
    "kap_config.json",
    "kap_connection.json",
    "kap_transport.py",
    "panel.py",
    "qnb_bank/__init__.py",
    "qnb_bank/peerpdf_scripts/qnb_dipnot.py",
    "qnb_bank/qnb_dipnot_guncelle.py",
    "qnb_bank/qnb_kap_history_indir.py",
    "qnb_bank/qnb_otomatik_guncelle.py",
    "requirements.txt",
    "sistem_kontrolu.py",
    "teb_bankasi/__init__.py",
    "teb_bankasi/peerpdf_scripts/teb_dipnot.py",
    "teb_bankasi/teb_dipnot_guncelle.py",
    "teb_bankasi/teb_kap_history_indir.py",
    "teb_bankasi/teb_otomatik_guncelle.py",
    "tests/__init__.py",
    "tests/helpers.py",
    "tests/test_denizbank_expense.py",
    "tests/test_engine.py",
    "tests/test_ing_excel_guncelle.py",
    "tests/test_kap_transport.py",
    "tests/test_panel.py",
    "tests/test_scaffold.py",
    "vakifbank/__init__.py",
    "vakifbank/peerpdf_scripts/vakifbank_dipnot.py",
    "vakifbank/vakifbank_dipnot_guncelle.py",
    "vakifbank/vakifbank_kap_history_indir.py",
    "vakifbank/vakifbank_otomatik_guncelle.py",
    "yapi_kredi_bankasi/__init__.py",
    "yapi_kredi_bankasi/peerpdf_scripts/yapi_kredi_dipnot.py",
    "yapi_kredi_bankasi/yapi_kredi_dipnot_guncelle.py",
    "yapi_kredi_bankasi/yapi_kredi_kap_history_indir.py",
    "yapi_kredi_bankasi/yapi_kredi_otomatik_guncelle.py",
    "ziraat_bankasi/__init__.py",
    "ziraat_bankasi/peerpdf_scripts/ziraat_dipnot.py",
    "ziraat_bankasi/ziraat_dipnot_guncelle.py",
    "ziraat_bankasi/ziraat_kap_history_indir.py",
    "ziraat_bankasi/ziraat_otomatik_guncelle.py",
)


def _is_link(path: Path) -> bool:
    return path.is_symlink() or getattr(path, "is_junction", lambda: False)()


def _checked_paths(files) -> tuple[PurePosixPath, ...]:
    paths = []
    seen = set()
    for value in files:
        if not isinstance(value, str) or not value or "\\" in value or ":" in value:
            raise ValueError(f"Gecersiz goreli dosya yolu: {value!r}")
        path = PurePosixPath(value)
        if path.is_absolute() or any(part in {"", ".", ".."} for part in value.split("/")):
            raise ValueError(f"Hedef disina cikan/gecersiz yol: {value!r}")
        key = value.casefold()
        if key in seen:
            raise ValueError(f"Tekrarlanan dosya yolu: {value}")
        seen.add(key)
        paths.append(path)
    for path in paths:
        if any(parent.as_posix().casefold() in seen for parent in path.parents if str(parent) != "."):
            raise ValueError(f"Dosya/klasor cakismasi: {path}")
    return tuple(paths)


def _check_location(path: Path) -> None:
    for item in (path, *path.parents):
        if _is_link(item):
            raise ValueError(f"Link/junction uzerinden dosya olusturulmaz: {item}")


def create_empty_project(target: Path, files=FILES) -> tuple[int, int]:
    paths = _checked_paths(files)
    root = Path(target).absolute()
    _check_location(root)
    root = root.resolve()
    if root == root.parent:
        raise ValueError("Disk/dosya sistemi koku hedef olarak kullanilamaz.")
    if root.exists() and not root.is_dir():
        raise ValueError(f"Hedef bir klasor degil: {root}")
    # Tum cakismalari yazmadan once kontrol et. Mevcut dosyalar silinmez,
    # bosaltilmaz, yeniden adlandirilmaz veya icerikleri okunmaz.
    for relative in paths:
        destination = root.joinpath(*relative.parts)
        _check_location(destination)
        if destination.exists() and not destination.is_file():
            raise ValueError(f"Dosya adinda bir klasor/ozel nesne var: {destination}")
        for parent in destination.parents:
            if parent == root.parent:
                break
            if parent.exists() and not parent.is_dir():
                raise ValueError(f"Klasor adinda bir dosya var: {parent}")
    root.mkdir(parents=True, exist_ok=True)
    created = skipped = 0
    for relative in paths:
        destination = root.joinpath(*relative.parts)
        _check_location(destination)
        destination.parent.mkdir(parents=True, exist_ok=True)
        _check_location(destination)
        try:
            # Exclusive creation: never truncate or overwrite an existing file.
            with destination.open("xb"):
                pass
            created += 1
        except FileExistsError:
            _check_location(destination)
            if not destination.is_file():
                raise ValueError(f"Dosya/klasor cakismasi: {destination}")
            skipped += 1
    return created, skipped


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Projenin dosya/klasor adlarini bos olarak olusturur. Kod indirmez.")
    parser.add_argument("--hedef", type=Path,
                        default=Path(__file__).resolve().parent / "peer-analysis-mini-bos",
                        help="Hedef klasor; varsayilan bu scriptin yanindaki peer-analysis-mini-bos.")
    parser.add_argument("--listele", action="store_true", help="Yalniz dosya adlarini gosterir; klasor/dosya olusturmaz.")
    args = parser.parse_args(argv)
    if args.listele:
        for name in FILES:
            print(name)
        print(f"Toplam: {len(FILES)} dosya")
        return 0
    try:
        created, skipped = create_empty_project(args.hedef)
    except (OSError, ValueError) as exc:
        print(f"HATA: {exc}")
        print("Mevcut dosyalarin uzerine yazilmadi. Kesinti olustuysa eksik dosyalar icin tekrar calistirabilirsiniz.")
        return 1
    print(f"Hedef: {args.hedef.resolve()}")
    print(f"Olusturulan bos dosya: {created}; mevcut oldugu icin atlanan: {skipped}")
    print("Bu yalniz bir ISKELETTIR. Yeni dosyalar 0 bayttir; uygulama kodu/ayar/Excel/driver icermez.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

