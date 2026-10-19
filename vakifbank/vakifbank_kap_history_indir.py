from __future__ import annotations

import argparse
import sys
from datetime import date
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from teb_bankasi import teb_kap_history_indir as _base


VAKIFBANK_MKK_MEMBER_OID = "4028e4a1415f4d9b01415fe3340f36e4"
DEFAULT_OUTPUT_DIR = Path(__file__).resolve().parent / "data" / "history"
DEFAULT_START_YEAR = 2020

def run(start_year: int, output_dir: Path, today: date, force: bool = False) -> list[str]:
    return _base.run(start_year, output_dir, today, force=force, member_oid=VAKIFBANK_MKK_MEMBER_OID, bank_name="VakifBank")


def run_incremental(output_dir: Path, today: date, force: bool = False) -> list[str]:
    return _base.run_incremental(output_dir, today, force=force, member_oid=VAKIFBANK_MKK_MEMBER_OID, bank_name="VakifBank")


def main() -> None:
    parser = argparse.ArgumentParser(description="Vakifbank KAP finansal rapor history dosyalarini ceyrek klasorlerine indirir.")
    parser.add_argument("--start-year", type=int, default=DEFAULT_START_YEAR, help="Baslangic yili.")
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR), help="History ana klasoru.")
    parser.add_argument("--today", help="Test icin bugun tarihi, YYYY-MM-DD.")
    parser.add_argument("--force", action="store_true", help="Mevcut Excel/PDF dosyalari olsa bile yeniden indir.")
    parser.add_argument("--all-history", action="store_true", help="Tum history araligini yeniden tara; varsayilan sadece yeni donemleri indirir.")
    args = parser.parse_args()

    today = date.fromisoformat(args.today) if args.today else date.today()
    if args.all_history:
        run(args.start_year, Path(args.output_dir), today, force=args.force)
    else:
        run_incremental(Path(args.output_dir), today, force=args.force)


if __name__ == "__main__":
    main()
