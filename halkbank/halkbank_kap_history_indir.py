from __future__ import annotations

import argparse
import sys
from datetime import date
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from teb_bankasi import teb_kap_history_indir as _base


HALKBANK_MKK_MEMBER_OID = "1DE05DAA82C3073AE0530A4A622A2EBD"
DEFAULT_OUTPUT_DIR = Path(__file__).resolve().parent / "data" / "history"
DEFAULT_START_YEAR = 2020

def run(start_year: int, output_dir: Path, today: date, force: bool = False) -> list[str]:
    return _base.run(start_year, output_dir, today, force=force, member_oid=HALKBANK_MKK_MEMBER_OID, bank_name="Halkbank")


def run_incremental(output_dir: Path, today: date, force: bool = False) -> list[str]:
    return _base.run_incremental(output_dir, today, force=force, member_oid=HALKBANK_MKK_MEMBER_OID, bank_name="Halkbank")


def main() -> None:
    parser = argparse.ArgumentParser(description="Halkbank KAP finansal rapor history dosyalarini indirir.")
    parser.add_argument("--start-year", type=int, default=DEFAULT_START_YEAR)
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR))
    parser.add_argument("--today")
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--all-history", action="store_true")
    args = parser.parse_args()
    today = date.fromisoformat(args.today) if args.today else date.today()
    if args.all_history:
        run(args.start_year, Path(args.output_dir), today, force=args.force)
    else:
        run_incremental(Path(args.output_dir), today, force=args.force)


if __name__ == "__main__":
    main()
