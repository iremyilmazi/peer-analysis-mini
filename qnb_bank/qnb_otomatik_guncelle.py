import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from banka_otomatik_guncelle import main_for_bank

raise SystemExit(main_for_bank("qnb"))
