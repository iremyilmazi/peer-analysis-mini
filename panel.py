from __future__ import annotations

import argparse
import importlib
import json
import os
import queue
import re
import subprocess
import sys
import threading
import time
import webbrowser
from dataclasses import dataclass
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from banka_otomatik_guncelle import BANKS as BANK_SPECS


ROOT = Path(__file__).resolve().parent
ANSI = sys.stdout.isatty()


@dataclass(frozen=True)
class BankEntry:
    key: str
    name: str
    script: Path


BANKS = tuple(
    BankEntry(spec.key, spec.display_name, ROOT / spec.folder /
              f"{spec.downloader_module.rsplit('.', 1)[1].removesuffix('_kap_history_indir')}_otomatik_guncelle.py")
    for spec in BANK_SPECS.values()
)

COLORS = {
    "reset": "\033[0m",
    "cyan": "\033[96m",
    "green": "\033[92m",
    "yellow": "\033[93m",
    "red": "\033[91m",
    "dim": "\033[2m",
    "bold": "\033[1m",
}


def color(name: str, text: str) -> str:
    return f"{COLORS[name]}{text}{COLORS['reset']}" if ANSI else text


def clear() -> None:
    os.system("cls" if os.name == "nt" else "clear")


def duration(seconds: float) -> str:
    seconds = max(0, int(seconds))
    minutes, sec = divmod(seconds, 60)
    hours, minutes = divmod(minutes, 60)
    return f"{hours:02d}:{minutes:02d}:{sec:02d}"


def header() -> None:
    print(color("cyan", "=" * 72))
    print(color("bold", "  PEER ANALYSIS  |  KAP BANKA GUNCELLEME MERKEZI"))
    print(color("cyan", "=" * 72))
    print("  PDF + Excel indirme  >  Excel verileri  >  PDF dipnotlari  >  Kontrol")
    print(color("dim", "  Her banka kendi bagimsiz otomatik guncelleme scriptiyle calisir."))
    print()


def validate_installation() -> list[str]:
    errors: list[str] = []
    if sys.version_info < (3, 11):
        errors.append('Python 3.11 veya yenisi gerekli. Test edilen surum Python 3.12.')
    for name in ('bs4','lxml','openpyxl','pandas','pdfplumber','pypdf'):
        try:
            importlib.import_module(name)
        except ImportError:
            errors.append(f'Eksik kutuphane: {name}. README.md kurulum adimlarini uygulayin.')
    for bank in BANKS:
        if not bank.script.is_file():
            errors.append(f"Eksik script: {bank.script}")
    for required in (ROOT / "banka_otomatik_guncelle.py", ROOT / "kap_config.json"):
        if not required.is_file():
            errors.append(f"Eksik ortak dosya: {required}")
    return errors


def workbook_context(bank_key: str = 'akbank') -> dict:
    workbook = ROOT / 'butunlesik_banka_verisi.xlsx'
    try:
        from banka_otomatik_guncelle import BANKS as specs,workbook_periods,period_to_quarter,next_period
        bank = find_bank(bank_key)
        if not workbook.is_file():
            raise ValueError('Güncel Excel’inizi bu klasöre butunlesik_banka_verisi.xlsx adıyla koyun.')
        periods = workbook_periods(workbook,specs[bank.key].company_name)
        if not periods:
            raise ValueError(f'Excel’in {bank.name} sayfasında geçerli bir dönem başlığı yok.')
        latest = max(periods,key=period_to_quarter)
        return {'ready':True,'bank':bank.key,'workbook':workbook.name,'latest_period':latest,'next_period':next_period(latest)}
    except Exception as exc:
        return {'ready':False,'workbook':workbook.name,'error':str(exc)}


def choose_bank() -> BankEntry | None:
    while True:
        clear()
        header()
        for index, bank in enumerate(BANKS, 1):
            print(f"  {color('cyan', f'{index:>2}')}. {bank.name}")
        print(f"  {color('yellow', ' 0')}. Cikis")
        print()
        answer = input("  Banka secimi: ").strip()
        if answer == "0":
            return None
        if answer.isdigit() and 1 <= int(answer) <= len(BANKS):
            return BANKS[int(answer) - 1]


def choose_mode() -> bool:
    print()
    print(f"  {color('cyan', '1')}. Guvenli deneme (ana Excel degismez)")
    print(f"  {color('cyan', '2')}. Gercek guncelleme (yedek alinir, sonra kaydedilir)")
    while True:
        answer = input("  Calisma modu [1]: ").strip() or "1"
        if answer in {"1", "2"}:
            return answer == "1"


def _reader(stream, messages: queue.Queue[str | None]) -> None:
    try:
        for line in iter(stream.readline, ""):
            messages.put(line.rstrip("\r\n"))
    finally:
        messages.put(None)


STEP_RE = re.compile(r"\]\s+([^\s]+)\s+")


class DashboardState:
    def __init__(self) -> None:
        self.lock = threading.Lock()
        self.process: subprocess.Popen[str] | None = None
        self.state = "IDLE"
        self.bank_key = ""
        self.bank_name = ""
        self.step = "Bekleniyor"
        self.progress = 0
        self.started = 0.0
        self.finished = 0.0
        self.last_response = 0.0
        self.logs: list[str] = []
        self.files: list[dict[str, str]] = []
        self.error = ""

    def snapshot(self, cursor: int = 0) -> dict:
        with self.lock:
            now = self.finished or time.monotonic()
            return {
                "state": self.state,
                "bank": self.bank_key,
                "bank_name": self.bank_name,
                "step": self.step,
                "progress": self.progress,
                "elapsed": now - self.started if self.started else 0,
                "last_response_seconds": now - self.last_response if self.last_response else None,
                "logs": self.logs[cursor:],
                "cursor": len(self.logs),
                "files": list(self.files),
                "error": self.error,
            }


DASHBOARD = DashboardState()
PROGRESS_BY_STEP = {
    "Baslangic": 4, "1-KAP": 18, "KAP": 18, "2-EXCEL": 43,
    "3-PDF": 68, "PDF": 68, "3B-BIRLESTIR": 82,
    "4-KONTROL": 94, "Kayit": 98, "Dry-run": 98, "Tamamlandi": 100,
}


def _dashboard_worker(bank: BankEntry, dry_run: bool) -> None:
    audit = ROOT / "outputs" / "otomatik_guncelleme_audit" / f"{bank.key}_panel_{time.time_ns()}.json"
    command = [sys.executable, '-B', "-u", str(bank.script), "--audit-report", str(audit), '--workers', '1']
    if dry_run:
        command.append("--dry-run")
    environment = os.environ.copy()
    environment["PYTHONUNBUFFERED"] = "1"
    environment.setdefault("PYTHONUTF8", "1")
    try:
        process = subprocess.Popen(
            command, cwd=ROOT, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace", bufsize=1, env=environment,
        )
        with DASHBOARD.lock:
            DASHBOARD.process = process
        assert process.stdout is not None
        for raw_line in iter(process.stdout.readline, ""):
            line = raw_line.rstrip("\r\n")
            match = STEP_RE.search(line)
            with DASHBOARD.lock:
                DASHBOARD.logs.append(line)
                DASHBOARD.last_response = time.monotonic()
                if match:
                    DASHBOARD.step = match.group(1)
                    DASHBOARD.progress = PROGRESS_BY_STEP.get(DASHBOARD.step, DASHBOARD.progress)
        return_code = process.wait()
        report: dict = {}
        if audit.is_file():
            try:
                report = json.loads(audit.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                report = {}
        files: list[dict[str, str]] = []
        for key, label in (("output", "Güncellenen Excel"), ("backup", "Güncelleme Öncesi Yedek")):
            value = report.get(key)
            if value:
                files.append({"label": label, "path": str(value)})
        if audit.is_file():
            files.append({"label": "Kontrol Raporu", "path": str(audit)})
        with DASHBOARD.lock:
            DASHBOARD.process = None
            DASHBOARD.files = files
            DASHBOARD.last_response = time.monotonic()
            DASHBOARD.finished = time.monotonic()
            if return_code == 0:
                if report.get('status') == 'NO_NEW_PERIOD':
                    DASHBOARD.state = 'NO_NEW_PERIOD'
                    DASHBOARD.step = 'Yeni dönem yok'
                    DASHBOARD.progress = 0
                    DASHBOARD.logs.append('[SONUÇ] Beklenen sonraki dönem bulunamadı. Ana Excel değişmedi.')
                elif report.get('status') == 'OK':
                    DASHBOARD.state = "OK"
                    DASHBOARD.step = "Tamamlandi"
                    DASHBOARD.progress = 100
                    DASHBOARD.logs.append(f"[BASARILI] {bank.name} akışı tamamlandı.")
                else:
                    DASHBOARD.state = 'ERROR'
                    DASHBOARD.error = 'İşlem tamamlandı ancak geçerli kontrol raporu oluşmadı.'
            else:
                DASHBOARD.state = "ERROR"
                DASHBOARD.error = str(report.get("error") or f"Script hata kodu: {return_code}")
                DASHBOARD.logs.append(f"[HATA] {DASHBOARD.error}")
    except Exception as exc:
        with DASHBOARD.lock:
            DASHBOARD.process = None
            DASHBOARD.state = "ERROR"
            DASHBOARD.error = f"{type(exc).__name__}: {exc}"
            DASHBOARD.last_response = time.monotonic()
            DASHBOARD.finished = time.monotonic()
            DASHBOARD.logs.append(f"[HATA] {DASHBOARD.error}")


class DashboardHandler(BaseHTTPRequestHandler):
    server_version = "PeerAnalysisPanel/1.0"

    def log_message(self, _format: str, *_args) -> None:
        return

    def _json(self, payload: dict, status: int = 200) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _body(self) -> dict:
        length = int(self.headers.get("Content-Length", "0"))
        if length < 0 or length > 64_000:
            raise ValueError("İstek çok büyük")
        if self.headers.get_content_type() != 'application/json':
            raise ValueError('JSON istegi gerekli')
        payload = json.loads(self.rfile.read(length).decode("utf-8") or "{}")
        if not isinstance(payload, dict):
            raise ValueError('JSON nesnesi gerekli')
        return payload

    def _trusted_request(self) -> bool:
        port = self.server.server_address[1]
        hosts = {f'127.0.0.1:{port}', f'localhost:{port}'}
        host = self.headers.get('Host', '')
        origin = self.headers.get('Origin')
        if host not in hosts or (origin and origin != f'http://{host}'):
            self._json({'error':'Yalniz yerel panel istekleri kabul edilir'},403)
            return False
        return True

    def do_GET(self) -> None:
        if not self._trusted_request():
            return
        parsed = urlparse(self.path)
        if parsed.path == "/":
            body = (ROOT / "banka_paneli" / "index.html").read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)
            return
        if parsed.path == '/app.js':
            body = (ROOT / 'banka_paneli' / 'app.js').read_bytes()
            self.send_response(200)
            self.send_header('Content-Type', 'text/javascript; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.send_header('Cache-Control', 'no-store')
            self.end_headers()
            self.wfile.write(body)
            return
        if parsed.path == "/api/banks":
            self._json({"banks": [{"key": item.key, "name": item.name} for item in BANKS]})
            return
        if parsed.path == '/api/context':
            query = parse_qs(parsed.query)
            self._json(workbook_context(query.get('bank', ['akbank'])[0]))
            return
        if parsed.path == "/api/status":
            query = parse_qs(parsed.query)
            try:
                cursor = max(0, int(query.get("cursor", ["0"])[0]))
            except ValueError:
                cursor = 0
            self._json(DASHBOARD.snapshot(cursor))
            return
        self._json({"error": "Bulunamadı"}, 404)

    def do_POST(self) -> None:
        if not self._trusted_request():
            return
        try:
            payload = self._body()
            if self.path == "/api/run":
                bank = find_bank(str(payload.get('bank','')))
                if not isinstance(payload.get('dry_run', True), bool):
                    raise ValueError('dry_run true veya false olmali')
                context = workbook_context(bank.key)
                if not context['ready']:
                    self._json({'error':context['error']},400)
                    return
                with DASHBOARD.lock:
                    if DASHBOARD.state == 'RUNNING':
                        self._json({"error": "Bir çalışma zaten devam ediyor."}, 409)
                        return
                    DASHBOARD.state = "RUNNING"
                    DASHBOARD.bank_key = bank.key
                    DASHBOARD.bank_name = bank.name
                    DASHBOARD.step = "Baslatiliyor"
                    DASHBOARD.progress = 2
                    DASHBOARD.started = DASHBOARD.last_response = time.monotonic()
                    DASHBOARD.finished = 0.0
                    DASHBOARD.logs = [f"[START] {bank.name} | {'Güvenli deneme' if payload.get('dry_run', True) else 'Gerçek güncelleme'}"]
                    DASHBOARD.files = []
                    DASHBOARD.error = ""
                threading.Thread(target=_dashboard_worker, args=(bank, bool(payload.get("dry_run", True))), daemon=True).start()
                self._json(DASHBOARD.snapshot(0), 202)
                return
            if self.path == "/api/stop":
                with DASHBOARD.lock:
                    process = DASHBOARD.process
                if process is not None and process.poll() is None:
                    process.terminate()
                    with DASHBOARD.lock:
                        DASHBOARD.logs.append("[STOP] Kullanıcı çalışmayı durdurdu.")
                self._json({"ok": True})
                return
            if self.path == "/api/open":
                path = Path(str(payload.get("path", ""))).resolve()
                with DASHBOARD.lock:
                    allowed_paths = {Path(item["path"]).resolve() for item in DASHBOARD.files}
                if path not in allowed_paths:
                    raise ValueError("İzin verilmeyen dosya yolu")
                target = path if path.is_dir() else path.parent
                if not target.exists():
                    raise ValueError("Klasör bulunamadı")
                subprocess.Popen(["explorer", str(target)])
                self._json({"ok": True})
                return
            self._json({"error": "Bulunamadı"}, 404)
        except (ValueError, KeyError, json.JSONDecodeError) as exc:
            self._json({"error": str(exc)}, 400)


def serve_dashboard(port: int = 0, open_browser: bool = True) -> int:
    server = ThreadingHTTPServer(("127.0.0.1", port), DashboardHandler)
    address = f"http://127.0.0.1:{server.server_address[1]}"
    print(f"Peer Analysis paneli: {address}", flush=True)
    print("Bu pencereyi kapattığınızda panel durur.", flush=True)
    if open_browser:
        threading.Timer(0.7, lambda: webbrowser.open(address)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0


def run_bank(bank: BankEntry, dry_run: bool, period: str | None = None) -> int:
    command = [sys.executable, "-u", str(bank.script)]
    if dry_run:
        command.append("--dry-run")
    if period:
        command.extend(("--period", period))

    clear()
    header()
    mode = "GUVENLI DENEME" if dry_run else "GERCEK GUNCELLEME"
    print(f"  Banka : {color('bold', bank.name)}")
    print(f"  Mod   : {color('yellow' if dry_run else 'green', mode)}")
    print(f"  Script: {bank.script.relative_to(ROOT)}")
    print(color("cyan", "-" * 72))

    started = last_response = time.monotonic()
    current_step = "Baslatiliyor"
    messages: queue.Queue[str | None] = queue.Queue()
    environment = os.environ.copy()
    environment["PYTHONUNBUFFERED"] = "1"
    environment.setdefault("PYTHONUTF8", "1")
    process = subprocess.Popen(
        command,
        cwd=ROOT,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
        env=environment,
    )
    assert process.stdout is not None
    threading.Thread(target=_reader, args=(process.stdout, messages), daemon=True).start()

    reader_done = False
    status_visible = False
    while not reader_done or process.poll() is None:
        try:
            line = messages.get(timeout=0.25)
            if line is None:
                reader_done = True
                continue
            if status_visible and ANSI:
                print("\r\033[2K", end="")
            last_response = time.monotonic()
            match = STEP_RE.search(line)
            if match:
                current_step = match.group(1)
            elapsed = duration(last_response - started)
            tone = "red" if "HATA" in line or "ERROR" in line else "green" if "Tamamlandi" in line else "cyan"
            print(f"  {color(tone, elapsed)}  {line}")
            status_visible = False
        except queue.Empty:
            now = time.monotonic()
            status = (
                f"  Calisiyor  |  Adim: {current_step:<14} |  Toplam: {duration(now-started)}"
                f"  |  Son yanit: {duration(now-last_response)} once"
            )
            if ANSI:
                print("\r\033[2K" + color("yellow", status), end="", flush=True)
                status_visible = True

    return_code = process.wait()
    if status_visible and ANSI:
        print("\r\033[2K", end="")
    print(color("cyan", "-" * 72))
    total = duration(time.monotonic() - started)
    if return_code == 0:
        print(color("green", f"  BASARILI  |  {bank.name}  |  Toplam sure: {total}"))
    else:
        print(color("red", f"  BASARISIZ |  {bank.name}  |  Kod: {return_code}  |  Sure: {total}"))
    return return_code


def find_bank(key: str) -> BankEntry:
    for bank in BANKS:
        if bank.key == key:
            return bank
    raise ValueError(f"Bilinmeyen banka: {key}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Banka otomatik guncelleme terminal paneli")
    parser.add_argument("--list", action="store_true", help="Banka scriptlerini listele ve cik.")
    parser.add_argument("--bank", choices=[bank.key for bank in BANKS])
    modes = parser.add_mutually_exclusive_group()
    modes.add_argument("--dry-run", action="store_true")
    modes.add_argument('--live', action='store_true', help='--bank ile acikca gercek guncelleme sec.')
    parser.add_argument("--period", help="Ornek: 30.06.2026")
    parser.add_argument("--terminal", action="store_true", help="Tarayici yerine klasik terminal menusunu ac.")
    parser.add_argument("--port", type=int, default=0, help="Yerel panel portu; 0 otomatik secer.")
    parser.add_argument("--no-browser", action="store_true", help="Paneli otomatik olarak tarayicida acma.")
    parser.add_argument('--check', action='store_true', help='Kurulum ve Excel durumunu kontrol edip cik.')
    return parser


def main() -> int:
    args = build_parser().parse_args()
    errors = validate_installation()
    if errors:
        for error in errors:
            print(color("red", error))
        return 2
    if args.check:
        print('KUTUPHANELER VE KAYNAK DOSYALAR HAZIR')
        context = workbook_context(args.bank or 'akbank')
        print(json.dumps(context,ensure_ascii=False,indent=2))
        return 0 if context['ready'] else 1
    if args.list:
        for bank in BANKS:
            print(f"{bank.key}\t{bank.name}\t{bank.script.relative_to(ROOT)}")
        return 0
    if args.bank:
        return run_bank(find_bank(args.bank), not args.live, args.period)

    if not args.terminal:
        return serve_dashboard(args.port, not args.no_browser)

    while True:
        bank = choose_bank()
        if bank is None:
            return 0
        dry_run = choose_mode()
        run_bank(bank, dry_run)
        print()
        input("  Menuye donmek icin Enter'a basin...")


if __name__ == "__main__":
    raise SystemExit(main())
