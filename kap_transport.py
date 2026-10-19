"""KAP HTTP transport, optionally through an explicitly local Edge WebDriver.

No Selenium Manager, driver download, certificate override or personal profile.
"""
from __future__ import annotations

import atexit
import base64
import io
import json
import os
import threading
import urllib.error
import urllib.parse
import urllib.request
from contextlib import contextmanager
from dataclasses import dataclass
from email.message import Message
from pathlib import Path

ROOT = Path(__file__).resolve().parent
MAX_RESPONSE_BYTES = 64 * 1024 * 1024
_LOCK = threading.RLock()
_CLIENT = None


@dataclass(frozen=True)
class Settings:
    mode: str
    driver: Path | None
    headless: bool
    binary: Path | None


def _local_path(value: str, root: Path) -> Path:
    path = Path(value).expanduser()
    return (path if path.is_absolute() else root / path).resolve()


def settings(root: Path | None = None) -> Settings:
    root = ROOT if root is None else root.resolve()
    config_file = root / 'kap_connection.json'
    config = json.loads(config_file.read_text(encoding='utf-8-sig')) if config_file.is_file() else {}
    if not isinstance(config, dict):
        raise ValueError('kap_connection.json bir JSON nesnesi olmali.')
    mode = os.environ.get('KAP_TRANSPORT', config.get('mode', 'auto'))
    if not isinstance(mode, str) or mode not in {'auto', 'edge', 'http'}:
        raise ValueError('KAP baglanti modu auto, edge veya http olmali.')
    headless = config.get('headless', False)
    if not isinstance(headless, bool):
        raise ValueError('headless true/false olmali.')
    driver_value = os.environ.get('KAP_EDGE_DRIVER') or config.get('driver_path')
    if driver_value is not None and (not isinstance(driver_value, str) or not driver_value.strip()):
        raise ValueError('driver_path bir dosya yolu veya null olmali.')
    driver = _local_path(driver_value, root) if driver_value else next(
        (path.resolve() for path in (root/'msedgedriver.exe', root/'drivers'/'msedgedriver.exe',
                                     root/'driver'/'msedgedriver.exe') if path.is_file()), None)
    if mode != 'http' and driver_value and not driver.is_file():
        raise FileNotFoundError(f'Yerel EdgeDriver bulunamadi: {driver}')
    if mode == 'auto':
        mode = 'edge' if driver is not None else 'http'
    if mode == 'edge' and (driver is None or not driver.is_file()):
        raise FileNotFoundError('Yerel msedgedriver.exe bulunamadi; proje kokune veya drivers/ klasorune koyun.')
    binary_value = config.get('edge_binary')
    if binary_value is not None and (not isinstance(binary_value, str) or not binary_value.strip()):
        raise ValueError('edge_binary bir dosya yolu veya null olmali.')
    binary = _local_path(binary_value, root) if binary_value else None
    if mode == 'edge' and binary is not None and not binary.is_file():
        raise FileNotFoundError(f'Edge tarayicisi bulunamadi: {binary}')
    return Settings(mode, driver, headless, binary)


def _kap_url(url: str) -> str:
    parsed = urllib.parse.urlsplit(url)
    if (parsed.scheme != 'https' or parsed.hostname not in {'kap.org.tr', 'www.kap.org.tr'}
            or parsed.port not in {None, 443} or parsed.username or parsed.password):
        raise ValueError('Edge baglantisi yalniz HTTPS KAP adreslerine izin verir.')
    return url


def _create_driver(config: Settings):
    try:
        from selenium import webdriver
        from selenium.webdriver.edge.options import Options
        from selenium.webdriver.edge.service import Service
    except ImportError as exc:
        raise RuntimeError('Edge icin Selenium gerekli; requirements.txt ile kutuphaneleri kurun.') from exc
    # Never call webdriver.Edge() without an existing executable. An explicit
    # Service prevents Selenium Manager from searching/downloading a driver.
    if config.driver is None or not config.driver.is_file():
        raise FileNotFoundError('Yerel EdgeDriver dosyasi bulunamadi.')
    options = Options()
    if config.headless:
        options.add_argument('--headless=new')
    if config.binary:
        options.binary_location = str(config.binary)
    service_env = dict(os.environ)
    service_env.pop('SE_EDGEDRIVER', None)
    service_env['MSEDGEDRIVER_TELEMETRY_OPTOUT'] = '1'
    class LocalService(Service):
        def env_path(self):
            # ChromiumDriver re-reads env_path immediately before starting.
            # Never let SE_EDGEDRIVER redirect to a different executable.
            return str(config.driver)

    service = LocalService(executable_path=str(config.driver), env=service_env)
    service.path = str(config.driver)
    try:
        return webdriver.Edge(service=service, options=options)
    except Exception as exc:
        service.stop()
        raise RuntimeError('Yerel EdgeDriver baslatilamadi. Edge/driver surumlerini ve kurum politikasini kontrol edin. '
                           f'{type(exc).__name__}: {exc}') from exc


_FETCH = r"""
const url = arguments[0], method = arguments[1], body = arguments[2];
const headers = arguments[3], timeout = arguments[4], limit = arguments[5];
const done = arguments[arguments.length - 1];
const controller = new AbortController();
const timer = setTimeout(() => controller.abort(), timeout);
fetch(url, {method, headers, body: body === null ? undefined : body,
            credentials:'same-origin', signal:controller.signal})
 .then(async response => {
   const size = Number(response.headers.get('content-length') || 0);
   if (size > limit) throw new Error('KAP yaniti 64 MiB sinirini asiyor.');
   const bytes = new Uint8Array(await response.arrayBuffer());
   if (bytes.length > limit) throw new Error('KAP yaniti 64 MiB sinirini asiyor.');
   const parts = [];
   for (let i=0; i<bytes.length; i+=32768)
     parts.push(String.fromCharCode(...bytes.subarray(i, i+32768)));
   done({status:response.status, url:response.url,
         headers:Array.from(response.headers.entries()), data:btoa(parts.join(''))});
 }).catch(error => done({error:String(error)})).finally(() => clearTimeout(timer));
"""


class EdgeClient:
    def __init__(self, config: Settings):
        self.config = config
        self.driver = None

    def close(self):
        driver, self.driver = self.driver, None
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass

    def open(self, request: urllib.request.Request, timeout: float):
        _kap_url(request.full_url)
        if self.driver is None:
            print(f'KAP baglantisi: yerel EdgeDriver ({self.config.driver})')
            self.driver = _create_driver(self.config)
            try:
                self.driver.set_page_load_timeout(60)
                self.driver.get('https://www.kap.org.tr/tr')
                _kap_url(self.driver.current_url)
            except Exception:
                self.close()
                raise
        self.driver.set_script_timeout(timeout + 15)
        # Normalize kap.org.tr / www.kap.org.tr to the visible KAP origin.
        origin = urllib.parse.urlsplit(_kap_url(self.driver.current_url))
        target = urllib.parse.urlsplit(request.full_url)._replace(netloc=origin.netloc).geturl()
        headers = {key:value for key,value in request.header_items()
                   if key.lower() in {'accept', 'content-type'}}
        body = request.data.decode('utf-8') if request.data is not None else None
        result = self.driver.execute_async_script(_FETCH, target, request.get_method(), body,
                                                 headers, int(timeout * 1000), MAX_RESPONSE_BYTES)
        if not isinstance(result, dict) or result.get('error'):
            raise RuntimeError(f"Edge KAP istegi basarisiz: {result.get('error') if isinstance(result, dict) else 'gecersiz yanit'}")
        _kap_url(result['url'])
        data = base64.b64decode(result['data'], validate=True)
        if len(data) > MAX_RESPONSE_BYTES:
            raise ValueError('KAP yaniti 64 MiB sinirini asiyor.')
        response_headers = Message()
        for name, value in result['headers']:
            response_headers[name] = value
        status = int(result['status'])
        if not 200 <= status < 300:
            raise urllib.error.HTTPError(request.full_url, status, 'Edge KAP HTTP hatasi',
                                         response_headers, io.BytesIO(data))
        response = io.BytesIO(data)
        response.headers = response_headers
        return response


def urlopen(request: urllib.request.Request, timeout: float = 60):
    global _CLIENT
    config = settings()
    if config.mode == 'http':
        return urllib.request.urlopen(request, timeout=timeout)
    with _LOCK:
        if _CLIENT is not None and _CLIENT.config != config:
            close_connection()
        if _CLIENT is None:
            _CLIENT = EdgeClient(config)
        return _CLIENT.open(request, timeout)


def close_connection():
    global _CLIENT
    with _LOCK:
        client, _CLIENT = _CLIENT, None
        if client is not None:
            client.close()


@contextmanager
def connection():
    try:
        yield
    finally:
        close_connection()


atexit.register(close_connection)


def main() -> int:
    import argparse
    parser = argparse.ArgumentParser(description='Excel degistirmeden KAP/yerel EdgeDriver baglantisini sinar.')
    parser.add_argument('--check', action='store_true', required=True)
    parser.parse_args()
    try:
        config = settings()
        print(f'KAP baglanti modu: {config.mode}')
        if config.driver:
            print(f'Yerel driver: {config.driver}')
        with connection():
            with urlopen(urllib.request.Request('https://www.kap.org.tr/tr'), timeout=60) as response:
                if not response.read(1024):
                    raise RuntimeError('KAP bos yanit verdi.')
            print('KAP ana sayfa baglanti kontrolu basarili. Excel degistirilmedi.')
            print('Bu kontrol finansal API/rapor indirme ve veri dogrulama testi degildir.')
        return 0
    except Exception as exc:
        print(f'KAP baglanti kontrolu basarisiz: {type(exc).__name__}: {exc}')
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
