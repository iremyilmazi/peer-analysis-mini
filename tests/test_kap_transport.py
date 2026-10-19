import base64
import contextlib
import io
import json
import os
import tempfile
import unittest
import urllib.error
import urllib.request
from pathlib import Path
from unittest.mock import Mock, patch

import kap_transport as transport


class TransportTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.env = patch.dict(os.environ, {}, clear=True)
        self.env.start()
        self.addCleanup(self.env.stop)
        self.addCleanup(transport.close_connection)
        self.driver = self.root / 'msedgedriver.exe'
        self.config = transport.Settings('edge', self.driver, False, None)

    def configure(self, **values):
        (self.root/'kap_connection.json').write_text(json.dumps(values), encoding='utf-8')

    def response(self, data=b'{}', status=200, **overrides):
        result = dict(data=base64.b64encode(data).decode('ascii'), status=status,
                      url='https://www.kap.org.tr/tr/api/example',
                      headers=[['content-disposition', 'attachment; filename="rapor.pdf"']])
        result.update(overrides)
        return result

    def client(self, result=None):
        driver = Mock()
        driver.current_url = 'https://www.kap.org.tr/tr'
        driver.execute_async_script.return_value = self.response() if result is None else result
        client = transport.EdgeClient(self.config)
        factory = patch.object(transport, '_create_driver', return_value=driver)
        factory.start()
        self.addCleanup(factory.stop)
        self.addCleanup(client.close)
        return client, driver

    def test_auto_without_driver_keeps_http(self):
        self.assertEqual('http', transport.settings(self.root).mode)

    def test_auto_detects_root_driver(self):
        self.driver.touch()
        self.assertEqual(self.config, transport.settings(self.root))

    def test_auto_detects_driver_subfolders(self):
        for folder in ('drivers', 'driver'):
            with self.subTest(folder=folder):
                path = self.root/folder/'msedgedriver.exe'
                path.parent.mkdir()
                path.touch()
                self.assertEqual(path, transport.settings(self.root).driver)
                path.unlink()

    def test_relative_path_is_project_relative_not_current_directory(self):
        self.driver.touch()
        self.configure(mode='edge', driver_path='msedgedriver.exe')
        self.assertEqual(self.driver.resolve(), transport.settings(self.root).driver)

    def test_environment_can_select_explicit_local_driver(self):
        self.driver.touch()
        with patch.dict(os.environ, {'KAP_EDGE_DRIVER':str(self.driver), 'KAP_TRANSPORT':'edge'}):
            self.assertEqual(self.config, transport.settings(self.root))

    def test_explicit_missing_driver_does_not_silently_use_http(self):
        self.configure(mode='auto', driver_path='missing.exe')
        with self.assertRaises(FileNotFoundError): transport.settings(self.root)

    def test_edge_mode_requires_local_driver(self):
        self.configure(mode='edge')
        with self.assertRaises(FileNotFoundError): transport.settings(self.root)

    def test_http_can_be_selected_even_with_local_driver(self):
        self.driver.touch()
        self.configure(mode='http')
        self.assertEqual('http', transport.settings(self.root).mode)

    def test_config_rejects_invalid_mode_and_nonboolean_headless(self):
        for values in ({'mode':'chrome'}, {'headless':'false'}, {'driver_path':17}, {'edge_binary':True}):
            self.configure(**values)
            with self.subTest(values=values), self.assertRaises(ValueError): transport.settings(self.root)

    def test_http_path_never_creates_webdriver(self):
        request = urllib.request.Request('https://www.kap.org.tr/tr')
        with patch.object(transport,'settings',return_value=transport.Settings('http',None,False,None)), \
             patch.object(transport.urllib.request,'urlopen',return_value='http') as direct, \
             patch.object(transport,'_create_driver') as factory:
            self.assertEqual('http',transport.urlopen(request,60))
            direct.assert_called_once_with(request,timeout=60)
            factory.assert_not_called()

    def test_selenium_service_is_pinned_to_existing_local_executable(self):
        self.driver.touch()
        with patch.dict(os.environ, {'SE_EDGEDRIVER':'unwanted-driver.exe'}), \
             patch('selenium.webdriver.Edge',return_value=Mock()) as factory:
            transport._create_driver(self.config)
            service = factory.call_args.kwargs['service']
            options = factory.call_args.kwargs['options']
            self.assertEqual(str(self.driver),service.path)
            self.assertEqual(str(self.driver),service.env_path())
            self.assertNotIn('SE_EDGEDRIVER',service.env)
            self.assertFalse(options.accept_insecure_certs)
            self.assertFalse(any('ignore-certificate' in arg or 'disable-web-security' in arg
                                 for arg in options.arguments))

    def test_real_selenium_driver_finder_skips_manager(self):
        from selenium.webdriver.common.driver_finder import DriverFinder
        self.driver.touch()
        with patch('selenium.webdriver.Edge',return_value=Mock()) as factory, \
             patch('selenium.webdriver.common.driver_finder.SeleniumManager') as manager:
            transport._create_driver(self.config)
            finder = DriverFinder(factory.call_args.kwargs['service'],factory.call_args.kwargs['options'])
            self.assertEqual(str(self.driver),finder.get_driver_path())
            manager.assert_not_called()

    def test_missing_executable_cannot_start_selenium_manager(self):
        with patch('selenium.webdriver.Edge') as factory, self.assertRaises(FileNotFoundError):
            transport._create_driver(self.config)
        factory.assert_not_called()

    def test_headless_and_custom_binary_are_explicit(self):
        self.driver.touch()
        binary = self.root/'msedge.exe'
        with patch('selenium.webdriver.Edge',return_value=Mock()) as factory:
            transport._create_driver(transport.Settings('edge',self.driver,True,binary))
            options = factory.call_args.kwargs['options']
            self.assertIn('--headless=new',options.arguments)
            self.assertEqual(str(binary),options.binary_location)

    def test_pdf_bytes_and_filename_survive_browser_transport(self):
        payload = b'%PDF-1.7\n\x00\xff\x80\r\n'
        client, driver = self.client(self.response(payload))
        with contextlib.redirect_stdout(io.StringIO()), client.open(
                urllib.request.Request('https://www.kap.org.tr/tr/api/file'),300) as response:
            self.assertEqual(payload,response.read())
            self.assertIn('rapor.pdf',response.headers['content-disposition'])
        driver.get.assert_called_once_with('https://www.kap.org.tr/tr')

    def test_post_json_uses_browser_and_normalizes_kap_origin(self):
        client, driver = self.client()
        request = urllib.request.Request('https://kap.org.tr/tr/api/example',data=b'{"year":2026}',
                                         headers={'Content-Type':'application/json','User-Agent':'not-forwarded'})
        with contextlib.redirect_stdout(io.StringIO()): client.open(request,60).close()
        args = driver.execute_async_script.call_args.args
        self.assertEqual(('https://www.kap.org.tr/tr/api/example','POST','{"year":2026}'),args[1:4])
        self.assertEqual({'Content-type':'application/json'},args[4])

    def test_connection_reuses_browser_for_metadata_and_attachment(self):
        client, driver = self.client()
        with contextlib.redirect_stdout(io.StringIO()):
            for _ in range(2): client.open(urllib.request.Request('https://www.kap.org.tr/tr/api/file'),60).close()
        driver.get.assert_called_once()
        self.assertEqual(2,driver.execute_async_script.call_count)

    def test_http_429_is_exposed_to_existing_retry_logic(self):
        client, _ = self.client(self.response(b'rate limited',429))
        with contextlib.redirect_stdout(io.StringIO()), self.assertRaises(urllib.error.HTTPError) as raised:
            client.open(urllib.request.Request('https://www.kap.org.tr/tr/api/file'),60)
        self.assertEqual(429,raised.exception.code)
        raised.exception.close()

    def test_foreign_and_insecure_urls_are_rejected_before_start(self):
        client, driver = self.client()
        for url in ('https://example.com/','http://www.kap.org.tr/tr','https://kap.org.tr.evil.com/',
                    'https://user:secret@www.kap.org.tr/','https://www.kap.org.tr:444/'):
            with self.subTest(url=url), self.assertRaises(ValueError):
                client.open(urllib.request.Request(url),60)
        driver.get.assert_not_called()

    def test_foreign_redirect_is_rejected(self):
        client, _ = self.client(self.response(url='https://example.com/login'))
        with contextlib.redirect_stdout(io.StringIO()), self.assertRaises(ValueError):
            client.open(urllib.request.Request('https://www.kap.org.tr/tr/api/file'),60)

    def test_edge_failure_never_falls_back_to_python_http(self):
        with patch.object(transport,'settings',return_value=self.config), \
             patch.object(transport,'_create_driver',side_effect=RuntimeError('driver error')), \
             patch.object(transport.urllib.request,'urlopen') as direct, \
             contextlib.redirect_stdout(io.StringIO()), self.assertRaises(RuntimeError):
            transport.urlopen(urllib.request.Request('https://www.kap.org.tr/tr'),60)
        direct.assert_not_called()

    def test_cleanup_runs_when_download_scope_raises(self):
        with patch.object(transport,'close_connection') as close, self.assertRaises(RuntimeError):
            with transport.connection(): raise RuntimeError('test')
        close.assert_called_once()

    def test_close_is_idempotent(self):
        client, driver = self.client()
        client.driver = driver
        client.close(); client.close()
        driver.quit.assert_called_once()

    def test_both_base_downloaders_use_shared_transport(self):
        from teb_bankasi import teb_kap_history_indir as teb
        from garanti_bankasi import garanti_kap_history_indir as garanti
        for module in (teb,garanti):
            self.assertIs(transport.urlopen,module.urlopen)


if __name__ == '__main__':
    unittest.main()
