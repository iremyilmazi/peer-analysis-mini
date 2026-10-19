import http.client
import json
import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import patch

import panel
from tests.helpers import workbook


class PanelTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.root_patch = patch.object(panel,'ROOT',self.root)
        self.root_patch.start()
        panel.DASHBOARD = panel.DashboardState()
        self.server = panel.ThreadingHTTPServer(('127.0.0.1',0),panel.DashboardHandler)
        self.thread = threading.Thread(target=self.server.serve_forever,daemon=True)
        self.thread.start()
        self.addCleanup(self.cleanup)

    def cleanup(self):
        self.server.shutdown(); self.server.server_close(); self.thread.join(timeout=2)
        self.root_patch.stop(); self.temp.cleanup()

    def request(self,method,path,payload=None,headers=None):
        connection = http.client.HTTPConnection('127.0.0.1',self.server.server_address[1],timeout=5)
        request_headers = {'Content-Type':'application/json'}
        request_headers.update(headers or {})
        connection.request(method,path,body=json.dumps(payload) if payload is not None else None,headers=request_headers)
        response = connection.getresponse()
        data = response.read()
        status = response.status
        connection.close()
        return status,json.loads(data)

    def test_thirteen_banks_are_exposed(self):
        code,data = self.request('GET','/api/banks')
        self.assertEqual(200,code)
        self.assertEqual(13,len(data['banks']))

    def test_missing_workbook_blocks_run(self):
        self.assertFalse(self.request('GET','/api/context?bank=ing')[1]['ready'])
        self.assertEqual(400,self.request('POST','/api/run',{'bank':'ing'})[0])

    def test_each_bank_gets_its_own_period(self):
        workbook(self.root/'butunlesik_banka_verisi.xlsx')
        for bank in panel.BANKS:
            context = self.request('GET','/api/context?bank='+bank.key)[1]
            self.assertTrue(context['ready'],bank.key)
            self.assertEqual(bank.key,context['bank'])
            self.assertEqual('31.03.2026',context['latest_period'])
            self.assertEqual('30.06.2026',context['next_period'])

    def test_unknown_bank_is_rejected(self):
        self.assertEqual(400,self.request('POST','/api/run',{'bank':'unknown'})[0])

    def test_default_dry_run_and_concurrency_guard(self):
        workbook(self.root/'butunlesik_banka_verisi.xlsx')
        called = threading.Event()
        def worker(bank,dry_run):
            self.assertEqual('ing',bank.key)
            self.assertTrue(dry_run)
            called.set()
        with patch.object(panel,'_dashboard_worker',side_effect=worker):
            self.assertEqual(202,self.request('POST','/api/run',{'bank':'ing'})[0])
            self.assertTrue(called.wait(timeout=2))
            self.assertEqual(409,self.request('POST','/api/run',{'bank':'akbank'})[0])

    def test_non_boolean_mode_is_rejected(self):
        self.assertEqual(400,self.request('POST','/api/run',{'bank':'ing','dry_run':'false'})[0])

    def test_foreign_origin_or_host_is_rejected(self):
        self.assertEqual(403,self.request('GET','/api/banks',headers={'Host':'example.invalid'})[0])
        self.assertEqual(403,self.request('POST','/api/run',{'bank':'ing'},headers={'Origin':'https://example.invalid'})[0])

    def test_json_array_is_rejected(self):
        self.assertEqual(400,self.request('POST','/api/run',[])[0])

    def test_cannot_open_arbitrary_file(self):
        self.assertEqual(400,self.request('POST','/api/open',{'path':str(self.root/'secret.txt')})[0])

    def test_fixed_elapsed_time_after_completion(self):
        state = panel.DashboardState()
        state.started, state.finished = 10,20
        self.assertEqual(10,state.snapshot()['elapsed'])

    def test_cli_requires_explicit_live(self):
        with patch.object(panel,'validate_installation',return_value=[]),patch.object(panel,'run_bank',return_value=0) as run:
            with patch('sys.argv',['panel.py','--bank','ing']): panel.main()
            self.assertTrue(run.call_args.args[1])
            with patch('sys.argv',['panel.py','--bank','ing','--live']): panel.main()
            self.assertFalse(run.call_args.args[1])
