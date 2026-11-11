import contextlib
import io
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

import yan_excel_otomatik_guncelle as side
from yan_excel_donem_ekle import payloads_for_new_periods, workbook_periods


class SideExcelTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.book = self.root / 'birinci_ikinci_asama_krediler.xlsx'
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = 'akbank'
        sheet['A1'] = '31.03.2026 - Önceki dönem'
        sheet['A2'] = 'Kalem'
        sheet['B2'] = 'Toplam'
        sheet['A3'] = 'Satır'
        sheet['B3'] = 1
        workbook.save(self.book)
        workbook.close()

    def test_only_newer_period_can_be_appended(self):
        payloads = [
            {'period': '31.03.2026'},
            {'period': '30.06.2026'},
        ]
        pending = payloads_for_new_periods(payloads, ['31.03.2026'], 'akbank')
        self.assertEqual(['30.06.2026'], [item['period'] for item in pending])
        with self.assertRaisesRegex(RuntimeError, 'tarihsel bosluk'):
            payloads_for_new_periods([{'period': '31.12.2025'}], ['31.03.2026'], 'akbank')

    def test_all_supported_bank_adapters_load(self):
        for bank in side.SIDE_MODULES:
            with self.subTest(bank=bank):
                engine, sheet = side.source_engine(bank)
                self.assertTrue(sheet)
                self.assertTrue(callable(engine.write_stage_workbook))
                self.assertTrue(callable(engine.write_deposit_workbook))

    def test_runner_is_independent_from_main_workbook_and_preserves_old_cells(self):
        pdf = self.root / '30.06.2026.pdf'
        pdf.write_bytes(b'%PDF-1.7')

        def writer(payloads, path):
            workbook = load_workbook(path)
            sheet = workbook['akbank']
            row = sheet.max_row + 3
            for payload in payloads:
                sheet.cell(row, 1, f"{payload['period']} - Yeni dönem")
                sheet.cell(row + 1, 1, 'Kalem')
                sheet.cell(row + 1, 2, 'Toplam')
                sheet.cell(row + 2, 1, 'Yeni satır')
                sheet.cell(row + 2, 2, 2)
            workbook.save(path)
            workbook.close()
            return len(payloads)

        engine = SimpleNamespace(
            list_pdf_files=lambda _folder: [pdf],
            extract_report_date=lambda _path: '30.06.2026',
            extract_history_payloads=lambda files, _cache, _kind, _workers: [{'period': '30.06.2026'} for _ in files],
            write_stage_workbook=writer,
            write_deposit_workbook=writer,
            DEFAULT_WORKERS=1,
        )
        report = self.root / 'audit.json'
        output = self.root / 'dry-run.xlsx'
        args = side.build_parser().parse_args([
            '--bank', 'akbank', '--product', 'stage', '--workbook', str(self.book),
            '--history-dir', str(self.root / 'history'), '--cache-dir', str(self.root / 'cache'),
            '--skip-download', '--dry-run', '--dry-run-output', str(output), '--audit-report', str(report),
        ])
        with patch.object(side, 'source_engine', return_value=(engine, 'akbank')), contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(0, side.run(args))

        self.assertEqual(['31.03.2026'], workbook_periods(self.book, 'akbank'))
        self.assertEqual(['31.03.2026', '30.06.2026'], workbook_periods(output, 'akbank'))
        original = load_workbook(self.book)
        self.assertEqual(1, original['akbank']['B3'].value)
        original.close()
