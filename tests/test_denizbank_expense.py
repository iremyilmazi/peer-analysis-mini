import copy
import importlib
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from teb_bankasi import teb_dipnot_guncelle as writer
from teb_bankasi.peerpdf_scripts import teb_dipnot as shared


class DenizbankExpenseTests(unittest.TestCase):
    def setUp(self):
        # Bank adapters patch shared globals at import; restore them after each test.
        before = dict(vars(shared))
        self.addCleanup(lambda: vars(shared).update(before))
        self.parser = importlib.import_module('denizbank.peerpdf_scripts.denizbank_dipnot')
        self.title = shared.OTHER_OPERATING_EXPENSE_TITLE

    def normalize(self, amount, label='Ornek gider -- aciklamasi'):
        return self.parser._normalize_expense_tables(
            [(self.title, ['Kalem', '', '', 'Toplam'], [[label, '', '', amount]])]
        )[0][2][0]

    def test_single_double_and_spaced_dash_glyphs_become_numeric_zero(self):
        for value in ('-', '--', '---', '\u2013', '\u2014', '\u2013\u2014', ' - - ', '\u00a0--\u00a0'):
            with self.subTest(value=value):
                row = self.normalize(value)
                self.assertEqual(0, row[3])
                self.assertIs(type(row[3]), int)
                self.assertEqual(['Ornek gider -- aciklamasi', '', ''], row[:3])

    def test_actual_blank_fields_are_not_invented_as_zero(self):
        for value in ('', '  ', None):
            with self.subTest(value=value):
                self.assertEqual(value, self.normalize(value)[3])

    def test_existing_amounts_and_negative_parentheses_are_preserved(self):
        for value in ('1.234', '12.345', '(123)', '123,45', 0, 42, 1.5):
            with self.subTest(value=value):
                self.assertEqual(value, self.normalize(value)[3])

    def test_unexpected_amount_text_fails_instead_of_silently_zeroing(self):
        for value in ('aciklama', 'N/A', '--123', '1 --', True, {}):
            with self.subTest(value=value), self.assertRaises(RuntimeError):
                self.normalize(value)

    def test_other_notes_and_input_rows_remain_unchanged(self):
        tables = [('Baska dipnot', ['Kalem', 'Toplam'], [['etiket', '--']]),
                  (self.title, ['Kalem', '', '', 'Toplam'], [['gider', '', '', '--']])]
        before = copy.deepcopy(tables)
        result = self.parser._normalize_expense_tables(tables)
        self.assertEqual(before, tables)
        self.assertEqual(before[0], result[0])
        self.assertEqual(0, result[1][2][0][3])

    def test_missing_empty_and_malformed_expense_table_fail_closed(self):
        for tables in ([], [(self.title, [], [])], [(self.title, [], [['gider', '--']])]):
            with self.subTest(tables=tables), self.assertRaises(RuntimeError):
                self.parser._normalize_expense_tables(tables)

    def test_full_parser_entrypoint_applies_local_rule(self):
        with patch.object(shared, 'extract_all_tables', return_value=[
            (self.title, ['Kalem', '', '', 'Toplam'], [['gider', '', '', '--']])
        ]) as extract:
            result = self.parser.extract_all_tables(Path('example.pdf'))
        extract.assert_called_once_with(Path('example.pdf'))
        self.assertIs(type(result[0][2][0][3]), int)

    def test_shared_extractor_is_not_replaced_with_bank_wrapper(self):
        self.assertIsNot(shared.extract_all_tables, self.parser.extract_all_tables)

    def test_zero_remains_numeric_through_shared_excel_writer(self):
        self.assertIs(type(writer._excel_number(self.normalize('--')[3])), int)
        self.assertEqual(0, writer._excel_number(self.normalize('--')[3]))

    def test_version_10_cache_is_rebuilt_and_version_11_reused(self):
        adapter = importlib.import_module('denizbank.denizbank_dipnot_guncelle')
        self.assertEqual(11, adapter.CACHE_VERSION)
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            pdf = root/'report_30062026.pdf'
            pdf.write_bytes(b'test PDF placeholder, extractor mocked')
            cache = root/'cache'
            cache.mkdir()
            with patch.object(writer, 'CACHE_VERSION', adapter.CACHE_VERSION), \
                 patch.object(writer, 'cache_path_for_pdf', return_value=cache/'report.json'):
                path = cache/'report.json'
                stat = pdf.stat()
                payload = dict(cache_version=10, source_size=stat.st_size, source_mtime=stat.st_mtime,
                               tables=[dict(title=self.title, headers=[], rows=[['gider', '', '', '--']])])
                path.write_text(json.dumps(payload), encoding='utf-8')
                fixed = [(self.title, [], [['gider', '', '', 0]])]
                with patch.object(writer, 'safe_extract_peerpdf_tables', return_value=fixed) as extract:
                    self.assertEqual(fixed, writer.load_or_extract_pdf_tables(pdf, cache))
                    self.assertEqual(fixed, writer.load_or_extract_pdf_tables(pdf, cache))
                extract.assert_called_once()
                self.assertEqual(11, json.loads(path.read_text(encoding='utf-8'))['cache_version'])


if __name__ == '__main__':
    unittest.main()
