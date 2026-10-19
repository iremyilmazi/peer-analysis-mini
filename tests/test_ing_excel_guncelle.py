from __future__ import annotations

import unittest

import pandas as pd

from ing_bank.ing_excel_guncelle import current_cumulative_column, income_rows


class IngIncomeTests(unittest.TestCase):
    def table(self, columns=(3, 4, 5, 6), width=8):
        rows = [[None] * width for _ in range(7)]
        headers = [
            'Cari Dönem 01.01.2026 - 30.06.2026',
            'Önceki Dönem 01.01.2025 - 30.06.2025',
            'Cari Dönem 3 Aylık 01.04.2026 - 30.06.2026',
            'Önceki Dönem 3 Aylık 01.04.2025 - 30.06.2025',
        ]
        for col, header in zip(columns, headers):
            rows[0][col] = header
        rows[1][1] = 'Kar veya Zarar Tablosu'
        for index, label, values in [
            (2, 'FAİZ GELİRLERİ', ['27.891', '18.000', '15.000', '10.000']),
            (3, 'FAİZ GİDERLERİ (-)', ['(19.916)', '(12.000)', '(11.000)', '(7.000)']),
            (4, 'TEMETTÜ GELİRLERİ', ['0', '-', '0', '-']),
            (5, 'Finansal Kiralama Faiz Gelirleri', ['-', '-', '-', '-']),
            (6, 'Eksik veri', [None, None, None, None]),
        ]:
            rows[index][1] = label
            for col, value in zip(columns, values):
                rows[index][col] = value
        return pd.DataFrame(rows)

    def test_compact_columns_use_current_cumulative_not_quarter(self):
        table = self.table()
        self.assertEqual(3, current_cumulative_column(table, '30.06.2026'))
        result = income_rows([table], '30.06.2026')
        self.assertEqual('27.891', result[0]['value'])
        self.assertEqual('(19.916)', result[1]['value'])
        self.assertEqual(4, len(result))

    def test_expanded_columns_preserve_legacy_layout(self):
        table = self.table((7, 8, 9, 10), 12)
        self.assertEqual(7, current_cumulative_column(table, '30.06.2026'))
        self.assertEqual('27.891', income_rows([table], '30.06.2026')[0]['value'])

    def test_reordered_columns_follow_header(self):
        table = self.table((6, 5, 4, 3), 8)
        self.assertEqual(6, current_cumulative_column(table, '30.06.2026'))
        self.assertEqual('27.891', income_rows([table], '30.06.2026')[0]['value'])

    def test_explicit_dash_is_zero_but_missing_remains_missing(self):
        result = income_rows([self.table()], '30.06.2026')
        self.assertEqual('0', result[2]['value'])
        self.assertEqual(0, result[3]['value'])
        self.assertNotIn('Eksik veri', [r['label'] for r in result])

    def test_wrong_period_stops_instead_of_selecting_other_column(self):
        with self.assertRaisesRegex(RuntimeError, 'tekil bulunamadi'):
            income_rows([self.table()], '30.09.2026')

    def test_quarterly_only_header_is_not_accepted_as_cumulative(self):
        table = self.table()
        table.iloc[0, 3] = 'Cari Dönem 3 Aylık 01.04.2026 - 30.06.2026'
        with self.assertRaisesRegex(RuntimeError, 'tekil bulunamadi'):
            income_rows([table], '30.06.2026')

    def test_duplicate_cumulative_header_is_rejected(self):
        table = self.table()
        table.iloc[0, 7] = table.iloc[0, 3]
        with self.assertRaisesRegex(RuntimeError, 'tekil bulunamadi'):
            income_rows([table], '30.06.2026')

    def test_empty_financial_block_is_rejected(self):
        table = self.table()
        table.iloc[2:, 3] = None
        with self.assertRaisesRegex(RuntimeError, 'verisi bulunamadi'):
            income_rows([table], '30.06.2026')


if __name__ == '__main__':
    unittest.main()
