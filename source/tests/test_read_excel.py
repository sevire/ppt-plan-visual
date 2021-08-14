from unittest import TestCase, skip

from source.visualiser.read_excel import read_excel


class TestReadExcel(TestCase):
    @skip
    def test_read_excel(self):
        excel_path = 'test_resources/input_files/config_files/KBT-VisualConfig.xlsx'
        sheet_name = 'FormatConfig'
        table = read_excel(excel_path, sheet_name)

        self.assertEqual('Default', table[0]['Format Name'])

