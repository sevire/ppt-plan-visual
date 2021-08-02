from unittest import TestCase

from source.visualiser.read_excel import read_excel


class TestReadExcel(TestCase):
    def test_read_excel(self):
        excel_path = '/Users/Development/PycharmProjects/ppt-plan-visual/source/tests/test_resources/input_files/config_files/KBT-VisualConfig.xlsx'
        sheet_name = 'FormatConfig'
        table = read_excel(excel_path, sheet_name)

        self.assertEqual('Default', table[0]['Format Name'])

