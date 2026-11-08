"""作者：Wangyuan；日期：2026-10-29。职责：外部合同标签不得注入Excel公式。"""
import tempfile
import unittest
from pathlib import Path
from openpyxl import load_workbook
from collection import package, write_template


class ReviewTests(unittest.TestCase):
    def test_labels_are_literal_strings(self):
        contract = package('=1+1', [{'id': '=3+3', 'origin': 'sample', 'destination': 'other', 'items': []}], ['=2+2'], [{'name': '=4+4', 'type': 'text', 'required': False}])
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'demo.xlsx'; write_template(contract, path)
            workbook = load_workbook(path)
            try:
                for address in ['A2', 'B2', 'C2', 'E1']:
                    self.assertEqual(workbook['收集表'][address].data_type, 's')
                    self.assertTrue(workbook['收集表'][address].value.startswith('='))
            finally: workbook.close()


if __name__ == '__main__': unittest.main()
