import unittest
from excel_utils.analysis import get_all_sheets_headers, analyze_column
from excel_utils.common import validate_row
import os
import tempfile

class TestExcelAnalysis(unittest.TestCase):
    def setUp(self):
        # Создаем тестовый Excel-файл
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test.xlsx")
        
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.append(["Header1", "Header2", "Header3"])
        ws.append(["Data1", "ValueA", "100"])
        ws.append(["Data2", "ValueB", "200"])
        ws.append(["Data3", "ValueA", "300"])
        wb.save(self.test_file)
    
    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir)
    
    def test_get_all_sheets_headers(self):
        """Проверяет корректность определения заголовков в Excel-файле"""
        headers = get_all_sheets_headers(self.test_file)
        self.assertIn("Sheet1", headers)
        self.assertEqual(headers["Sheet1"][0], ["Header1", "Header2", "Header3"])
    
    def test_analyze_column(self):
        """Проверяет сбор уникальных значений из колонки"""
        headers = get_all_sheets_headers(self.test_file)
        valid_sheets = {k: v for k, v in headers.items() if v[0] is not None}
        
        categories = analyze_column(self.test_file, valid_sheets, "Header2")
        self.assertEqual(categories, ["ValueA", "ValueB"])
    
    def test_validate_row(self):
        """Проверяет фильтрацию строк по условиям"""
        headers = ["Header1", "Header2", "Header3"]
        row = ["Data1", "ValueA", "100"]
        
        # Проверяем фильтр по одной колонке
        self.assertTrue(validate_row(row, headers, 1, {"Header2": "ValueA"}))
        self.assertFalse(validate_row(row, headers, 1, {"Header2": "ValueB"}))
        
        # Проверяем фильтр по нескольким колонкам
        self.assertTrue(validate_row(row, headers, 1, {"Header1": "Data1", "Header2": "ValueA"}))
        self.assertFalse(validate_row(row, headers, 1, {"Header1": "Data1", "Header2": "ValueB"}))

if __name__ == '__main__':
    unittest.main()