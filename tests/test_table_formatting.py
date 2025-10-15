import unittest
import os
import tempfile
import openpyxl
from excel_utils.workbook import create_filtered_file
from excel_utils.analysis import get_all_sheets_headers

class TestTableFormatting(unittest.TestCase):
    def setUp(self):
        # Создаем тестовый Excel-файл
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_table.xlsx")
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "TestSheet"
        
        # Добавляем заголовки
        ws.append(["ID", "Value", "Category"])
        
        # Добавляем данные
        for i in range(1, 11):
            ws.append([i, i*10, "A" if i % 2 == 0 else "B"])
        
        wb.save(self.test_file)
    
    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir)
    
    def test_table_creation(self):
        """Проверяет создание таблиц в результирующем файле"""
        sheet_headers = get_all_sheets_headers(self.test_file)
        valid_sheets = {k: v for k, v in sheet_headers.items() if v[0] is not None}
        
        # Создаем простой фильтр
        filters = {}
        
        output_file = os.path.join(self.temp_dir, "output.xlsx")
        result = create_filtered_file(self.test_file, output_file, valid_sheets, filters)
        
        self.assertIsNotNone(result)
        self.assertTrue(os.path.exists(result))
        
        # Проверяем, что таблица создана
        wb = openpyxl.load_workbook(result)
        ws = wb["TestSheet"]
        
        # Проверяем, что есть таблица
        self.assertTrue(len(ws.tables) > 0)
        
        # Проверяем, что таблица имеет правильный диапазон
        table = list(ws.tables.values())[0]
        self.assertEqual(table.ref, "A1:C11")
    
    def test_table_with_filtering(self):
        """Проверяет создание таблицы при фильтрации данных"""
        sheet_headers = get_all_sheets_headers(self.test_file)
        valid_sheets = {k: v for k, v in sheet_headers.items() if v[0] is not None}
        
        # Создаем фильтр для категории A
        filters = {"Category": "A"}
        
        output_file = os.path.join(self.temp_dir, "output_filtered.xlsx")
        result = create_filtered_file(self.test_file, output_file, valid_sheets, filters)
        
        self.assertIsNotNone(result)
        self.assertTrue(os.path.exists(result))
        
        # Проверяем, что таблица создана
        wb = openpyxl.load_workbook(result)
        ws = wb["TestSheet"]
        
        # Проверяем, что есть таблица
        self.assertTrue(len(ws.tables) > 0)
        
        # Проверяем, что таблица имеет правильный диапазон
        table = list(ws.tables.values())[0]
        self.assertEqual(table.ref, "A1:C6")  # 1 заголовок + 5 строк данных

if __name__ == '__main__':
    unittest.main()