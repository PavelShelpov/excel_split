import unittest
import os
import tempfile
import openpyxl
from excel_utils.workbook import create_filtered_file
from excel_utils.analysis import get_all_sheets_headers
from excel_utils.common import validate_row

class TestSpecialFormats(unittest.TestCase):
    def setUp(self):
        # Создаем тестовый Excel-файл с особыми стилями и условным форматированием
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_special.xlsx")
        self.output_file = os.path.join(self.temp_dir, "output.xlsx")
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "TestSheet"
        
        # Добавляем заголовки
        ws.append(["ID", "Value", "Category"])
        
        # Добавляем данные
        for i in range(1, 11):
            ws.append([i, i*10, "A" if i % 2 == 0 else "B"])
        
        # Добавляем условное форматирование
        from openpyxl.formatting import Rule
        from openpyxl.formatting.rule import ColorScaleRule
        
        # Создаем правило условного форматирования
        rule = ColorScaleRule(
            start_type='min', start_color='FF0000',
            mid_type='percentile', mid_color='FFFF00',
            end_type='max', end_color='00FF00'
        )
        
        # Добавляем правило к диапазону
        ws.conditional_formatting.add('B2:B11', rule)
        
        # Добавляем специфический стиль (имитируем SAPBEXstdItem)
        from openpyxl.styles import NamedStyle
        sap_style = NamedStyle(name="SAPBEXstdItem")
        sap_style.font = openpyxl.styles.Font(bold=True)
        sap_style.fill = openpyxl.styles.PatternFill(start_color="00FF0000", end_color="00FF0000", fill_type="solid")
        
        # Применяем стиль к ячейке
        ws['A2'].style = sap_style
        
        wb.save(self.test_file)
    
    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir)
    
    def test_special_style_handling(self):
        """Проверяет обработку специфических стилей"""
        sheet_headers = get_all_sheets_headers(self.test_file)
        valid_sheets = {k: v for k, v in sheet_headers.items() if v[0] is not None}
        
        # Создаем простой фильтр
        filters = {"Category": "A"}
        
        # Пытаемся создать фильтрованный файл
        result = create_filtered_file(self.test_file, self.output_file, valid_sheets, filters)
        
        self.assertIsNotNone(result)
        self.assertTrue(os.path.exists(result))
    
    def test_conditional_formatting(self):
        """Проверяет копирование условного форматирования"""
        sheet_headers = get_all_sheets_headers(self.test_file)
        valid_sheets = {k: v for k, v in sheet_headers.items() if v[0] is not None}
        
        # Создаем простой фильтр
        filters = {"Category": "A"}
        
        # Пытаемся создать фильтрованный файл
        result = create_filtered_file(self.test_file, self.output_file, valid_sheets, filters)
        
        self.assertIsNotNone(result)
        
        # Проверяем, что условное форматирование присутствует в выходном файле
        wb = openpyxl.load_workbook(result)
        ws = wb["TestSheet"]
        
        # Проверяем, что условное форматирование существует
        self.assertTrue(len(ws.conditional_formatting) > 0)
    
    def test_empty_filter(self):
        """Проверяет работу с пустым фильтром"""
        sheet_headers = get_all_sheets_headers(self.test_file)
        valid_sheets = {k: v for k, v in sheet_headers.items() if v[0] is not None}
        
        # Пустой фильтр
        filters = {}
        
        # Пытаемся создать фильтрованный файл
        result = create_filtered_file(self.test_file, self.output_file, valid_sheets, filters)
        
        self.assertIsNotNone(result)
        
        # Проверяем, что все данные сохранены
        wb = openpyxl.load_workbook(result)
        ws = wb["TestSheet"]
        
        # Должно быть 11 строк (заголовок + 10 данных)
        self.assertEqual(ws.max_row, 11)
    
    def test_no_matching_data(self):
        """Проверяет работу с фильтром, не находящим совпадений"""
        sheet_headers = get_all_sheets_headers(self.test_file)
        valid_sheets = {k: v for k, v in sheet_headers.items() if v[0] is not None}
        
        # Фильтр, не должен находить совпадений
        filters = {"Category": "C"}
        
        # Пытаемся создать фильтрованный файл
        result = create_filtered_file(self.test_file, self.output_file, valid_sheets, filters)
        
        self.assertIsNone(result)
        self.assertFalse(os.path.exists(self.output_file))

if __name__ == '__main__':
    unittest.main()