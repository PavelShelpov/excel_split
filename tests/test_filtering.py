import unittest
import os
import tempfile
from excel_utils.common import validate_row

class TestFiltering(unittest.TestCase):
    def test_validate_row_empty_filters(self):
        """Проверяет, что строка проходит валидацию без фильтров"""
        headers = ["Name", "Value", "Category"]
        row = ["Item1", "100", "A"]
        self.assertTrue(validate_row(row, headers, 1, {}))
    
    def test_validate_row_single_filter_match(self):
        """Проверяет, что строка соответствует одиночному фильтру"""
        headers = ["Name", "Value", "Category"]
        row = ["Item1", "100", "A"]
        self.assertTrue(validate_row(row, headers, 1, {"Category": "A"}))
    
    def test_validate_row_single_filter_no_match(self):
        """Проверяет, что строка не соответствует одиночному фильтру"""
        headers = ["Name", "Value", "Category"]
        row = ["Item1", "100", "A"]
        self.assertFalse(validate_row(row, headers, 1, {"Category": "B"}))
    
    def test_validate_row_multiple_filters_match(self):
        """Проверяет, что строка соответствует нескольким фильтрам"""
        headers = ["Name", "Value", "Category"]
        row = ["Item1", "100", "A"]
        self.assertTrue(validate_row(row, headers, 1, {"Name": "Item1", "Category": "A"}))
    
    def test_validate_row_multiple_filters_no_match(self):
        """Проверяет, что строка не соответствует нескольким фильтрам"""
        headers = ["Name", "Value", "Category"]
        row = ["Item1", "100", "A"]
        self.assertFalse(validate_row(row, headers, 1, {"Name": "Item2", "Category": "A"}))
    
    def test_validate_row_missing_column(self):
        """Проверяет обработку отсутствующего столбца в фильтре"""
        headers = ["Name", "Value", "Category"]
        row = ["Item1", "100", "A"]
        self.assertFalse(validate_row(row, headers, 1, {"NonExistent": "Value"}))
    
    def test_validate_row_with_none_values(self):
        """Проверяет обработку None значений"""
        headers = ["Name", "Value", "Category"]
        row = ["Item1", None, "A"]
        self.assertTrue(validate_row(row, headers, 1, {"Value": ""}))
    
    def test_validate_row_case_sensitivity(self):
        """Проверяет обработку регистра"""
        headers = ["Name", "Value", "Category"]
        row = ["Item1", "100", "A"]
        self.assertTrue(validate_row(row, headers, 1, {"category": "A"}))  # Строчные буквы в ключе

if __name__ == '__main__':
    unittest.main()