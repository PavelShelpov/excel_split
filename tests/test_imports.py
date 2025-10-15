import unittest
import sys
import os

# Добавляем корневую директорию проекта в путь для импорта
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

class TestModuleImports(unittest.TestCase):
    """Тест на проверку корректности импортов модулей"""
    
    def test_import_main(self):
        """Проверяет, что main.py импортируется без ошибок"""
        try:
            import main
            self.assertTrue(True)
        except Exception as e:
            self.fail(f"Failed to import main: {str(e)}")
    
    def test_import_cli(self):
        """Проверяет, что cli модуль импортируется без ошибок"""
        try:
            from cli import interface
            self.assertTrue(True)
        except Exception as e:
            self.fail(f"Failed to import cli: {str(e)}")
    
    def test_import_core(self):
        """Проверяет, что core модуль импортируется без ошибок"""
        try:
            from core import processing
            self.assertTrue(True)
        except Exception as e:
            self.fail(f"Failed to import core: {str(e)}")
    
    def test_import_excel_utils(self):
        """Проверяет, что excel_utils модуль импортируется без ошибок"""
        try:
            from excel_utils import analysis, filtering, formatting, workbook
            self.assertTrue(True)
        except Exception as e:
            self.fail(f"Failed to import excel_utils: {str(e)}")

if __name__ == '__main__':
    unittest.main()