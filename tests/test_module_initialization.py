import unittest
import sys
import os
import logging

# Добавляем корневую директорию проекта в путь для импорта
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

class TestModuleInitialization(unittest.TestCase):
    """Тест на проверку корректной инициализации всех модулей"""
    
    def setUp(self):
        # Настройка логгера для тестов
        logging.basicConfig(level=logging.DEBUG)
        self.logger = logging.getLogger('test_initialization')
    
    def test_excel_utils_initialization(self):
        """Проверяет, что пакет excel_utils инициализируется без ошибок"""
        try:
            import excel_utils
            self.assertIsNotNone(excel_utils)
            self.logger.debug("excel_utils package imported successfully")
        except Exception as e:
            self.fail(f"Failed to import excel_utils package: {str(e)}")
    
    def test_all_modules_import(self):
        """Проверяет импорт всех модулей проекта"""
        modules = [
            'excel_utils.analysis',
            'excel_utils.filtering',
            'excel_utils.formatting',
            'excel_utils.workbook',
            'excel_utils.common',
            'core.processing',
            'cli.interface'
        ]
        
        for module in modules:
            try:
                __import__(module)
                self.logger.debug(f"Successfully imported {module}")
            except Exception as e:
                self.fail(f"Failed to import {module}: {str(e)}")
    
    def test_function_availability(self):
        """Проверяет доступность критических функций"""
        try:
            from excel_utils.formatting import generate_short_filename
            self.assertTrue(callable(generate_short_filename))
            
            from excel_utils.common import validate_row
            self.assertTrue(callable(validate_row))
            
            from excel_utils.workbook import create_filtered_file
            self.assertTrue(callable(create_filtered_file))
        except Exception as e:
            self.fail(f"Critical functions not available: {str(e)}")

if __name__ == '__main__':
    unittest.main()