import unittest
import os
import tempfile
from excel_utils.formatting import sanitize_filename, shorten_category_name, generate_short_filename

class TestFilenameShortening(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
    
    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir)
    
    def test_sanitize_filename(self):
        """Проверяет удаление недопустимых символов из названия файла"""
        self.assertEqual(sanitize_filename("Test:File?Name*"), "Test_File_Name_")
        self.assertEqual(sanitize_filename("File with spaces"), "File with spaces")
        self.assertEqual(sanitize_filename("File<>|"), "File____")
    
    def test_shorten_category_name(self):
        """Проверяет сокращение длинных названий категорий"""
        # Тест с коротким именем
        self.assertEqual(shorten_category_name("IT"), "IT")
        
        # Тест с одним длинным словом
        self.assertEqual(shorten_category_name("Department"), "Dep")
        
        # Тест с несколькими словами
        self.assertEqual(shorten_category_name("Human Resources"), "HR")
        self.assertEqual(shorten_category_name("Apparatus Management"), "AM")
        self.assertEqual(shorten_category_name("Apparatus Management Development"), "AMD")
    
    def test_generate_short_filename(self):
        """Проверяет генерацию коротких имен файлов"""
        # Базовый случай
        base_path = os.path.join(self.temp_dir, "base_file")
        filters = {
            "Department": "Human Resources",
            "Team": "Development",
            "Project": "New Project"
        }
        filename = generate_short_filename(base_path, filters)
        self.assertEqual(filename, "base_file_HR_Dev_New Project.xlsx")
        
        # Сокращение для длинных имен
        base_path = os.path.join(self.temp_dir, "base_file")
        filters = {
            "Department": "Apparatus Management Development Department",
            "Team": "Development Team for New Projects"
        }
        filename = generate_short_filename(base_path, filters)
        self.assertEqual(filename, "base_file_AMDD_DTNPNP.xlsx")
        
        # Проверка ограничения длины
        base_path = os.path.join(self.temp_dir, "very_long_base_name" * 10)
        filters = {
            "Department": "Very Long Department Name " * 10,
            "Team": "Very Long Team Name " * 10
        }
        filename = generate_short_filename(base_path, filters)
        self.assertTrue(len(filename) <= 200)
        
        # Проверка хэширования при очень длинных именах
        base_path = os.path.join(self.temp_dir, "a" * 200)
        filters = {
            "Department": "b" * 200,
            "Team": "c" * 200
        }
        filename = generate_short_filename(base_path, filters)
        self.assertIn("short_", filename)
        self.assertEqual(len(filename), 200)

if __name__ == '__main__':
    unittest.main()