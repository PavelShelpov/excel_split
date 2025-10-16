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
    
    def test_single_category(self):
        """Проверяет, что при одной категории используется полное название"""
        self.assertEqual(shorten_category_name("Технический директорат", is_last=True), "Технический директорат")
        self.assertEqual(shorten_category_name("Департамент", is_last=True), "Департамент")
    
    def test_multiple_categories(self):
        """Проверяет сокращение нескольких категорий по новым правилам"""
        # Проверка основного правила
        self.assertEqual(shorten_category_name("Технический директорат", is_last=False), "ТеД")
        self.assertEqual(shorten_category_name("Департамент инженерно-технического обеспечения", is_last=False), "ДеИ")
        
        # Проверка сокращения с несколькими словами
        self.assertEqual(shorten_category_name("Технический директорат_Департамент", is_last=False), "ТеД")
        self.assertEqual(shorten_category_name("Департамент инженерно-технического обеспечения_Отдел", is_last=False), "ДеИ")
        
        # Проверка обработки регистра
        self.assertEqual(shorten_category_name("ТЕХНИЧЕСКИЙ ДИРЕКТОРАТ", is_last=False), "ТеД")
        self.assertEqual(shorten_category_name("департамент инженерно-технического обеспечения", is_last=False), "ДеИ")
        
        # Проверка специфических случаев
        self.assertEqual(shorten_category_name("a", is_last=False), "A")
        self.assertEqual(shorten_category_name("ab", is_last=False), "Ab")
        self.assertEqual(shorten_category_name("abc", is_last=False), "Ab")
        self.assertEqual(shorten_category_name("abc def", is_last=False), "AbD")
    
    def test_generate_short_filename(self):
        """Проверяет генерацию коротких имен файлов с новыми правилами"""
        # Базовый случай с одной категорией
        base_path = os.path.join(self.temp_dir, "base_file")
        filters = {"Department": "Human Resources"}
        filename = generate_short_filename(base_path, filters)
        self.assertEqual(filename, "base_file_Human Resources.xlsx")
        
        # Сценарий с несколькими категориями
        filters = {
            "Department": "Технический директорат",
            "Team": "Департамент инженерно-технического обеспечения",
            "Project": "Отдел по ремонту"
        }
        filename = generate_short_filename(base_path, filters)
        self.assertEqual(filename, "base_file_ТеД_ДеИ_Отдел по ремонту.xlsx")
        
        # Проверка обработки регистра
        filters = {
            "Department": "ТЕХНИЧЕСКИЙ ДИРЕКТОРАТ",
            "Team": "ДЕПАРТАМЕНТ ИНЖЕНЕРНО-ТЕХНИЧЕСКОГО ОБЕСПЕЧЕНИЯ",
            "Project": "ОТДЕЛ ПО РЕМОНТУ"
        }
        filename = generate_short_filename(base_path, filters)
        self.assertEqual(filename, "base_file_ТеД_ДеИ_ОТДЕЛ ПО РЕМОНТУ.xlsx")
        
        # Проверка длины пути
        base_path = os.path.join(self.temp_dir, "very_long_base_name" * 10)
        filters = {
            "Department": "Технический директорат" * 10,
            "Team": "Департамент инженерно-технического обеспечения" * 10
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