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
        """Проверяет сокращение названий категорий по новым правилам"""
        # Проверка с предлогами
        self.assertEqual(shorten_category_name("Отдел по учету и налогам"), "ОУиН")
        self.assertEqual(shorten_category_name("Департамент по развитию и инвестициям"), "ДРиИ")
        self.assertEqual(shorten_category_name("Управление для анализа и контроля"), "УАиК")
        
        # Проверка союзов
        self.assertEqual(shorten_category_name("Технический директорат и управление"), "ТДиУ")
        self.assertEqual(shorten_category_name("Департамент развития или анализа"), "ДРиА")
        self.assertEqual(shorten_category_name("Отдел учета а также контроля"), "ОУаК")
        
        # Проверка коротких результатов
        self.assertEqual(shorten_category_name("Кадровый директорат"), "КаД")
        self.assertEqual(shorten_category_name("Департамент анализа"), "ДаА")
        self.assertEqual(shorten_category_name("Управление"), "Уп")
        
        # Проверка дефисов
        self.assertEqual(shorten_category_name("Финансово-инвестиционный департамент"), "ФИД")
        self.assertEqual(shorten_category_name("Техническо-инженерный отдел"), "ТИО")
        self.assertEqual(shorten_category_name("Административно-хозяйственный отдел"), "АХО")
        
        # Проверка последней категории (не сокращается)
        self.assertEqual(shorten_category_name("Кадровый директорат", is_last=True), "Кадровый директорат")
        self.assertEqual(shorten_category_name("Департамент развития", is_last=True), "Департамент развития")
        self.assertEqual(shorten_category_name("Отдел учета", is_last=True), "Отдел учета")
        
        # Проверка обработки регистра
        self.assertEqual(shorten_category_name("ТЕХНИЧЕСКИЙ ДИРЕКТОРАТ"), "ТеД")
        self.assertEqual(shorten_category_name("департамент развития"), "ДеР")
        self.assertEqual(shorten_category_name("отдел учета"), "ОУ")
    
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
            "Team": "Департамент по развитию",
            "Project": "Отдел по ремонту"
        }
        filename = generate_short_filename(base_path, filters)
        self.assertEqual(filename, "base_file_ТеД_ДР_Отдел по ремонту.xlsx")
        
        # Проверка обработки регистра
        filters = {
            "Department": "ТЕХНИЧЕСКИЙ ДИРЕКТОРАТ",
            "Team": "ДЕПАРТАМЕНТ ПО РАЗВИТИЮ",
            "Project": "ОТДЕЛ ПО РЕМОНТУ"
        }
        filename = generate_short_filename(base_path, filters)
        self.assertEqual(filename, "base_file_ТеД_ДР_ОТДЕЛ ПО РЕМОНТУ.xlsx")
        
        # Проверка длины пути
        base_path = os.path.join(self.temp_dir, "very_long_base_name" * 10)
        filters = {
            "Department": "Технический директорат" * 10,
            "Team": "Департамент по развитию" * 10
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