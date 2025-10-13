# copy_excel.py
import os
import shutil
import openpyxl  # Установите: pip install openpyxl

def get_excel_headers(file_path):
    """Читает заголовки из первого листа Excel-файла."""
    try:
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        headers = [cell.value for cell in ws[1] if cell.value is not None]
        return headers
    except Exception as e:
        raise ValueError(f"Не удалось прочитать Excel: {str(e)}")

def main():
    print("=== Копирование Excel-файла ===")
    
    # Запрос путей
    source = input("Введите полный путь к исходному Excel-файлу: ").strip('"')
    destination = input("Введите путь к целевой директории: ").strip('"')
    
    # Нормализация путей
    source = os.path.normpath(source)
    destination = os.path.normpath(destination)
    
    # Проверка существования исходного файла
    if not os.path.exists(source):
        print(f"❌ Ошибка: Исходный файл не найден: {source}")
        return
    if not os.path.isfile(source):
        print(f"❌ Ошибка: Указанный путь не является файлом: {source}")
        return
    
    # Проверка расширения .xlsx
    if not source.lower().endswith('.xlsx'):
        print("❌ Ошибка: Файл должен иметь расширение .xlsx")
        return
    
    # Обработка Excel: поиск заголовков
    try:
        headers = get_excel_headers(source)
        if not headers:
            print("❌ Ошибка: Заголовки не найдены (пустая первая строка)")
            return
        print("\n🔍 Найдены заголовки таблицы:")
        print(", ".join(str(h) for h in headers))
    except Exception as e:
        print(f"❌ Ошибка анализа Excel: {str(e)}")
        return
    
    # Копирование файла
    os.makedirs(destination, exist_ok=True)
    try:
        shutil.copy2(source, destination)
        print(f"\n✅ Успешно скопировано в: {os.path.join(destination, os.path.basename(source))}")
    except Exception as e:
        print(f"❌ Ошибка копирования: {str(e)}")

if __name__ == "__main__":
    main()