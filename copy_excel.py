# copy_excel.py
import os
import shutil
import openpyxl

def get_all_sheets_headers(file_path, max_scan_rows=10):
    """Анализирует все листы в Excel-файле, возвращает заголовки для каждого."""
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True)
        sheet_results = {}
        
        for ws in wb.worksheets:
            max_non_empty = 0
            header_row = None
            header_row_idx = 0

            # Поиск строки с максимальным количеством данных
            for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan_rows), start=1):
                non_empty_count = sum(1 for cell in row if cell.value is not None)
                if non_empty_count > max_non_empty:
                    max_non_empty = non_empty_count
                    header_row = row
                    header_row_idx = row_idx

            # Сохранение результатов
            if max_non_empty > 0:
                headers = [cell.value for cell in header_row if cell.value is not None]
                sheet_results[ws.title] = (headers, header_row_idx)
            else:
                sheet_results[ws.title] = (None, None)
                
        return sheet_results
    except Exception as e:
        raise ValueError(f"Ошибка анализа Excel: {str(e)}")

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
    
    # Анализ Excel: заголовки во всех листах
    try:
        sheet_headers = get_all_sheets_headers(source)
        valid_sheets = {sheet: data for sheet, data in sheet_headers.items() if data[0] is not None}
        
        # Проверка наличия данных
        if not valid_sheets:
            print("❌ Ошибка: Ни в одном листе не найдены заголовки")
            return
        
        # Вывод информации по листам
        print("\n🔍 Анализ листов:")
        for sheet, (headers, row_idx) in sheet_headers.items():
            if headers is None:
                print(f"  - {sheet}: заголовки не найдены")
            else:
                print(f"  - {sheet} (строка {row_idx}): {', '.join(str(h) for h in headers)}")
        
        # Поиск пересечения заголовков
        all_headers = [set(headers) for headers, _ in valid_sheets.values()]
        common_headers = set.intersection(*all_headers)
        
        # Вывод результата
        if not common_headers:
            print("\n⚠️ Не найдено общих заголовков между листами")
        else:
            print(f"\n✅ Общие заголовки во всех листах: {', '.join(common_headers)}")
    except Exception as e:
        print(f"❌ {str(e)}")
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