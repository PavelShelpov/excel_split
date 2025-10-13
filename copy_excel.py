# copy_excel.py
import os
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

            for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan_rows), start=1):
                non_empty_count = sum(1 for cell in row if cell.value is not None)
                if non_empty_count > max_non_empty:
                    max_non_empty = non_empty_count
                    header_row = row
                    header_row_idx = row_idx

            if max_non_empty > 0:
                headers = [cell.value for cell in header_row if cell.value is not None]
                sheet_results[ws.title] = (headers, header_row_idx)
            else:
                sheet_results[ws.title] = (None, None)
                
        return sheet_results
    except Exception as e:
        raise ValueError(f"Error analyzing Excel: {str(e)}")

def analyze_column(file_path, valid_sheets, selected_column):
    """Собирает уникальные значения из указанной колонки по всем листам."""
    try:
        wb = openpyxl.load_workbook(file_path)
        categories = set()
        
        for sheet_name, (headers, row_idx) in valid_sheets.items():
            ws = wb[sheet_name]
            try:
                col_index = headers.index(selected_column)
            except ValueError:
                continue
            
            for row in ws.iter_rows(min_row=row_idx + 1, values_only=True):
                cell_value = row[col_index] if col_index < len(row) else None
                if cell_value is not None and str(cell_value).strip() != "":
                    categories.add(str(cell_value).strip())
        
        return sorted(categories)
    except Exception as e:
        raise ValueError(f"Error analyzing data: {str(e)}")

def create_filtered_file(source, target, valid_sheets, selected_column, category):
    """Создаёт новый файл с отфильтрованными данными без внешних зависимостей."""
    try:
        wb_source = openpyxl.load_workbook(source, read_only=True)
        wb_new = openpyxl.Workbook()
        wb_new.remove(wb_new.active)  # Удаляем дефолтный лист
        
        for sheet_name in wb_source.sheetnames:
            ws_source = wb_source[sheet_name]
            ws_new = wb_new.create_sheet(title=sheet_name)
            
            # Обработка листов с заголовками
            if sheet_name in valid_sheets:
                headers, row_idx = valid_sheets[sheet_name]
                try:
                    col_index = headers.index(selected_column)
                except ValueError:
                    # Копируем все строки без фильтрации
                    for row in ws_source.iter_rows(values_only=True):
                        ws_new.append(row)
                    continue
                
                # Копируем заголовки
                header_row = list(ws_source.iter_rows(min_row=row_idx, max_row=row_idx, values_only=True))[0]
                ws_new.append(header_row)
                
                # Фильтрация данных
                for row in ws_source.iter_rows(min_row=row_idx + 1, values_only=True):
                    cell_value = row[col_index] if col_index < len(row) else None
                    if str(cell_value).strip() == str(category).strip():
                        ws_new.append(row)
            else:
                # Копируем лист без изменений
                for row in ws_source.iter_rows(values_only=True):
                    ws_new.append(row)
        
        wb_new.save(target)
        return target
    except Exception as e:
        raise ValueError(f"Error during filtering: {str(e)}")

def main():
    print("=== Copy Excel File ===")
    
    # Запрос путей
    source = input("Enter full path to source Excel file: ").strip('"')
    destination = input("Enter target directory path: ").strip('"')
    
    # Нормализация путей
    source = os.path.normpath(source)
    destination = os.path.normpath(destination)
    
    # Проверка существования исходного файла
    if not os.path.exists(source):
        print(f"Error: Source file not found: {source}")
        return
    if not os.path.isfile(source):
        print(f"Error: Path is not a file: {source}")
        return
    
    # Анализ Excel: заголовки во всех листах
    try:
        sheet_headers = get_all_sheets_headers(source)
        valid_sheets = {sheet: data for sheet, data in sheet_headers.items() if data[0] is not None}
        
        if not valid_sheets:
            print("Error: No headers found in any sheet")
            return
        
        # Вывод информации по листам
        print("\nSheet analysis:")
        for sheet, (headers, row_idx) in sheet_headers.items():
            if headers is None:
                print(f"  - {sheet}: headers not found")
            else:
                print(f"  - {sheet} (row {row_idx}): {', '.join(str(h) for h in headers)}")
        
        # Поиск пересечения заголовков
        all_headers = [set(headers) for headers, _ in valid_sheets.values()]
        common_headers = set.intersection(*all_headers) if all_headers else set()
        
        if not common_headers:
            print("\nWarning: No common headers found between sheets")
        else:
            print(f"\nCommon headers in all sheets: {', '.join(common_headers)}")
            
            # Выбор колонки для анализа
            print("\nAvailable columns for analysis:", ", ".join(common_headers))
            selected_column = input("Enter column name for analysis: ").strip()
            
            if selected_column not in common_headers:
                print(f"Error: Column '{selected_column}' not found in common headers")
            else:
                categories = analyze_column(source, valid_sheets, selected_column)
                if not categories:
                    print(f"Warning: No data found in column '{selected_column}'")
                else:
                    print(f"\nCategories in column '{selected_column}':")
                    for i, cat in enumerate(categories, 1):
                        print(f"  {i}. {cat}")
                    
                    # Выбор категории для фильтрации
                    selected_category = input("\nEnter category for filtering: ").strip()
                    
                    if selected_category not in categories:
                        print(f"Error: Category '{selected_category}' not found in the list")
                    else:
                        # Создание отфильтрованного файла
                        os.makedirs(destination, exist_ok=True)
                        target_file = os.path.join(destination, os.path.basename(source))
                        
                        # Удаляем расширение и добавляем .xlsx (гарантированно чистый файл)
                        if target_file.lower().endswith(('.xlsx', '.xlsm', '.xlsb')):
                            target_file = os.path.splitext(target_file)[0] + ".xlsx"
                        else:
                            target_file += ".xlsx"
                        
                        create_filtered_file(source, target_file, valid_sheets, selected_column, selected_category)
                        print(f"\nFiltered file saved at: {target_file}")
                        return
    except Exception as e:
        print(f"Error: {str(e)}")
        return
    
    # Стандартное копирование
    os.makedirs(destination, exist_ok=True)
    target_file = os.path.join(destination, os.path.basename(source))
    if target_file.lower().endswith(('.xlsx', '.xlsm', '.xlsb')):
        target_file = os.path.splitext(target_file)[0] + ".xlsx"
    else:
        target_file += ".xlsx"
    
    # Создаём пустой файл (без данных, только структура)
    wb = openpyxl.Workbook()
    wb.save(target_file)
    print(f"\nEmpty file created at: {target_file}")

if __name__ == "__main__":
    main()