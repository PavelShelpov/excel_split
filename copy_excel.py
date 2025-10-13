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

def find_data_range(ws, header_row_idx, selected_column, category):
    """Определяет границы данных: начало, конец и технические строки."""
    # Начало таблицы (строка заголовков)
    data_start = header_row_idx + 1
    
    # Конец данных: последняя строка с непустыми ячейками в области заголовков
    data_end = ws.max_row
    for row_idx in range(ws.max_row, header_row_idx, -1):
        row_has_data = any(cell.value is not None for cell in ws[row_idx])
        if row_has_data:
            data_end = row_idx
            break
    
    # Технические строки ниже: строки после data_end
    return data_start, data_end

def create_filtered_file(source, target, valid_sheets, selected_column, category):
    """Создаёт новый файл с отфильтрованными данными и полным сохранением форматирования."""
    try:
        wb_source = openpyxl.load_workbook(source, read_only=False)
        wb_new = openpyxl.Workbook()
        wb_new.remove(wb_new.active)
        
        for sheet_name in wb_source.sheetnames:
            ws_source = wb_source[sheet_name]
            ws_new = wb_new.create_sheet(title=sheet_name)
            
            # Копирование ширины столбцов
            for col_letter, dim in ws_source.column_dimensions.items():
                ws_new.column_dimensions[col_letter].width = dim.width
            
            # Копирование высоты строк
            for row_idx, dim in ws_source.row_dimensions.items():
                ws_new.row_dimensions[row_idx].height = dim.height
            
            # Копирование объединенных ячеек
            for merged_cell in ws_source.merged_cells.ranges:
                ws_new.merge_cells(str(merged_cell))
            
            # Обработка листов с заголовками
            if sheet_name in valid_sheets:
                headers, header_row_idx = valid_sheets[sheet_name]
                try:
                    col_index = headers.index(selected_column)
                except ValueError:
                    col_index = None
                
                # Определение границ данных
                data_start, data_end = find_data_range(ws_source, header_row_idx, selected_column, category)
                
                # 1. Копирование технических строк выше таблицы (до заголовков)
                for row_idx in range(1, header_row_idx):
                    for col_idx in range(1, ws_source.max_column + 1):
                        cell = ws_source.cell(row=row_idx, column=col_idx)
                        if cell.value is not None or cell.has_style:
                            new_cell = ws_new.cell(row=row_idx, column=col_idx, value=cell.value)
                            if cell.has_style:
                                new_cell.style = cell.style
                                new_cell.font = cell.font.copy()
                                new_cell.border = cell.border.copy()
                                new_cell.fill = cell.fill.copy()
                                new_cell.alignment = cell.alignment.copy()
                
                # 2. Копирование заголовков
                for col_idx in range(1, ws_source.max_column + 1):
                    cell = ws_source.cell(row=header_row_idx, column=col_idx)
                    if cell.value is not None or cell.has_style:
                        new_cell = ws_new.cell(row=header_row_idx, column=col_idx, value=cell.value)
                        if cell.has_style:
                            new_cell.style = cell.style
                            new_cell.font = cell.font.copy()
                            new_cell.border = cell.border.copy()
                            new_cell.fill = cell.fill.copy()
                            new_cell.alignment = cell.alignment.copy()
                
                # 3. Фильтрация данных
                new_row_idx = header_row_idx + 1
                for row_idx in range(data_start, data_end + 1):
                    cell = ws_source.cell(row=row_idx, column=col_index + 1)
                    cell_value = cell.value if cell.value is not None else ""
                    
                    if str(cell_value).strip() == str(category).strip():
                        for col_idx in range(1, ws_source.max_column + 1):
                            source_cell = ws_source.cell(row=row_idx, column=col_idx)
                            if source_cell.value is not None or source_cell.has_style:
                                new_cell = ws_new.cell(row=new_row_idx, column=col_idx, value=source_cell.value)
                                if source_cell.has_style:
                                    new_cell.style = source_cell.style
                                    new_cell.font = source_cell.font.copy()
                                    new_cell.border = source_cell.border.copy()
                                    new_cell.fill = source_cell.fill.copy()
                                    new_cell.alignment = source_cell.alignment.copy()
                        new_row_idx += 1
                
                # 4. Копирование технических строк ниже таблицы
                for row_idx in range(data_end + 1, ws_source.max_row + 1):
                    for col_idx in range(1, ws_source.max_column + 1):
                        cell = ws_source.cell(row=row_idx, column=col_idx)
                        if cell.value is not None or cell.has_style:
                            new_cell = ws_new.cell(row=row_idx, column=col_idx, value=cell.value)
                            if cell.has_style:
                                new_cell.style = cell.style
                                new_cell.font = cell.font.copy()
                                new_cell.border = cell.border.copy()
                                new_cell.fill = cell.fill.copy()
                                new_cell.alignment = cell.alignment.copy()
            else:
                # Копирование листа без изменений (без заголовков)
                for row_idx in range(1, ws_source.max_row + 1):
                    for col_idx in range(1, ws_source.max_column + 1):
                        cell = ws_source.cell(row=row_idx, column=col_idx)
                        if cell.value is not None or cell.has_style:
                            new_cell = ws_new.cell(row=row_idx, column=col_idx, value=cell.value)
                            if cell.has_style:
                                new_cell.style = cell.style
                                new_cell.font = cell.font.copy()
                                new_cell.border = cell.border.copy()
                                new_cell.fill = cell.fill.copy()
                                new_cell.alignment = cell.alignment.copy()
        
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
    
    wb = openpyxl.Workbook()
    wb.save(target_file)
    print(f"\nEmpty file created at: {target_file}")

if __name__ == "__main__":
    main()