# copy_excel.py
import os
import re
import openpyxl
from copy import copy

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

def analyze_column(file_path, valid_sheets, selected_column, filters=None):
    """Собирает уникальные значения из указанной колонки с учетом фильтров."""
    if filters is None:
        filters = {}
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
                if not validate_row(row, headers, row_idx, filters):
                    continue
                
                cell_value = row[col_index] if col_index < len(row) else None
                if cell_value is not None and str(cell_value).strip() != "":
                    categories.add(str(cell_value).strip())
        
        return sorted(categories)
    except Exception as e:
        raise ValueError(f"Error analyzing data: {str(e)}")

def validate_row(row, headers, header_row_idx, filters):
    """Проверяет соответствие строки условиям фильтров."""
    if not filters:
        return True
    
    for col, value in filters.items():
        try:
            col_index = headers.index(col)
            cell_value = row[col_index] if col_index < len(row) else None
            if str(cell_value).strip() != str(value).strip():
                return False
        except ValueError:
            return False
    return True

def sanitize_filename(name):
    """Удаляет недопустимые символы из названия файла."""
    return re.sub(r'[\\/*?:"<>|]', '_', name)

def create_filtered_file_single(source, target, valid_sheets, selected_column, category):
    """Создаёт файл с фильтрацией по одной колонке и категории."""
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
            
            if sheet_name in valid_sheets:
                headers, header_row_idx = valid_sheets[sheet_name]
                try:
                    col_index = headers.index(selected_column)
                except ValueError:
                    # Если колонка не найдена, копируем все без изменений
                    for row in ws_source.iter_rows(values_only=True):
                        ws_new.append(row)
                    continue
                
                # 1. Технические строки выше таблицы
                for row_idx in range(1, header_row_idx):
                    for col_idx in range(1, ws_source.max_column + 1):
                        cell = ws_source.cell(row=row_idx, column=col_idx)
                        if cell.value is not None or cell.has_style:
                            new_cell = ws_new.cell(row=row_idx, column=col_idx, value=cell.value)
                            if cell.has_style:
                                new_cell.style = cell.style
                                new_cell.font = copy(cell.font)
                                new_cell.border = copy(cell.border)
                                new_cell.fill = copy(cell.fill)
                                new_cell.alignment = copy(cell.alignment)
                                new_cell.number_format = cell.number_format
                
                # 2. Заголовки
                for col_idx in range(1, ws_source.max_column + 1):
                    cell = ws_source.cell(row=header_row_idx, column=col_idx)
                    if cell.value is not None or cell.has_style:
                        new_cell = ws_new.cell(row=header_row_idx, column=col_idx, value=cell.value)
                        if cell.has_style:
                            new_cell.style = cell.style
                            new_cell.font = copy(cell.font)
                            new_cell.border = copy(cell.border)
                            new_cell.fill = copy(cell.fill)
                            new_cell.alignment = copy(cell.alignment)
                            new_cell.number_format = cell.number_format
                
                # 3. Фильтрация данных
                new_row_idx = header_row_idx + 1
                for row_idx in range(header_row_idx + 1, ws_source.max_row + 1):
                    cell = ws_source.cell(row=row_idx, column=col_index + 1)
                    cell_value = cell.value if cell.value is not None else ""
                    
                    if str(cell_value).strip() == str(category).strip():
                        for col_idx in range(1, ws_source.max_column + 1):
                            source_cell = ws_source.cell(row=row_idx, column=col_idx)
                            if source_cell.value is not None or source_cell.has_style:
                                new_cell = ws_new.cell(row=new_row_idx, column=col_idx, value=source_cell.value)
                                if source_cell.has_style:
                                    new_cell.style = source_cell.style
                                    new_cell.font = copy(source_cell.font)
                                    new_cell.border = copy(source_cell.border)
                                    new_cell.fill = copy(source_cell.fill)
                                    new_cell.alignment = copy(source_cell.alignment)
                                    new_cell.number_format = source_cell.number_format
                        new_row_idx += 1
            else:
                # Копирование листа без изменений (без заголовков)
                for row_idx in range(1, ws_source.max_row + 1):
                    for col_idx in range(1, ws_source.max_column + 1):
                        cell = ws_source.cell(row=row_idx, column=col_idx)
                        if cell.value is not None or cell.has_style:
                            new_cell = ws_new.cell(row=row_idx, column=col_idx, value=cell.value)
                            if cell.has_style:
                                new_cell.style = cell.style
                                new_cell.font = copy(cell.font)
                                new_cell.border = copy(cell.border)
                                new_cell.fill = copy(cell.fill)
                                new_cell.alignment = copy(cell.alignment)
                                new_cell.number_format = cell.number_format
        
        wb_new.save(target)
        return target
    except Exception as e:
        raise ValueError(f"Error during filtering: {str(e)}")

def create_filtered_file_hierarchy(source, target, valid_sheets, filters):
    """Создаёт файл с фильтрацией по иерархии колонок."""
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
            
            if sheet_name in valid_sheets:
                headers, header_row_idx = valid_sheets[sheet_name]
                
                # 1. Технические строки выше таблицы
                for row_idx in range(1, header_row_idx):
                    for col_idx in range(1, ws_source.max_column + 1):
                        cell = ws_source.cell(row=row_idx, column=col_idx)
                        if cell.value is not None or cell.has_style:
                            new_cell = ws_new.cell(row=row_idx, column=col_idx, value=cell.value)
                            if cell.has_style:
                                new_cell.style = cell.style
                                new_cell.font = copy(cell.font)
                                new_cell.border = copy(cell.border)
                                new_cell.fill = copy(cell.fill)
                                new_cell.alignment = copy(cell.alignment)
                                new_cell.number_format = cell.number_format
                
                # 2. Заголовки
                for col_idx in range(1, ws_source.max_column + 1):
                    cell = ws_source.cell(row=header_row_idx, column=col_idx)
                    if cell.value is not None or cell.has_style:
                        new_cell = ws_new.cell(row=header_row_idx, column=col_idx, value=cell.value)
                        if cell.has_style:
                            new_cell.style = cell.style
                            new_cell.font = copy(cell.font)
                            new_cell.border = copy(cell.border)
                            new_cell.fill = copy(cell.fill)
                            new_cell.alignment = copy(cell.alignment)
                            new_cell.number_format = cell.number_format
                
                # 3. Фильтрация данных
                new_row_idx = header_row_idx + 1
                for row_idx in range(header_row_idx + 1, ws_source.max_row + 1):
                    row = ws_source[row_idx]
                    if not validate_row([cell.value for cell in row], headers, header_row_idx, filters):
                        continue
                    
                    for col_idx in range(1, ws_source.max_column + 1):
                        source_cell = ws_source.cell(row=row_idx, column=col_idx)
                        if source_cell.value is not None or source_cell.has_style:
                            new_cell = ws_new.cell(row=new_row_idx, column=col_idx, value=source_cell.value)
                            if source_cell.has_style:
                                new_cell.style = source_cell.style
                                new_cell.font = copy(source_cell.font)
                                new_cell.border = copy(source_cell.border)
                                new_cell.fill = copy(source_cell.fill)
                                new_cell.alignment = copy(source_cell.alignment)
                                new_cell.number_format = source_cell.number_format
                    new_row_idx += 1
            else:
                for row_idx in range(1, ws_source.max_row + 1):
                    for col_idx in range(1, ws_source.max_column + 1):
                        cell = ws_source.cell(row=row_idx, column=col_idx)
                        if cell.value is not None or cell.has_style:
                            new_cell = ws_new.cell(row=row_idx, column=col_idx, value=cell.value)
                            if cell.has_style:
                                new_cell.style = cell.style
                                new_cell.font = copy(cell.font)
                                new_cell.border = copy(cell.border)
                                new_cell.fill = copy(cell.fill)
                                new_cell.alignment = copy(cell.alignment)
                                new_cell.number_format = cell.number_format
        
        wb_new.save(target)
        return target
    except Exception as e:
        raise ValueError(f"Error during filtering: {str(e)}")

def generate_hierarchy_files(source, destination, valid_sheets, hierarchy_columns, base_name):
    """Создает файлы для всех уровней иерархии."""
    created_files = []
    filter_stack = [({})]  # Начинаем с пустого фильтра
    
    for level, column in enumerate(hierarchy_columns):
        new_stack = []
        
        for current_filters in filter_stack:
            # Получаем значения для текущего уровня
            categories = analyze_column(source, valid_sheets, column, current_filters)
            
            for category in categories:
                # Создаем новый фильтр
                new_filters = current_filters.copy()
                new_filters[column] = category
                
                # Формируем имя файла
                safe_parts = [sanitize_filename(v) for v in new_filters.values()]
                suffix = "_".join(safe_parts) if safe_parts else "All"
                target_file = os.path.join(destination, f"{base_name}_{suffix}.xlsx")
                
                # Создаем файл
                create_filtered_file_hierarchy(source, target_file, valid_sheets, new_filters)
                created_files.append(target_file)
                
                # Добавляем в стек для следующего уровня
                new_stack.append(new_filters)
        
        filter_stack = new_stack  # Переходим к следующему уровню
    
    return created_files

def main():
    print("=== Copy Excel File ===")
    print("To cancel the operation, press Ctrl+C at any time")
    
    try:
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
            
            # Выбор колонок для фильтрации
            print("\nAvailable columns for filtering:", ", ".join(common_headers))
            columns_input = input("Enter columns for filtering (comma-separated): ").strip()
            
            if columns_input.lower() == "cancel":
                print("Operation cancelled by user")
                return
                
            hierarchy_columns = [col.strip() for col in columns_input.split(",") if col.strip()]
            
            # Проверка валидности выбранных колонок
            invalid_columns = [col for col in hierarchy_columns if col not in common_headers]
            if invalid_columns:
                print(f"Error: Invalid columns: {', '.join(invalid_columns)}")
                return
            if not hierarchy_columns:
                print("Error: No valid columns selected")
                return
            
            # Выбор категорий
            if len(hierarchy_columns) == 1:
                selected_column = hierarchy_columns[0]
                categories = analyze_column(source, valid_sheets, selected_column)
                
                if not categories:
                    print(f"Warning: No data found in column '{selected_column}'")
                    return
                
                print(f"\nCategories in column '{selected_column}':")
                for i, cat in enumerate(categories, 1):
                    print(f"  {i}. {cat}")
                
                selected_category = input("\nEnter category for filtering (or 'all' for all categories): ").strip()
                if selected_category.lower() == "cancel":
                    print("Operation cancelled by user")
                    return
                    
                if selected_category.lower() == "all":
                    os.makedirs(destination, exist_ok=True)
                    base_name = os.path.splitext(os.path.basename(source))[0]
                    for category in categories:
                        target_file = os.path.join(destination, f"{base_name}_{sanitize_filename(category)}.xlsx")
                        create_filtered_file_single(source, target_file, valid_sheets, selected_column, category)
                    print(f"Created {len(categories)} filtered files")
                else:
                    if selected_category not in categories:
                        print(f"Error: Category '{selected_category}' not found in the list")
                        return
                    os.makedirs(destination, exist_ok=True)
                    base_name = os.path.splitext(os.path.basename(source))[0]
                    target_file = os.path.join(destination, f"{base_name}_{sanitize_filename(selected_category)}.xlsx")
                    create_filtered_file_single(source, target_file, valid_sheets, selected_column, selected_category)
                    print(f"\nFiltered file saved at: {target_file}")
            else:
                # Обработка иерархии
                os.makedirs(destination, exist_ok=True)
                base_name = os.path.splitext(os.path.basename(source))[0]
                created_files = generate_hierarchy_files(source, destination, valid_sheets, hierarchy_columns, base_name)
                
                if created_files:
                    print(f"\nCreated {len(created_files)} hierarchical files:")
                    for file in created_files:
                        print(f"  - {file}")
                else:
                    print("Warning: No files created (no data matched the filters)")
                return
    except KeyboardInterrupt:
        print("\nOperation cancelled by user (Ctrl+C)")
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