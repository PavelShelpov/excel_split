import os
import logging
from excel_utils import (
    get_all_sheets_headers,
    analyze_column,
    get_all_combinations,
    select_categories_sequentially,
    sanitize_filename,
    create_filtered_file
)

logger = logging.getLogger('excel_splitter')

def process_file():
    """Обрабатывает один файл: выбор файла, директории, колонок, категорий, создание файлов."""
    logger.info("Starting file processing")
    print("\n=== Copy Excel File ===")
    print("To cancel the operation, press Ctrl+C at any time")
    try:
        # Шаг 0: Выбор исходного файла
        while True:
            source = input("\nEnter full path to source Excel file: ").strip('"')
            if source.lower() == "cancel":
                print("Operation cancelled by user")
                return False
            if os.path.exists(source) and os.path.isfile(source):
                # Проверка формата файла
                if not (source.lower().endswith('.xlsx') or source.lower().endswith('.xlsm')):
                    print("Error: File must have .xlsx or .xlsm extension")
                    continue
                break
            print(f"Error: Source file not found or is not a file: {source}")
        
        # Шаг 1: Выбор целевой директории
        while True:
            destination = input("Enter target directory path: ").strip('"')
            if destination.lower() == "cancel":
                print("Operation cancelled by user")
                return False
            if os.path.isdir(destination):
                break
            print(f"Error: Target directory does not exist: {destination}")
        
        # Анализ Excel: заголовки во всех листах
        sheet_headers = get_all_sheets_headers(source)
        valid_sheets = {sheet: data for sheet, data in sheet_headers.items() if data[0] is not None}
        if not valid_sheets:
            logger.error("No headers found in any sheet")
            print("Error: No headers found in any sheet")
            return False
        
        # Поиск пересечения заголовков
        all_headers = [set(headers) for headers, _ in valid_sheets.values()]
        common_headers = set.intersection(*all_headers) if all_headers else set()
        if not common_headers:
            logger.warning("No common headers found between sheets")
            print("\nWarning: No common headers found between sheets")
            return False
        
        # Шаг 2: Выбор колонок для фильтрации
        print("\nAvailable columns for filtering:")
        common_headers_list = list(common_headers)
        for i, col in enumerate(common_headers_list, 1):
            print(f"  {i}. {col}")
        print("  b. Назад")
        print("  c. Отмена")
        while True:
            columns_input = input("Enter columns for filtering (comma-separated numbers or names): ").strip()
            if columns_input.lower() in ["c", "cancel", "отмена"]:
                print("Operation cancelled by user")
                return False
            if columns_input.lower() in ["b", "back", "назад"]:
                return False  # Возврат к началу
            # Обработка номеров колонок
            hierarchy_columns = []
            invalid_inputs = []
            for item in columns_input.split(","):
                item = item.strip()
                if item.isdigit():
                    idx = int(item) - 1
                    if 0 <= idx < len(common_headers_list):
                        hierarchy_columns.append(common_headers_list[idx])
                    else:
                        invalid_inputs.append(item)
                else:
                    hierarchy_columns.append(item)
            # Проверка валидности
            invalid_columns = [col for col in hierarchy_columns if col not in common_headers]
            if invalid_columns or invalid_inputs:
                invalid_list = invalid_columns + invalid_inputs
                print(f"Error: Invalid columns: {', '.join(invalid_list)}")
                continue
            if not hierarchy_columns:
                print("Error: No valid columns selected")
                continue
            break
        
        # Шаг 3: Последовательный выбор категорий
        print("\nStarting sequential category selection...")
        all_combinations = select_categories_sequentially(source, valid_sheets, hierarchy_columns)
        if not all_combinations:
            print("No combinations selected")
            return False
        
        # Создание файлов
        os.makedirs(destination, exist_ok=True)
        base_name = os.path.splitext(os.path.basename(source))[0]
        created_files = []
        for filters in all_combinations:
            # Формируем имя файла с расширением .xlsx
            safe_parts = [sanitize_filename(v) for v in filters.values()]
            suffix = "_".join(safe_parts) if safe_parts else "All"
            target_file = os.path.join(destination, f"{base_name}_{suffix}.xlsx")
            # Создаем файл
            created_file = create_filtered_file(source, target_file, valid_sheets, filters)
            if created_file is not None:
                created_files.append(created_file)
        
        # Вывод результатов
        if created_files:
            print(f"\nCreated {len(created_files)} files:")
            for file in created_files:
                print(f"  - {file}")
        else:
            print("Warning: No files created (no data matched the filters)")
        return True
    except KeyboardInterrupt:
        logger.info("Operation cancelled by user (Ctrl+C)")
        print("\nOperation cancelled by user (Ctrl+C)")
        return False
    except Exception as e:
        logger.exception("Unexpected error during file processing")
        print(f"Error: {str(e)}")
        return False