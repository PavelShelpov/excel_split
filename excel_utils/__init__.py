# Импорты для пакета excel_utils
try:
    from .analysis import get_all_sheets_headers, analyze_column
    from .filtering import get_all_combinations, select_categories_sequentially
    from .formatting import sanitize_filename, generate_short_filename
    from .workbook import create_filtered_file
    from .common import validate_row
    
    __all__ = [
        'get_all_sheets_headers',
        'analyze_column',
        'get_all_combinations',
        'select_categories_sequentially',
        'sanitize_filename',
        'generate_short_filename',
        'create_filtered_file',
        'validate_row'
    ]
    
    # Убираем логгирование из __init__.py
    # Логгирование должно происходить в main.py после полной инициализации
except Exception as e:
    # Добавляем импорт logging ДО использования
    import logging
    logger = logging.getLogger('excel_splitter')
    logger.error(f"Failed to initialize excel_utils package: {str(e)}")
    raise