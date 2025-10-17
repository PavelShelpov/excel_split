import logging
from copy import copy

logger = logging.getLogger('excel_splitter')

def validate_row(row, headers, header_row_idx, filters):
    """Проверяет соответствие строки условиям фильтров."""
    logger.debug(f"Validating row: {row}, headers: {headers}, filters: {filters}")
    if not filters:
        logger.debug("No filters provided, row is valid")
        return True
    # Нормализуем заголовки к нижнему регистру
    normalized_headers = [str(header).lower() if header is not None else "" for header in headers]
    for col, value in filters.items():
        # Нормализуем имя колонки к нижнему регистру
        normalized_col = str(col).lower()
        try:
            col_index = normalized_headers.index(normalized_col)
            cell_value = row[col_index] if col_index < len(row) else None
            str_value = str(cell_value).strip() if cell_value is not None else ""
            str_filter = str(value).strip()
            logger.debug(f"Checking column '{col}': cell value='{str_value}', filter='{str_filter}'")
            # Сравниваем без учета регистра
            if str_value.lower() != str_filter.lower():
                logger.debug(f"Row does not match filter for column '{col}'")
                return False
        except ValueError:
            logger.warning(f"Column '{col}' not found in headers")
            return False
    logger.debug("Row matches all filters")
    return True

def copy_cell_style(source_cell, target_cell):
    """Копирует стили из исходной ячейки в целевую."""
    if source_cell.has_style:
        try:
            target_cell.font = copy(source_cell.font)
            target_cell.border = copy(source_cell.border)
            target_cell.fill = copy(source_cell.fill)
            target_cell.alignment = copy(source_cell.alignment)
            target_cell.number_format = source_cell.number_format
        except Exception as e:
            logger.debug(f"Error copying cell style: {str(e)}")