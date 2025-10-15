import os
import re
import logging
import openpyxl
from copy import copy
from contextlib import contextmanager
from openpyxl.worksheet.table import Table, TableStyleInfo
from excel_utils.common import validate_row

logger = logging.getLogger('excel_splitter')

def get_column_letter(col_idx):
    """Конвертирует индекс столбца в букву (1 -> A, 26 -> Z, 27 -> AA и т.д.)"""
    letters = []
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        letters.append(chr(65 + remainder))
    return ''.join(reversed(letters))

def is_openpyxl_new_version():
    """Проверяет, является ли версия openpyxl новой (>= 3.0)"""
    try:
        from openpyxl import __version__
        major_version = int(__version__.split('.')[0])
        return major_version >= 3
    except (ImportError, ValueError):
        return False

@contextmanager
def safe_workbook(file_path, read_only=False):
    """Контекстный менеджер для безопасной работы с файлами Excel."""
    wb = None
    try:
        logger.debug(f"Opening workbook: {file_path}")
        wb = openpyxl.load_workbook(file_path, read_only=read_only)
        yield wb
    finally:
        if wb:
            try:
                wb.close()
                logger.debug(f"Workbook closed: {file_path}")
            except Exception as e:
                logger.error(f"Error closing workbook: {str(e)}")

def create_filtered_file(source, target, valid_sheets, filters):
    """Создаёт файл с фильтрацией по комбинации условий."""
    logger.info(f"Creating filtered file: {target} with filters {filters}")
    
    # Добавлена проверка на пустой фильтр
    if not filters:
        logger.info("Empty filters, copying all data")
    
    try:
        # Всегда сохраняем как .xlsx
        if target.lower().endswith('.xlsm'):
            logger.debug("Converting .xlsm to .xlsx format")
            target = target[:-5] + '.xlsx'
        
        with safe_workbook(source, read_only=False) as wb_source:
            wb_new = openpyxl.Workbook()
            wb_new.remove(wb_new.active)
            has_data = False  # Флаг наличия данных
            logger.debug(f"Processing {len(wb_source.sheetnames)} sheets")
            
            for sheet_name in wb_source.sheetnames:
                ws_source = wb_source[sheet_name]
                # Игнорируем скрытые листы
                if ws_source.sheet_state != 'visible':
                    logger.debug(f"Skipping hidden sheet: {sheet_name}")
                    continue
                
                ws_new = wb_new.create_sheet(title=sheet_name)
                logger.debug(f"Processing sheet: {sheet_name}")
                
                # Копирование ширины столбцов
                if hasattr(ws_source, 'column_dimensions'):
                    for col_letter, dim in ws_source.column_dimensions.items():
                        try:
                            ws_new.column_dimensions[col_letter].width = dim.width
                        except Exception as e:
                            logger.debug(f"Error copying column width for {col_letter}: {str(e)}")
                
                # Копирование высоты строк
                if hasattr(ws_source, 'row_dimensions'):
                    for row_idx, dim in ws_source.row_dimensions.items():
                        try:
                            ws_new.row_dimensions[row_idx].height = dim.height
                        except Exception as e:
                            logger.debug(f"Error copying row height for {row_idx}: {str(e)}")
                
                # Копирование объединенных ячеек
                if hasattr(ws_source, 'merged_cells'):
                    for merged_cell in ws_source.merged_cells.ranges:
                        try:
                            ws_new.merge_cells(str(merged_cell))
                        except Exception as e:
                            logger.debug(f"Error copying merged cells: {str(e)}")
                
                # Копирование условного форматирования
                if hasattr(ws_source, 'conditional_formatting'):
                    for cf in ws_source.conditional_formatting:
                        try:
                            # Проверяем тип объекта условного форматирования
                            if hasattr(cf, 'cfRule') and hasattr(cf, 'cfRules'):
                                # Новые версии openpyxl
                                for rule in cf.cfRules:
                                    ws_new.conditional_formatting.add(cf._get_range(), rule)
                            elif hasattr(cf, 'rules'):
                                # Средние версии
                                for rule in cf.rules:
                                    ws_new.conditional_formatting.add(cf._get_range(), rule)
                            else:
                                # Старые версии
                                ws_new.conditional_formatting.add(cf._range, cf)
                        except Exception as e:
                            logger.debug(f"Error copying conditional formatting: {str(e)}")
                
                if sheet_name in valid_sheets:
                    headers, header_row_idx = valid_sheets[sheet_name]
                    logger.debug(f"Headers for sheet {sheet_name}: {headers}")
                    logger.debug(f"Header row index: {header_row_idx}")
                    
                    # 1. Технические строки выше таблицы
                    for row_idx in range(1, header_row_idx):
                        for col_idx in range(1, ws_source.max_column + 1):
                            try:
                                cell = ws_source.cell(row=row_idx, column=col_idx)
                                if cell.value is not None or cell.has_style:
                                    new_cell = ws_new.cell(row=row_idx, column=col_idx, value=cell.value)
                                    if cell.has_style:
                                        # Копируем отдельные стили вместо общего
                                        try:
                                            new_cell.font = copy(cell.font)
                                            new_cell.border = copy(cell.border)
                                            new_cell.fill = copy(cell.fill)
                                            new_cell.alignment = copy(cell.alignment)
                                            new_cell.number_format = cell.number_format
                                        except Exception as e:
                                            logger.debug(f"Error copying individual style properties: {str(e)}")
                            except Exception as e:
                                logger.debug(f"Error copying cell at row {row_idx}, col {col_idx}: {str(e)}")
                    
                    # 2. Заголовки
                    for col_idx in range(1, ws_source.max_column + 1):
                        try:
                            cell = ws_source.cell(row=header_row_idx, column=col_idx)
                            if cell.value is not None or cell.has_style:
                                new_cell = ws_new.cell(row=header_row_idx, column=col_idx, value=cell.value)
                                if cell.has_style:
                                    # Копируем отдельные стили вместо общего
                                    try:
                                        new_cell.font = copy(cell.font)
                                        new_cell.border = copy(cell.border)
                                        new_cell.fill = copy(cell.fill)
                                        new_cell.alignment = copy(cell.alignment)
                                        new_cell.number_format = cell.number_format
                                    except Exception as e:
                                        logger.debug(f"Error copying individual style properties for header: {str(e)}")
                        except Exception as e:
                            logger.debug(f"Error copying header at col {col_idx}: {str(e)}")
                    
                    # 3. Фильтрация данных
                    new_row_idx = header_row_idx + 1
                    filtered_count = 0
                    for row_idx in range(header_row_idx + 1, ws_source.max_row + 1):
                        try:
                            row = ws_source[row_idx]
                            # Добавлена проверка на пустой фильтр
                            if not filters:
                                should_include = True
                            else:
                                should_include = validate_row([cell.value for cell in row], headers, header_row_idx, filters)
                            
                            if should_include:
                                filtered_count += 1
                                for col_idx in range(1, ws_source.max_column + 1):
                                    try:
                                        source_cell = ws_source.cell(row=row_idx, column=col_idx)
                                        if source_cell.value is not None or source_cell.has_style:
                                            new_cell = ws_new.cell(row=new_row_idx, column=col_idx, value=source_cell.value)
                                            if source_cell.has_style:
                                                # Копируем отдельные стили вместо общего
                                                try:
                                                    new_cell.font = copy(source_cell.font)
                                                    new_cell.border = copy(source_cell.border)
                                                    new_cell.fill = copy(source_cell.fill)
                                                    new_cell.alignment = copy(source_cell.alignment)
                                                    new_cell.number_format = source_cell.number_format
                                                except Exception as e:
                                                    logger.debug(f"Error copying individual style properties for data cell: {str(e)}")
                                    except Exception as e:
                                        logger.debug(f"Error copying data cell at row {row_idx}, col {col_idx}: {str(e)}")
                                new_row_idx += 1
                        except Exception as e:
                            logger.debug(f"Error processing row {row_idx}: {str(e)}")
                    
                    logger.debug(f"Filtered {filtered_count} rows out of {ws_source.max_row - header_row_idx} possible")
                    
                    # Проверка наличия данных
                    if new_row_idx > header_row_idx + 1:
                        has_data = True
                        # Форматируем данные как таблицу
                        last_col_letter = get_column_letter(ws_source.max_column)
                        table_range = f"A{header_row_idx}:{last_col_letter}{new_row_idx-1}"
                        
                        # Создаем таблицу
                        table = Table(displayName=f"Table{sheet_name}", ref=table_range)
                        
                        # Создаем стиль таблицы с только поддерживаемыми параметрами
                        try:
                            # Пытаемся использовать полный набор параметров
                            style = TableStyleInfo(
                                name="TableStyleMedium9",
                                showFirstColumn=False,
                                showLastColumn=False,
                                showRowHeaders=True,
                                showColumnHeaders=True
                            )
                        except TypeError:
                            # Используем базовые параметры, если некоторые не поддерживаются
                            try:
                                style = TableStyleInfo(
                                    name="TableStyleMedium9",
                                    showFirstColumn=False,
                                    showLastColumn=False,
                                    showColumnHeaders=True
                                )
                            except TypeError:
                                # Используем минимальный набор параметров
                                style = TableStyleInfo(
                                    name="TableStyleMedium9"
                                )
                        
                        table.tableStyleInfo = style
                        ws_new.add_table(table)
                    else:
                        # Удаляем лист без данных
                        wb_new.remove(ws_new)
                        logger.debug(f"Removed sheet {sheet_name} due to no matching data")
                        continue  # Переходим к следующему листу
                
                else:
                    logger.debug(f"Copying entire sheet {sheet_name} without filtering")
                    for row_idx in range(1, ws_source.max_row + 1):
                        for col_idx in range(1, ws_source.max_column + 1):
                            try:
                                cell = ws_source.cell(row=row_idx, column=col_idx)
                                if cell.value is not None or cell.has_style:
                                    new_cell = ws_new.cell(row=row_idx, column=col_idx, value=cell.value)
                                    if cell.has_style:
                                        # Копируем отдельные стили вместо общего
                                        try:
                                            new_cell.font = copy(cell.font)
                                            new_cell.border = copy(cell.border)
                                            new_cell.fill = copy(cell.fill)
                                            new_cell.alignment = copy(cell.alignment)
                                            new_cell.number_format = cell.number_format
                                        except Exception as e:
                                            logger.debug(f"Error copying individual style properties: {str(e)}")
                            except Exception as e:
                                logger.debug(f"Error copying cell at row {row_idx}, col {col_idx}: {str(e)}")
            
            if not has_data:
                logger.warning("No data matched the filters, file not created")
                return None
            
            # Удаляем целевой файл, если он существует
            if os.path.exists(target):
                logger.info(f"Removing existing target file: {target}")
                os.remove(target)
            
            # Сохраняем как .xlsx
            logger.info(f"Saving filtered file: {target}")
            wb_new.save(target)
            return target
    except Exception as e:
        logger.exception(f"Error during filtering: {str(e)}")
        raise ValueError(f"Error during filtering: {str(e)}")