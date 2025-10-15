import re
import os
import logging
from openpyxl.utils import get_column_letter

logger = logging.getLogger('excel_splitter')

def sanitize_filename(name):
    """Удаляет недопустимые символы из названия файла."""
    # Удаляем недопустимые символы
    name = re.sub(r'[\\/*?:"<>|]', '_', name)
    # Убираем лишние пробелы
    name = re.sub(r'\s+', ' ', name).strip()
    return name

def shorten_category_name(name):
    """
    Сокращает длинные названия категорий по правилам:
    - Для одного слова: первые 3 буквы
    - Для двух и более слов: первая буква каждого слова
    """
    if not name:
        return ""
    
    # Удаляем недопустимые символы
    name = sanitize_filename(name)
    
    # Если название короткое, оставляем как есть
    if len(name) <= 15:
        return name
    
    # Разбиваем на слова и берем первую букву каждого
    words = name.split()
    if len(words) > 1:
        # Берем первую букву каждого слова
        short_name = ''.join(word[0] for word in words if word)
    else:
        # Для одного слова берем первые 3 буквы
        short_name = name[:3]
    
    return short_name

def generate_short_filename(base_name, filters, max_length=150):
    """
    Генерирует короткое имя файла с учетом максимальной длины.
    Если длина превышает max_length, использует хэш для уникальности.
    """
    # Создаем список сокращенных названий категорий
    safe_parts = []
    for i, (col, value) in enumerate(filters.items()):
        if i == len(filters) - 1:  # Последняя категория - не сокращаем
            safe_parts.append(sanitize_filename(value))
        else:
            safe_parts.append(shorten_category_name(value))
    
    # Формируем суффикс
    suffix = "_".join(safe_parts) if safe_parts else "All"
    
    # Проверяем длину полного пути
    full_path = os.path.join(os.path.dirname(base_name), f"{os.path.basename(base_name)}_{suffix}.xlsx")
    
    # Если длина слишком большая, сокращаем
    if len(full_path) > max_length:
        logger.warning(f"Filename is too long ({len(full_path)} characters), shortening...")
        
        # Оставляем только последние N символов из суффикса
        max_suffix_length = max_length - len(base_name) - 5  # Учитываем '.xlsx' и '_'
        
        if max_suffix_length <= 0:
            # Если даже базовое имя слишком длинное, используем хэш
            import hashlib
            hash_suffix = hashlib.sha1(suffix.encode()).hexdigest()[:8]
            suffix = f"short_{hash_suffix}"
        else:
            # Сокращаем суффикс до допустимой длины
            suffix = suffix[:max_suffix_length]
    
    return f"{os.path.basename(base_name)}_{suffix}.xlsx"