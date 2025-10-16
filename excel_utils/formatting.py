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

def shorten_category_name(name, is_last=False):
    """
    Сокращает длинные названия категорий по новым правилам:
    - Для одной категории: полное название
    - Для нескольких категорий: все, кроме последней, сокращаются по правилам:
        * Две первые буквы первого слова (первая в верхнем, вторая в нижнем регистре)
        * Первая буква каждого последующего слова в верхнем регистре
    """
    if not name or is_last:
        return name
    
    # Удаляем недопустимые символы
    name = sanitize_filename(name)
    # Нормализуем пробелы
    name = re.sub(r'\s+', ' ', name).strip()
    
    # Делаем сокращение только если длина превышает 3 символа
    if len(name) <= 3:
        return name
    
    # Разбиваем на слова
    words = name.split()
    if len(words) == 0:
        return name
    
    # Формируем сокращение
    if len(words[0]) >= 2:
        # Берем первые две буквы первого слова (первая заглавная, вторая строчная)
        short_name = words[0][0].upper() + words[0][1].lower()
    else:
        # Если первое слово короткое, берем первую букву в верхнем регистре
        short_name = words[0][0].upper() if len(words[0]) > 0 else ""
    
    # Добавляем первые буквы остальных слов в верхнем регистре
    for word in words[1:]:
        if word:
            short_name += word[0].upper()
    
    return short_name

def generate_short_filename(base_name, filters, max_length=150):
    """
    Генерирует короткое имя файла с учетом максимальной длины.
    Если длина превышает max_length, использует хэш для уникальности.
    """
    # Создаем список сокращенных названий категорий
    safe_parts = []
    category_names = list(filters.values())
    
    for i, value in enumerate(category_names):
        is_last = (i == len(category_names) - 1)
        short_name = shorten_category_name(value, is_last)
        safe_parts.append(sanitize_filename(short_name))
    
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