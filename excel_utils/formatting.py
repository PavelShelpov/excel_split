import re

def sanitize_filename(name):
    """Удаляет недопустимые символы из названия файла."""
    return re.sub(r'[\\/*?:"<>|]', '_', name)