import logging
import sys

def setup_logging():
    """Настраивает систему логирования для приложения."""
    logger = logging.getLogger('excel_splitter')
    logger.setLevel(logging.INFO)
    
    # Создаем обработчик для вывода в консоль
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    
    # Формат логов
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    console_handler.setFormatter(formatter)
    
    # Добавляем обработчик, если он еще не добавлен
    if not logger.handlers:
        logger.addHandler(console_handler)
    
    return logger

# Создаем глобальный логгер
logger = setup_logging()