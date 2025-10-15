import sys
from logging_config import logger  # Импортируем уже настроенный логгер
from cli.interface import main as cli_main

def run():
    """Точка входа в приложение. Может быть расширена для поддержки GUI в будущем"""
    logger.info("Application started")
    cli_main()
    
def check_dependencies():
    try:
        import openpyxl
        version = openpyxl.__version__
        # Проверяем минимальную версию
        major, minor, _ = map(int, version.split('.'))
        if major < 3:
            logger.error("OpenPyXL version 3.0+ is required. Current version: %s", version)
            return False
        return True
    except ImportError:
        logger.error("OpenPyXL is not installed. Please install it with 'pip install openpyxl'")
        return False

if __name__ == "__main__":
    try:
        if check_dependencies():
            run()
    except Exception as e:
        logger.exception("Critical error in application")
        sys.exit(1)