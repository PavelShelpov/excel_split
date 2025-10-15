import sys
from cli.interface import main as cli_main

def run():
    """Точка входа в приложение. Может быть расширена для поддержки GUI в будущем"""
    cli_main()

if __name__ == "__main__":
    run()