import sys
import os

def run_cli():
    """Запускает CLI версию приложения"""
    from cli.interface import main as cli_main
    cli_main()

def run_gui():
    """Запускает GUI версию приложения"""
    from gui.main import launch_gui
    launch_gui()

def main():
    """Точка входа в приложение с выбором режима работы"""
    if len(sys.argv) > 1 and sys.argv[1] == "cli":
        run_cli()
    elif len(sys.argv) > 1 and sys.argv[1] == "gui":
        run_gui()
    else:
        print("Excel Splitter")
        print("1. Command Line Interface (CLI)")
        print("2. Graphical User Interface (GUI)")
        print("3. Exit")
        
        choice = input("Enter your choice (1/2/3): ").strip()
        
        if choice == "1":
            run_cli()
        elif choice == "2":
            run_gui()
        elif choice == "3":
            print("Exiting...")
        else:
            print("Invalid choice. Please enter 1, 2, or 3.")
            main()

if __name__ == "__main__":
    main()