from core.processing import process_file

def main():
    """Главный цикл программы: обработка файлов."""
    while True:
        success = process_file()
        # Спрашиваем, хочет ли пользователь продолжить
        if success:
            cont = input("\nDo you want to process another file? (y/n): ").strip().lower()
            if cont != 'y':
                print("Program terminated by user")
                break
        else:
            cont = input("\nDo you want to try again? (y/n): ").strip().lower()
            if cont != 'y':
                print("Program terminated by user")
                break