# copy_excel.py
import os
import shutil

def main():
    print("=== Копирование Excel-файла ===")
    
    # Запрос путей с обработкой пробелов
    source = input("Введите полный путь к исходному Excel-файлу: ").strip('"')
    destination = input("Введите путь к целевой директории: ").strip('"')
    
    # Коррекция пути для Windows (удаление лишних кавычек)
    source = os.path.normpath(source)
    destination = os.path.normpath(destination)
    
    # Проверки
    if not os.path.exists(source):
        print(f"Ошибка: Исходный файл не найден: {source}")
        return
    
    if not os.path.isfile(source):
        print(f"Ошибка: Указанный путь не является файлом: {source}")
        return
    
    # Создание директории
    os.makedirs(destination, exist_ok=True)
    
    # Копирование
    try:
        shutil.copy2(source, destination)
        print(f"Успешно скопировано в: {os.path.join(destination, os.path.basename(source))}")
    except Exception as e:
        print(f"Ошибка копирования: {str(e)}")

if __name__ == "__main__":
    main()