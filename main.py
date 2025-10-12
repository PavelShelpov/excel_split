import pandas as pd
from file_handler import load_xlsx, save_xlsx
from filter_logic import get_column_names, get_unique_values, apply_filter

def main():
    """
    Основная функция приложения для фильтрации XLSX файла.
    """
    print("--- Приложение фильтрации XLSX файла ---")

    # 1. Пользователь выбирает файл (ввод пути в консоли)
    file_path = input("Введите путь к XLSX файлу: ").strip()
    if not file_path:
        print("Путь к файлу не может быть пустым.")
        return

    # 2. Загрузка файла
    try:
        print(f"Загрузка файла: {file_path}")
        df = load_xlsx(file_path)
        print(f"Файл успешно загружен. Размер: {df.shape[0]} строк, {df.shape[1]} колонок.")
    except FileNotFoundError as fnf_error:
        print(fnf_error)
        return
    except ValueError as ve:
        print(ve)
        return
    except Exception as e:
        print(f"Произошла непредвиденная ошибка при загрузке: {e}")
        return

    # 3. Пользователь выбирает колонку
    columns = get_column_names(df)
    print("\nДоступные колонки:")
    for i, col in enumerate(columns):
        print(f"{i + 1}. {col}")

    while True:
        try:
            col_choice = input("\nВведите номер или название колонки для фильтрации: ").strip()
            # Проверяем, является ли ввод числом и соответствует ли номер колонке
            if col_choice.isdigit():
                col_index = int(col_choice) - 1
                if 0 <= col_index < len(columns):
                    selected_column = columns[col_index]
                    break
                else:
                    print("Неверный номер колонки. Пожалуйста, попробуйте снова.")
            # Проверяем, является ли ввод названием колонки
            elif col_choice in columns:
                selected_column = col_choice
                break
            else:
                print("Колонка с таким названием не найдена. Пожалуйста, попробуйте снова.")
        except (ValueError, IndexError):
            print("Неверный ввод. Пожалуйста, введите номер или название колонки.")

    print(f"Выбрана колонка: '{selected_column}'")

    # 4. Пользователь выбирает категорию
    unique_values = get_unique_values(df, selected_column)
    print(f"\nУникальные значения в колонке '{selected_column}':")
    for i, val in enumerate(unique_values):
        # Отображаем NaN как строку 'NaN' для ясности ввода
        display_val = "NaN" if pd.isna(val) else val
        print(f"{i + 1}. {display_val}")

    # Инициализируем selected_category до цикла
    selected_category = None
    while True:
        try:
            cat_choice = input(f"\nВведите номер или значение категории из '{selected_column}': ").strip()
            # Проверяем, является ли ввод числом и соответствует ли номер значению
            if cat_choice.isdigit():
                val_index = int(cat_choice) - 1
                if 0 <= val_index < len(unique_values):
                    selected_category = unique_values[val_index]
                    break
                else:
                    print("Неверный номер категории. Пожалуйста, попробуйте снова.")
            # Проверяем, является ли ввод значением (строка или число)
            else:
                # Для NaN ввод должен быть 'NaN'
                if cat_choice == "NaN":
                    selected_category = pd.NA
                else:
                    # Пытаемся найти строку/число в списке уникальных значений
                    found = False
                    for val in unique_values:
                        # Сравниваем как строки, но учитываем pd.NA
                        if pd.isna(val) and pd.isna(pd.NA): # Этот случай не сработает как ожидалось
                            # Правильная проверка: если cat_choice == "NaN", то уже выше selected_category = pd.NA
                            # Иначе ищем по строковому представлению
                            continue # Пропускаем, так как NaN уже обработан
                        elif str(val) == cat_choice:
                            selected_category = val
                            found = True
                            break
                    if found:
                        break
                    else:
                        print(f"Категория '{cat_choice}' не найдена. Пожалуйста, попробуйте снова.")

        except (ValueError, IndexError):
            print("Неверный ввод. Пожалуйста, введите номер или значение категории.")

    # Проверка на случай, если цикл каким-то образом завершится без присвоения
    # (хотя по логике этого не должно произойти)
    if selected_category is None:
        print("Не удалось определить выбранную категорию. Программа завершена.")
        return

    # Обработка NaN: pandas не может сравнить NaN с NaN через ==, используем pd.isna для фильтрации
    if pd.isna(selected_category):
        print(f"Выбрана категория: 'NaN' (пустое значение)")
    else:
        print(f"Выбрана категория: '{selected_category}'")


    # 5. Фильтрация файла по заданной категории
    print("\nПрименение фильтра...")
    # Если выбрана категория NaN, используем специальную логику
    if pd.isna(selected_category):
        filtered_df = df[df[selected_column].isna()].copy()
    else:
        filtered_df = apply_filter(df, selected_column, selected_category)

    print(f"Фильтрация завершена. Найдено {filtered_df.shape[0]} строк.")

    # 6. Данные, не попавшие в фильтр, "удаляются" (не сохраняются в filtered_df)

    # 7. Сохранение файла
    output_filename = input("\nВведите имя для сохраняемого файла (без расширения .xlsx): ").strip()
    if not output_filename:
        output_filename = "filtered_output"
    output_path = f"{output_filename}.xlsx"

    try:
        save_xlsx(filtered_df, output_path)
        print(f"\nФайл успешно сохранен как: {output_path}")
    except Exception as e:
        print(f"Ошибка при сохранении файла: {e}")

if __name__ == "__main__":
    main()