import pandas as pd
from file_handler import load_xlsx, save_xlsx
from filter_logic import get_column_names, get_unique_values, apply_filter
import os
import unicodedata
import re

def normalize_string(name: str) -> str:
    """
    Нормализует строку: приведение к нижнему регистру,
    удаление лишних пробелов в начале и конце, нормализация Unicode,
    удаление всех пробелов внутри, удаление спецсимволов (оставляем alnum и -, _, .).
    """
    if not isinstance(name, str):
        name = str(name)
    # Нормализуем Unicode
    normalized = unicodedata.normalize('NFKD', name)
    # Приводим к нижнему регистру
    lower = normalized.lower()
    # Удаляем лишние пробелы в начале и конце
    stripped = lower.strip()
    # Удаляем все пробелы внутри
    no_spaces = re.sub(r'\s+', '', stripped)
    # Оставляем только разрешённые символы (после удаления пробелов, пробел в разрешённых не нужен)
    safe_str = "".join(c for c in no_spaces if c.isalnum() or c in ('-', '_', '.'))
    return safe_str

def main():
    """
    Основная функция приложения для фильтрации XLSX файла по категориям в колонке.
    """
    print("--- Приложение фильтрации XLSX файла по категориям ---")

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

    # --- Нормализация имён колонок ---
    original_columns = get_column_names(df)
    # Функция нормализации для колонок
    def normalize_col_name(name: str) -> str:
        return normalize_string(name)

    # Создаём словарь для сопоставления нормализованного имени с оригинальным
    normalized_to_original = {normalize_col_name(col): col for col in original_columns}

    print("\nДоступные колонки:")
    for i, col in enumerate(original_columns):
        print(f"{i + 1}. {col}")

    selected_column = None
    while selected_column is None:
        try:
            col_choice = input("\nВведите номер или название колонки для фильтрации: ").strip()
            # Проверяем, является ли ввод числом и соответствует ли номер колонке
            if col_choice.isdigit():
                col_index = int(col_choice) - 1
                if 0 <= col_index < len(original_columns):
                    selected_column = original_columns[col_index]
                    break
                else:
                    print("Неверный номер колонки. Пожалуйста, попробуйте снова.")
            # Проверяем, является ли ввод названием колонки (с нормализацией)
            else:
                normalized_choice = normalize_col_name(col_choice)
                if normalized_choice in normalized_to_original:
                    selected_column = normalized_to_original[normalized_choice]
                    break
                else:
                    print("Колонка с таким названием не найдена. Пожалуйста, попробуйте снова.")
        except (ValueError, IndexError):
            print("Неверный ввод. Пожалуйста, введите номер или название колонки.")

    print(f"Выбрана колонка: '{selected_column}'")

    # 4. Получение всех уникальных значений (категорий) из выбранной колонки
    print(f"\nУникальные значения в колонке '{selected_column}' (идентичные нормализованы):")
    unique_values = get_unique_values(df, selected_column)
    
    # --- НОВАЯ ЛОГИКА: Группировка уникальных значений по нормализации ---
    # Словарь: нормализованная_строка -> [список_оригинальных_значений]
    normalized_groups = {}
    for val in unique_values:
        if pd.isna(val):
            # Для NaN используем специальный ключ, так как он не нормализуется как строка
            norm_key = "__NAN__" # Используем уникальный ключ для NaN
        else:
            norm_key = normalize_string(str(val))
        if norm_key not in normalized_groups:
            normalized_groups[norm_key] = []
        normalized_groups[norm_key].append(val)

    # Создаём список пар: (нормализованный_ключ, первое_оригинальное_значение_из_группы)
    display_pairs = []
    for norm_key, original_vals in normalized_groups.items():
        # Берём первое оригинальное значение из группы для отображения
        first_original = original_vals[0]
        display_pairs.append((norm_key, first_original))

    # Создаём список нормализованных ключей и список отображаемых имён для удобства
    display_keys = [pair[0] for pair in display_pairs]
    display_names = [pair[1] for pair in display_pairs]

    for i, display_name in enumerate(display_names):
        # Отображаем первое оригинальное имя из группы
        print(f"{i + 1}. {display_name}")

    # 5. Пользователю предлагается выбор: фильтровать по всем или нескольким категориям
    print("\n--- Выбор категорий для фильтрации ---")
    choice = None
    while choice not in ['все', 'несколько']:
        choice_input = input("Фильтровать по 'всем' категориям или 'несколько'? Введите 'все' или 'несколько': ").strip().lower()
        if choice_input in ['все', 'всем', 'all']:
            choice = 'все'
        elif choice_input in ['несколько', 'некоторые', 'several', 'some']:
            choice = 'несколько'
        else:
            print("Пожалуйста, введите 'все' или 'несколько'.")

    print(f"Выбран режим фильтрации: {choice}")

    # Определение списка нормализованных ключей для фильтрации
    keys_to_filter = []
    if choice == 'все':
        # Добавляем все нормализованные ключи
        keys_to_filter = list(normalized_groups.keys())
    else: # choice == 'несколько'
        selected_indices = set() # Используем set для уникальности
        print("\nВведите номера нормализованных категорий для фильтрации (например, 1 3 5 или 1-3 5). Нажмите Enter, когда закончите:")
        while True:
            selection_input = input("Выберите номер(а) или диапазон(ы) через пробел (или Enter для завершения): ").strip()
            if not selection_input:
                break # Пользователь закончил ввод

            parts = selection_input.split()
            valid_selection = True
            for part in parts:
                if '-' in part:
                    # Обработка диапазона (например, 1-3)
                    try:
                        start, end = map(int, part.split('-'))
                        if start > end:
                            start, end = end, start # Поменять местами, если начальный больше
                        for num in range(start, end + 1):
                            idx = num - 1 # Индексация с 0
                            if 0 <= idx < len(display_keys):
                                selected_indices.add(idx)
                            else:
                                print(f"Номер {num} вне диапазона. Пропущен.")
                                valid_selection = False
                    except ValueError:
                        print(f"Неверный формат диапазона: {part}. Пропущен.")
                        valid_selection = False
                else:
                    # Обработка одиночного номера
                    try:
                        num = int(part)
                        idx = num - 1 # Индексация с 0
                        if 0 <= idx < len(display_keys):
                            selected_indices.add(idx)
                        else:
                            print(f"Номер {num} вне диапазона. Пропущен.")
                            valid_selection = False
                    except ValueError:
                        print(f"Неверный номер: {part}. Пропущен.")
                        valid_selection = False

            if not valid_selection:
                print("Пожалуйста, повторите ввод.")

        # Формирование списка нормализованных ключей: для каждого выбранного индекса
        keys_to_filter = [display_keys[idx] for idx in selected_indices]

        if not keys_to_filter:
            print("Не выбрано ни одной категории. Программа завершена.")
            return

        print(f"Выбраны следующие категории для фильтрации (с учётом нормализации):")
        for norm_key in keys_to_filter:
            first_original = normalized_groups[norm_key][0]
            print(f" - {first_original} (и {len(normalized_groups[norm_key]) - 1} других)")


    # 6. Цикл по выбранным нормализованным ключам (группам категорий)
    print(f"\nНачинается процесс фильтрации по {len(keys_to_filter)} выбранной(ым) группе(ам) категорий...")
    saved_files_count = 0
    used_normalized_filenames = set() # Отслеживаем уже использованные НОРМАЛИЗОВАННЫЕ имена файлов
    for norm_key in keys_to_filter:
        original_vals_in_group = normalized_groups[norm_key]
        first_original_display_name = original_vals_in_group[0]
        print(f"\nФильтрация по группе: '{first_original_display_name}' (включая {len(original_vals_in_group) - 1} других)...")

        # 7. Фильтрация файла по группе оригинальных значений
        # Создаём маску для фильтрации: строка в selected_column соответствует ЛЮБОМУ значению в группе
        # Используем isin() для проверки принадлежности к списку
        # Обрабатываем NaN отдельно, если он есть в группе
        mask = df[selected_column].isin(original_vals_in_group)
        if norm_key == "__NAN__":
             # Если группа - это NaN, добавляем строки, где selected_column IS NULL
            mask = mask | df[selected_column].isna()

        filtered_df = df[mask].copy()

        print(f"  Найдено {filtered_df.shape[0]} строк для группы '{first_original_display_name}'.")

        # 8. Сохранение файла для текущей группы
        # Имя файла формируется на основе ПЕРВОГО ОРИГИНАЛЬНОГО значения в группе
        # Очищаем first_original_display_name от недопустимых символов (не нормализуя!)
        safe_category_str = "".join(c for c in first_original_display_name if c.isalnum() or c in (' ', '-', '_', '.')).rstrip()
        if not safe_category_str:
             # Если после очистки имя пустое, используем заглушку
            safe_category_str = "unnamed_category"

        # --- Исправление для уникальности имён файлов ---
        # Нормализуем safe_category_str для проверки конфликта
        normalized_safe_category_str = normalize_string(safe_category_str)
        # Получаем имя исходного файла без расширения
        original_filename = os.path.splitext(os.path.basename(file_path))[0]
        base_output_filename = f"{original_filename}__{safe_category_str}.xlsx"
        output_filename = base_output_filename

        # Проверяем, существует ли такое НОРМАЛИЗОВАННОЕ имя уже, и если да, добавляем суффикс к ОРИГИНАЛЬНОМУ имени
        counter = 1
        while normalized_safe_category_str in used_normalized_filenames:
            name_part, ext_part = os.path.splitext(base_output_filename)
            output_filename = f"{name_part}_{counter}{ext_part}"
            # Пересчитываем нормализованное имя для нового output_filename
            temp_safe_part = "".join(c for c in f"{safe_category_str}_{counter}" if c.isalnum() or c in (' ', '-', '_', '.')).rstrip()
            normalized_safe_category_str = normalize_string(temp_safe_part)
            counter += 1

        # Добавляем НОРМАЛИЗОВАННОЕ имя файла в список использованных
        used_normalized_filenames.add(normalized_safe_category_str)

        try:
            save_xlsx(filtered_df, output_filename)
            print(f"  Файл для группы '{first_original_display_name}' успешно сохранен как: {output_filename}")
            saved_files_count += 1
        except Exception as e:
            print(f"  Ошибка при сохранении файла для группы '{first_original_display_name}': {e}")

    print(f"\n--- Процесс завершен ---")
    print(f"Всего уникальных нормализованных категорий в колонке '{selected_column}': {len(display_keys)}")
    print(f"Выбрано для фильтрации: {len(keys_to_filter)} групп(ы)")
    print(f"Успешно сохранено файлов: {saved_files_count}")

if __name__ == "__main__":
    main()