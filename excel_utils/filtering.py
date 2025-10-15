from .analysis import analyze_column
import logging

logger = logging.getLogger('excel_splitter')

def get_all_combinations(source, valid_sheets, hierarchy_columns, filters=None, level=0):
    """Возвращает все возможные комбинации фильтров, включая частичные уровни."""
    if filters is None:
        filters = {}
    if level >= len(hierarchy_columns):
        return [filters.copy()]
    column = hierarchy_columns[level]
    logger.debug(f"Getting categories for column {column} with filters {filters}")
    categories = analyze_column(source, valid_sheets, column, filters)
    combinations = []
    # Добавляем комбинации для текущего уровня без добавления следующих уровней
    for category in categories:
        new_filters = filters.copy()
        new_filters[column] = category
        combinations.append(new_filters.copy())
    # Добавляем комбинации для следующих уровней
    for category in categories:
        new_filters = filters.copy()
        new_filters[column] = category
        combinations.extend(get_all_combinations(source, valid_sheets, hierarchy_columns, new_filters, level + 1))
    return combinations

def select_categories_sequentially(source, valid_sheets, hierarchy_columns):
    """Последовательно запрашивает выбор категорий у пользователя с отображением вариантов для каждой комбинации."""
    logger.info("Starting sequential category selection")
    all_combinations = []
    
    def generate_combinations(level, current_filters, include_all=False):
        if level >= len(hierarchy_columns):
            all_combinations.append(current_filters.copy())
            return
        column = hierarchy_columns[level]
        categories = analyze_column(source, valid_sheets, column, current_filters)
        if not categories:
            return
        # Проверяем, является ли текущий уровень последним
        is_last_level = (level == len(hierarchy_columns) - 1)
        # Если это не первый уровень и есть предыдущие фильтры
        if level > 0:
            logger.debug(f"Current filters: {current_filters}")
            print(f"\nCurrent filters:")
            for col, value in current_filters.items():
                print(f"  - {col}: {value}")
            # Если не последний уровень, спрашиваем, хочет ли пользователь выбрать все комбинации
            if not is_last_level:
                while True:
                    all_comb = input(f"Include all categories from '{hierarchy_columns[level]}' for current filters? (y/n): ").strip().lower()
                    if all_comb == 'y':
                        # Анализируем все возможные комбинации
                        for category in categories:
                            new_filters = current_filters.copy()
                            new_filters[column] = category
                            generate_combinations(level + 1, new_filters, True)
                        return
                    elif all_comb == 'n':
                        break
                    else:
                        print("Please enter 'y' or 'n'")
        # Выводим доступные категории с номерами
        print(f"\nAvailable categories for column '{column}':")
        for i, cat in enumerate(categories, 1):
            print(f"  {i}. {cat}")
        print("  b. Назад")
        print("  c. Отмена")
        # Запрашиваем выбор
        while True:
            selection = input(f"Enter categories for '{column}' (comma-separated numbers, 'all', 'b' for back, 'c' for cancel): ").strip()
            # Обработка специальных команд
            if selection.lower() in ["c", "cancel", "отмена"]:
                print("Operation cancelled by user")
                return
            if selection.lower() in ["b", "back", "назад"]:
                return
            # Обработка "all" - выбираем все категории
            if selection.lower() == "all":
                selected_categories = categories
                break
            # Обработка номеров
            user_categories = []
            invalid_inputs = []
            for item in selection.split(","):
                item = item.strip()
                if item.isdigit():
                    idx = int(item) - 1
                    if 0 <= idx < len(categories):
                        user_categories.append(categories[idx])
                    else:
                        invalid_inputs.append(item)
                else:
                    user_categories.append(item)
            # Проверка валидности
            invalid_categories = [cat for cat in user_categories if cat not in categories]
            if invalid_categories or invalid_inputs:
                invalid_list = invalid_categories + invalid_inputs
                print(f"Error: Invalid categories: {', '.join(invalid_list)}")
                continue
            selected_categories = user_categories
            break
        # Обрабатываем выбор
        for category in selected_categories:
            new_filters = current_filters.copy()
            new_filters[column] = category
            # Если пользователь выбрал "all" для предыдущего уровня
            if include_all and level > 0:
                generate_combinations(level + 1, new_filters, True)
            else:
                generate_combinations(level + 1, new_filters)
    
    # Начинаем генерацию комбинаций с первого уровня
    generate_combinations(0, {})
    # Добавляем частичные уровни фильтрации
    final_combinations = []
    for filters in all_combinations:
        # Добавляем фильтры всех подуровней
        for i in range(len(hierarchy_columns)):
            partial_filter = {}
            for j in range(i + 1):
                if hierarchy_columns[j] in filters:
                    partial_filter[hierarchy_columns[j]] = filters[hierarchy_columns[j]]
            final_combinations.append(partial_filter)
    # Удаляем дубликаты частичных фильтров
    unique_combinations = []
    seen = set()
    for filters in final_combinations:
        filter_tuple = tuple(sorted(filters.items()))
        if filter_tuple not in seen:
            seen.add(filter_tuple)
            unique_combinations.append(filters)
    return unique_combinations