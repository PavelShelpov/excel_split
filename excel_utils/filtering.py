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
    categories = analyze_column(source, valid_sheets, column, filters)
    combinations = []
    
    # Добавляем комбинации для текущего уровня
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
    
    def generate_combinations(level, current_filters):
        """Рекурсивная функция генерации комбинаций с явным выбором режима для каждого уровня"""
        if level >= len(hierarchy_columns):
            all_combinations.append(current_filters.copy())
            return
            
        column = hierarchy_columns[level]
        categories = analyze_column(source, valid_sheets, column, current_filters)
        
        if not categories:
            logger.warning(f"No categories found for column '{column}' at level {level}")
            return
        
        # Показываем текущие фильтры, если это не первый уровень
        if level > 0:
            print(f"\nCurrent filters:")
            for col, value in current_filters.items():
                print(f"  - {col}: {value}")
        
        # Выводим информацию о текущем уровне
        print(f"\n--- Level {level + 1}/{len(hierarchy_columns)} ---")
        print(f"Column for filtering: '{column}'")
        
        # Спрашиваем, хочет ли пользователь выбрать все категории для этого уровня
        while True:
            all_choice = input(f"Include all categories for this level? (y/n): ").strip().lower()
            if all_choice == 'y':
                # Обрабатываем выбор "all"
                logger.info(f"User chose 'all' for column '{column}' at level {level}")
                for category in categories:
                    new_filters = current_filters.copy()
                    new_filters[column] = category
                    generate_combinations(level + 1, new_filters)
                return
            elif all_choice == 'n':
                # Обрабатываем точечный выбор
                break
            else:
                print("Please enter 'y' or 'n'")
        
        # Выводим доступные категории с номерами
        print(f"\nAvailable categories for column '{column}':")
        for i, cat in enumerate(categories, 1):
            print(f"  {i}. {cat}")
        
        print("  a. All (this level only)")
        print("  b. Назад")
        print("  c. Отмена")
        
        # Запрашиваем выбор
        while True:
            selection = input(f"Enter categories for '{column}' (comma-separated numbers, 'a' for all this level, 'b' for back, 'c' for cancel): ").strip()
            
            # Обработка специальных команд
            if selection.lower() in ["c", "cancel", "отмена"]:
                print("Operation cancelled by user")
                return
            if selection.lower() in ["b", "back", "назад"]:
                return
            if selection.lower() == "a":
                # Обрабатываем выбор "all" только для этого уровня
                logger.info(f"User chose 'all' for column '{column}' at level {level} (this level only)")
                for category in categories:
                    new_filters = current_filters.copy()
                    new_filters[column] = category
                    generate_combinations(level + 1, new_filters)
                return
            
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
            
            # Обработка выбора
            for category in user_categories:
                new_filters = current_filters.copy()
                new_filters[column] = category
                generate_combinations(level + 1, new_filters)
            
            return
    
    # Начинаем генерацию комбинаций с первого уровня
    generate_combinations(0, {})
    
    # Возвращаем все уникальные комбинации
    unique_combinations = []
    seen = set()
    for filters in all_combinations:
        filter_tuple = tuple(sorted(filters.items()))
        if filter_tuple not in seen:
            seen.add(filter_tuple)
            unique_combinations.append(filters)
    
    return unique_combinations