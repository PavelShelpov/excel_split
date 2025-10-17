from .analysis import analyze_column
import logging
logger = logging.getLogger('excel_splitter')

def get_all_combinations(source, valid_sheets, hierarchy_columns, filters=None, level=0):
    """Возвращает все возможные комбинации фильтров, включая частичные уровни."""
    if filters is None:
        filters = {}
    
    # Если достигли конца иерархии, возвращаем текущие фильтры
    if level >= len(hierarchy_columns):
        return [filters.copy()]
    
    column = hierarchy_columns[level]
    categories = analyze_column(source, valid_sheets, column, filters)
    
    # Если нет категорий, возвращаем пустой список
    if not categories:
        return []
    
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
        """Рекурсивная функция генерации комбинаций"""
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
            print("\nCurrent filters:")
            for col, value in current_filters.items():
                print(f"  - {col}: {value}")
        
        # Выводим информацию о текущем уровне
        print(f"\n--- Level {level + 1}/{len(hierarchy_columns)} ---")
        print(f"Column for filtering: '{column}'")
        
        # Выводим доступные категории с номерами
        print(f"\nAvailable categories for column '{column}':")
        for i, cat in enumerate(categories, 1):
            print(f"  {i}. {cat}")
        
        print("  a. All (for this level and all subsequent levels)")
        print("  s. Select specific categories (for this level only)")
        print("  b. Back")
        print("  c. Cancel")
        
        # Запрашиваем выбор
        while True:
            selection = input(f"Enter selection for '{column}' (a/s/b/c): ").strip().lower()
            
            # Обработка специальных команд
            if selection in ["c", "cancel"]:
                print("Operation cancelled by user")
                return
            if selection in ["b", "back"]:
                return
            if selection in ["a", "all"]:
                # Обрабатываем выбор "all" для всех оставшихся уровней
                logger.info(f"User chose 'all' for column '{column}' at level {level} and all subsequent levels")
                
                # Если это последний уровень, просто добавляем все категории
                if level == len(hierarchy_columns) - 1:
                    for category in categories:
                        new_filters = current_filters.copy()
                        new_filters[column] = category
                        all_combinations.append(new_filters)
                else:
                    # Для промежуточных уровней генерируем все возможные комбинации
                    all_combinations_recursive = get_all_combinations(
                        source, valid_sheets, hierarchy_columns, current_filters, level
                    )
                    for combo in all_combinations_recursive:
                        all_combinations.append(combo)
                return
            
            if selection in ["s", "select"]:
                print("\nYou can select specific categories by numbers or enter 'all' for this level only.")
                print("Enter categories (comma-separated numbers or 'all' for this level):")
                
                while True:
                    category_selection = input(f"Enter categories for '{column}': ").strip().lower()
                    
                    if category_selection in ["c", "cancel"]:
                        print("Operation cancelled by user")
                        return
                    if category_selection in ["b", "back"]:
                        return
                    
                    # Обработка "all" только для текущего уровня
                    if category_selection == "all":
                        logger.info(f"User chose 'all' for column '{column}' at level {level} (this level only)")
                        for category in categories:
                            new_filters = current_filters.copy()
                            new_filters[column] = category
                            generate_combinations(level + 1, new_filters)
                        return
                    
                    # Обработка номеров
                    user_categories = []
                    invalid_inputs = []
                    for item in category_selection.split(","):
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
            
            print("Please enter 'a' for all, 's' for select, 'b' for back, or 'c' for cancel")
    
    # Начинаем генерацию комбинаций с первого уровня
    generate_combinations(0, {})
    
    # Возвращаем уникальные комбинации
    unique_combinations = []
    seen = set()
    for filters in all_combinations:
        filter_tuple = tuple(sorted(filters.items()))
        if filter_tuple not in seen:
            seen.add(filter_tuple)
            unique_combinations.append(filters)
    
    return unique_combinations