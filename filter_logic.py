import pandas as pd

def get_column_names(dataframe: pd.DataFrame) -> list[str]:
    """
    Возвращает список названий колонок из DataFrame.

    :param dataframe: Входной DataFrame.
    :return: Список названий колонок.
    """
    return list(dataframe.columns)

def get_unique_values(dataframe: pd.DataFrame, column_name: str) -> list:
    """
    Возвращает уникальные значения (категории) в указанной колонке.

    :param dataframe: Входной DataFrame.
    :param column_name: Название колонки.
    :return: Список уникальных значений. Включает NaN, если они есть.
    """
    # Используем dropna=False, чтобы NaN оставались в списке уникальных значений,
    # если они присутствуют. Это позволяет пользователю их выбрать.
    # Если NaN не нужны, можно использовать dataframe[column_name].dropna().unique()
    unique_series = dataframe[column_name].unique()
    # Конвертируем numpy array в list для удобства
    return unique_series.tolist()

def apply_filter(dataframe: pd.DataFrame, column_name: str, category_value) -> pd.DataFrame:
    """
    Применяет фильтр к DataFrame по указанной колонке и значению категории.

    :param dataframe: Входной DataFrame.
    :param column_name: Название колонки для фильтрации.
    :param category_value: Значение категории для фильтрации.
    :return: Новый, отфильтрованный DataFrame.
    """
    # Используем .copy() для создания независимой копии отфильтрованных данных
    filtered_df = dataframe[dataframe[column_name] == category_value].copy()
    return filtered_df