import pandas as pd

def load_xlsx(file_path: str) -> pd.DataFrame:
    """
    Загружает XLSX-файл в pandas DataFrame.

    :param file_path: Путь к XLSX-файлу.
    :return: Загруженный DataFrame.
    :raises: FileNotFoundError, ValueError (если файл не найден или имеет неверный формат).
    """
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        return df
    except FileNotFoundError:
        raise FileNotFoundError(f"Файл не найден: {file_path}")
    except Exception as e:
        # pandas может выбросить разные ошибки при чтении (например, xlrd в старых версиях)
        # Лучше уточнить тип ошибки, но для простоты - общий Exception
        raise ValueError(f"Ошибка при чтении файла: {e}")

def save_xlsx(dataframe: pd.DataFrame, output_path: str) -> None:
    """
    Сохраняет pandas DataFrame в XLSX-файл.

    :param dataframe: DataFrame для сохранения.
    :param output_path: Путь для сохранения XLSX-файла.
    """
    dataframe.to_excel(output_path, index=False, engine='openpyxl')