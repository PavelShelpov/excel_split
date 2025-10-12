import os
import sys
from pathlib import Path

def get_script_dir() -> Path:
    """Определяем корректную директорию скрипта/EXE"""
    if getattr(sys, 'frozen', False):  # Если запущен как EXE
        return Path(sys.executable).parent  # Директория EXE-файла
    else:
        return Path(__file__).parent  # Директория скрипта

def should_ignore(path, base_path, ignore_patterns):
    rel_path = path.relative_to(base_path)
    
    # Игнорируем системные директории
    if any(part in ignore_patterns['dirs'] for part in rel_path.parts):
        return True
    
    # Игнорируем по расширению
    if path.suffix in ignore_patterns['extensions']:
        return True
    
    # Игнорируем специальные файлы
    if path.name in ignore_patterns['files']:
        return True
    
    # Игнорируем скрытые файлы/папки
    if any(part.startswith('.') for part in rel_path.parts):
        return True
    
    return False

def get_project_structure(start_path, output_file, ignore_patterns):
    with open(output_file, 'w', encoding='utf-8') as result:
        base_path = Path(start_path)
        
        # Записываем дерево директорий
        result.write("=== PROJECT STRUCTURE ===\n\n")
        for root, dirs, files in os.walk(start_path, topdown=True):
            # Фильтрация директорий
            dirs[:] = [d for d in dirs if not should_ignore(Path(root)/d, base_path, ignore_patterns)]
            
            # Фильтрация файлов
            files = [f for f in files if not should_ignore(Path(root)/f, base_path, ignore_patterns)]
            
            rel_path = Path(root).relative_to(base_path)
            level = len(rel_path.parts)
            
            if rel_path == Path('.'):
                result.write(f"./\n")
            else:
                indent = '│   ' * (level-1) + '├── '
                result.write(f"{indent}{rel_path.name}/\n")
            
            for file in files:
                file_indent = '│   ' * level + '├── '
                result.write(f"{file_indent}{file}\n")
        
        # Добавляем содержимое файлов
        result.write("\n\n=== FILE CONTENTS ===\n\n")
        for root, dirs, files in os.walk(start_path):
            dirs[:] = [d for d in dirs if not should_ignore(Path(root)/d, base_path, ignore_patterns)]
            files = [f for f in files if not should_ignore(Path(root)/f, base_path, ignore_patterns)]
            
            for file in files:
                file_path = Path(root) / file
                rel_path = file_path.relative_to(base_path)
                
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                except UnicodeDecodeError:
                    content = "BINARY FILE CONTENT (NOT SHOWN)"
                except Exception as e:
                    content = f"ERROR READING FILE: {str(e)}"
                
                result.write(f"\n─── FILE: {rel_path} ───\n\n")
                result.write(content + '\n')

if __name__ == "__main__":
    # Используем новую функцию определения директории
    script_dir = get_script_dir()
    output_filename = "project_dump.txt"
    
    ignore_patterns = {
        'dirs': {'__pycache__', '.git', '.idea', 'venv', '.venv'},
        'extensions': {'.pyc', '.pyo', '.pyd', '.so', '.dll', '.exe'},
        'files': {output_filename, Path(sys.argv[0]).name}  # Игнорируем EXE-файл
    }
    
    # Явно указываем путь для сохранения
    output_path = script_dir / output_filename
    
    # Удаляем предыдущий файл
    if output_path.exists():
        try:
            output_path.unlink()
        except Exception as e:
            print(f"Error deleting old file: {e}")
            sys.exit(1)
    
    # Запускаем генерацию
    try:
        get_project_structure(script_dir, output_path, ignore_patterns)
        print(f"Success! File saved to: {output_path}")
    except Exception as e:
        print(f"Critical error: {str(e)}")
        sys.exit(1)
    
    # Дополнительная проверка для EXE
    if getattr(sys, 'frozen', False):
        input("Press Enter to exit...")  # Чтобы окно не закрывалось сразу