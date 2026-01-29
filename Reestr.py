import os
import shutil
from pathlib import Path

# ==================== НАСТРОЙКИ ====================
# Быстро меняйте параметры поиска здесь:

# Исходная папка для поиска 
SOURCE_DIRECTORY = r'M:\Финансовый департамент\Treasury'  # ЗАМЕНИТЕ НА СВОЙ ПУТЬ

# Целевая папка для копирования 
TARGET_DIRECTORY = r'\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\Test'  # ЗАМЕНИТЕ НА СВОЙ ПУТЬ

# Ключевое слово в названии файла 
FILE_NAME_KEYWORD = "Чешенко"

# Типы файлов для поиска
FILE_EXTENSIONS = ["pdf", "docx", "xlsx"]  # Например: ["txt", "jpg", "png"]

# Чувствительность к регистру при поиске
CASE_SENSITIVE = False  # True - учитывает регистр, False - не учитывает

# Создать целевую папку, если её нет
CREATE_TARGET_DIR = True

# Показывать подробный процесс работы
SHOW_DETAILS = True

# Сохранять структуру папок при копировании
PRESERVE_FOLDER_STRUCTURE = False  # True - сохранит структуру папок, False - все файлы в одну папку

# ==================== КОНЕЦ НАСТРОЕК ====================

def find_and_copy_files():
    """
    Находит файлы по заданным критериям и копирует их в целевую папку
    """
    
    if CREATE_TARGET_DIR:
        Path(TARGET_DIRECTORY).mkdir(parents=True, exist_ok=True)
    
    found_files = 0
    copied_files = 0
    errors = 0
    
    print("=" * 60)
    print("ПОИСК И КОПИРОВАНИЕ ФАЙЛОВ")
    print("=" * 60)
    print(f"Ключевое слово: '{FILE_NAME_KEYWORD}'")
    print(f"Типы файлов: {', '.join(FILE_EXTENSIONS)}")
    print(f"Чувствительность к регистру: {'Да' if CASE_SENSITIVE else 'Нет'}")
    print(f"Сохранять структуру папок: {'Да' if PRESERVE_FOLDER_STRUCTURE else 'Нет'}")
    print(f"Ищем в: {SOURCE_DIRECTORY}")
    print(f"Копируем в: {TARGET_DIRECTORY}")
    print("-" * 60)
    

    if not os.path.exists(SOURCE_DIRECTORY):
        print(f"❌ Ошибка: Папка '{SOURCE_DIRECTORY}' не существует!")
        return
    
    # Подготавливаем расширения для сравнения
    target_extensions = [f".{ext.lower().lstrip('.')}" for ext in FILE_EXTENSIONS]
    
    # Рекурсивно обходим все файлы в исходной директории
    for root, dirs, files in os.walk(SOURCE_DIRECTORY):
        for file in files:
            file_path = Path(root) / file
            
            # Получаем имя файла без расширения и само расширение
            filename = file_path.stem
            extension = file_path.suffix.lower()
            
            # Проверяем расширение файла
            if not FILE_EXTENSIONS or extension in target_extensions:
                # Подготавливаем строки для сравнения
                search_name = FILE_NAME_KEYWORD if CASE_SENSITIVE else FILE_NAME_KEYWORD.lower()
                current_name = filename if CASE_SENSITIVE else filename.lower()
                
                # Проверяем соответствие имени файла
                if search_name in current_name:
                    found_files += 1
                    
                    if SHOW_DETAILS:
                        print(f"✅ Найден: {file_path}")
                    
                    try:
                        if PRESERVE_FOLDER_STRUCTURE:
                            # Сохраняем структуру папок
                            relative_path = Path(root).relative_to(SOURCE_DIRECTORY)
                            target_subdir = Path(TARGET_DIRECTORY) / relative_path
                            target_subdir.mkdir(parents=True, exist_ok=True)
                            target_file_path = target_subdir / file
                        else:
                            # Все файлы в одну папку
                            target_file_path = Path(TARGET_DIRECTORY) / file
                        
                        # Если файл с таким именем уже существует, добавляем номер
                        counter = 1
                        original_target = target_file_path
                        while target_file_path.exists():
                            name = original_target.stem
                            suffix = original_target.suffix
                            target_file_path = original_target.parent / f"{name}_{counter}{suffix}"
                            counter += 1
                        
                        # КОПИРУЕМ файл (вместо перемещения)
                        shutil.copy2(str(file_path), str(target_file_path))
                        copied_files += 1
                        
                        if SHOW_DETAILS:
                            print(f"   📁 Скопирован в: {target_file_path}")
                        
                    except Exception as e:
                        errors += 1
                        print(f"   ❌ Ошибка при копировании {file}: {e}")
    
    # Выводим итоговую статистику
    print("-" * 60)
    print("📊 РЕЗУЛЬТАТЫ:")
    print(f"   Найдено файлов: {found_files}")
    print(f"   Успешно скопировано: {copied_files}")
    print(f"   Ошибок: {errors}")
    
    if found_files == 0:
        print("   ℹ️  Файлы, соответствующие критериям поиска, не найдены.")
    else:
        print(f"   📂 Файлы скопированы в: {TARGET_DIRECTORY}")
        if PRESERVE_FOLDER_STRUCTURE:
            print("   📁 Структура папок сохранена")
    
    print("=" * 60)

if __name__ == "__main__":
    find_and_copy_files()
