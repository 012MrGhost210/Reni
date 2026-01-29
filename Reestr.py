import os
import shutil
from pathlib import Path
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

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

# Настройки Excel отчета
CREATE_EXCEL_REPORT = True  # Создавать ли Excel файл со списком всех файлов
EXCEL_FILENAME = "file_list.xlsx"  # Название Excel файла
EXCEL_COLUMNS = ["Имя файла", "Тип файла", "Дата изменения", "Полный путь"]  # Заголовки столбцов

# ==================== КОНЕЦ НАСТРОЕК ====================

def create_excel_report(all_files_list, target_excel_path):
    """
    Создает Excel файл со списком всех файлов
    """
    try:
        # Создаем DataFrame
        df = pd.DataFrame(all_files_list, columns=EXCEL_COLUMNS)
        
        # Сохраняем в Excel с помощью openpyxl для форматирования
        with pd.ExcelWriter(target_excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Список файлов')
            
            # Получаем рабочую книгу и лист для форматирования
            workbook = writer.book
            worksheet = writer.sheets['Список файлов']
            
            # Устанавливаем ширину столбцов
            column_widths = {
                'A': 40,  # Имя файла
                'B': 15,  # Тип файла
                'C': 20,  # Дата изменения
                'D': 80   # Полный путь
            }
            
            for col, width in column_widths.items():
                worksheet.column_dimensions[col].width = width
            
            # Форматируем заголовки
            header_font = Font(bold=True, color="FFFFFF")
            header_fill = "4472C4"  # Синий цвет
            
            for cell in worksheet[1]:
                cell.font = header_font
                cell.fill = pd.styles.PatternFill(start_color=header_fill, 
                                                  end_color=header_font, 
                                                  fill_type="solid")
                cell.alignment = Alignment(horizontal='center')
        
        print(f"📊 Excel отчет создан: {target_excel_path}")
        return True
        
    except Exception as e:
        print(f"❌ Ошибка при создании Excel отчета: {e}")
        return False

def find_and_copy_files():
    """
    Находит файлы по заданным критериям и копирует их в целевую папку
    Создает Excel отчет со списком всех файлов
    """
    
    if CREATE_TARGET_DIR:
        Path(TARGET_DIRECTORY).mkdir(parents=True, exist_ok=True)
    
    found_files = 0
    copied_files = 0
    errors = 0
    
    # Список для хранения информации о всех файлах
    all_files_data = []
    
    print("=" * 60)
    print("ПОИСК, КОПИРОВАНИЕ ФАЙЛОВ И СОЗДАНИЕ ОТЧЕТА")
    print("=" * 60)
    print(f"Ключевое слово: '{FILE_NAME_KEYWORD}'")
    print(f"Типы файлов: {', '.join(FILE_EXTENSIONS)}")
    print(f"Чувствительность к регистру: {'Да' if CASE_SENSITIVE else 'Нет'}")
    print(f"Сохранять структуру папок: {'Да' if PRESERVE_FOLDER_STRUCTURE else 'Нет'}")
    print(f"Создать Excel отчет: {'Да' if CREATE_EXCEL_REPORT else 'Нет'}")
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
            
            # Получаем информацию о файле
            filename = file_path.name
            file_extension = file_path.suffix.lower().lstrip('.')
            if file_extension == '':
                file_extension = 'без расширения'
            
            # Получаем дату изменения файла
            try:
                mod_time = os.path.getmtime(file_path)
                mod_date = datetime.fromtimestamp(mod_time).strftime('%Y-%m-%d %H:%M:%S')
            except:
                mod_date = 'Недоступно'
            
            # Добавляем информацию о ВСЕХ файлах в список
            all_files_data.append([
                filename,  # Имя файла
                file_extension,  # Тип файла (расширение)
                mod_date,  # Дата изменения
                str(file_path)  # Полный путь
            ])
            
            # Проверяем расширение файла
            if not FILE_EXTENSIONS or file_path.suffix.lower() in target_extensions:
                # Подготавливаем строки для сравнения
                search_name = FILE_NAME_KEYWORD if CASE_SENSITIVE else FILE_NAME_KEYWORD.lower()
                current_name = file_path.stem if CASE_SENSITIVE else file_path.stem.lower()
                
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
    
    # Создаем Excel отчет если нужно
    excel_report_created = False
    if CREATE_EXCEL_REPORT and all_files_data:
        excel_path = Path(TARGET_DIRECTORY) / EXCEL_FILENAME
        excel_report_created = create_excel_report(all_files_data, excel_path)
    
    # Выводим итоговую статистику
    print("-" * 60)
    print("📊 РЕЗУЛЬТАТЫ:")
    print(f"   Всего найдено файлов в директории: {len(all_files_data)}")
    print(f"   Файлов по критериям поиска: {found_files}")
    print(f"   Успешно скопировано: {copied_files}")
    print(f"   Ошибок: {errors}")
    
    if CREATE_EXCEL_REPORT:
        print(f"   Excel отчет создан: {'Да' if excel_report_created else 'Нет'}")
        if excel_report_created:
            print(f"   📋 Записано строк в отчет: {len(all_files_data)}")
    
    if found_files == 0:
        print("   ℹ️  Файлы, соответствующие критериям поиска, не найдены.")
    else:
        print(f"   📂 Файлы скопированы в: {TARGET_DIRECTORY}")
        if PRESERVE_FOLDER_STRUCTURE:
            print("   📁 Структура папок сохранена")
    
    print("=" * 60)

if __name__ == "__main__":
    find_and_copy_files()
