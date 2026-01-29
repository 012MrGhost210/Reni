import os
import pandas as pd
from pathlib import Path
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# ==================== НАСТРОЙКИ ====================
# Быстро меняйте параметры поиска здесь:

# Исходная папка для анализа
SOURCE_DIRECTORY = r'M:\Финансовый департамент\Treasury'  # ЗАМЕНИТЕ НА СВОЙ ПУТЬ

# Папка для сохранения Excel файла
OUTPUT_DIRECTORY = r'\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\Test'  # ЗАМЕНИТЕ НА СВОЙ ПУТЬ

# Название Excel файла
EXCEL_FILENAME = "анализ_файлов.xlsx"

# Показывать подробный процесс работы
SHOW_DETAILS = True

# Создать папку для отчета, если её нет
CREATE_OUTPUT_DIR = True

# ==================== КОНЕЦ НАСТРОЕК ====================

def format_excel_file(excel_path, worksheet):
    """
    Форматирует Excel файл: настраивает ширину столбцов, стили заголовков
    """
    # Устанавливаем ширину столбцов
    column_widths = {
        'A': 40,  # Имя файла
        'B': 20,  # Тип файла
        'C': 25,  # Дата изменения
        'D': 100  # Полный путь
    }
    
    for col, width in column_widths.items():
        worksheet.column_dimensions[col].width = width
    
    # Форматируем заголовки
    header_font = Font(bold=True, color="FFFFFF", size=12)
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    
    for col in range(1, 5):  # 4 колонки
        cell = worksheet.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Делаем автофильтр для заголовков
    worksheet.auto_filter.ref = worksheet.dimensions
    
    # Замораживаем первую строку
    worksheet.freeze_panes = "A2"

def create_excel_report(files_data, output_path):
    """
    Создает Excel файл со списком всех файлов
    """
    try:
        # Создаем DataFrame с нужными колонками
        df = pd.DataFrame(files_data, columns=[
            "Имя файла", 
            "Тип файла", 
            "Дата изменения", 
            "Полный путь"
        ])
        
        # Сортируем по имени файла
        df = df.sort_values("Имя файла")
        
        # Создаем Excel файл
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Все файлы')
            
            # Получаем рабочую книгу и лист для форматирования
            workbook = writer.book
            worksheet = writer.sheets['Все файлы']
            
            # Применяем форматирование
            format_excel_file(output_path, worksheet)
            
            # Добавляем итоговую строку
            total_row = len(files_data) + 3
            worksheet.cell(row=total_row, column=1, value=f"Всего файлов: {len(files_data)}")
            worksheet.cell(row=total_row, column=1).font = Font(bold=True, color="FF0000")
            
            # Сохраняем
            workbook.save(output_path)
        
        return True
        
    except Exception as e:
        print(f"❌ Ошибка при создании Excel отчета: {e}")
        return False

def analyze_directory_files():
    """
    Анализирует все файлы в директории и создает Excel отчет
    """
    if CREATE_OUTPUT_DIR:
        Path(OUTPUT_DIRECTORY).mkdir(parents=True, exist_ok=True)
    
    # Полный путь к Excel файлу
    excel_path = Path(OUTPUT_DIRECTORY) / EXCEL_FILENAME
    
    # Список для хранения информации о всех файлах
    files_data = []
    
    print("=" * 60)
    print("АНАЛИЗ ФАЙЛОВ В ДИРЕКТОРИИ")
    print("=" * 60)
    print(f"Анализируемая папка: {SOURCE_DIRECTORY}")
    print(f"Отчет будет сохранен: {excel_path}")
    print("-" * 60)
    
    if not os.path.exists(SOURCE_DIRECTORY):
        print(f"❌ Ошибка: Папка '{SOURCE_DIRECTORY}' не существует!")
        return
    
    # Счетчики
    total_files = 0
    processed_files = 0
    
    # Рекурсивно обходим все файлы в исходной директории
    for root, dirs, files in os.walk(SOURCE_DIRECTORY):
        for file in files:
            total_files += 1
            file_path = Path(root) / file
            
            try:
                # Получаем информацию о файле
                filename = file_path.name
                
                # Тип файла (расширение)
                file_extension = file_path.suffix.lower()
                if file_extension:
                    file_type = file_extension.lstrip('.')
                else:
                    file_type = "без расширения"
                
                # Дата последнего изменения
                try:
                    mod_time = os.path.getmtime(file_path)
                    mod_date = datetime.fromtimestamp(mod_time).strftime('%Y-%m-%d %H:%M:%S')
                except:
                    mod_date = 'Недоступно'
                
                # Полный путь к файлу
                full_path = str(file_path)
                
                # Добавляем информацию в список
                files_data.append([
                    filename,
                    file_type.upper() if file_type != "без расширения" else file_type,
                    mod_date,
                    full_path
                ])
                
                processed_files += 1
                
                if SHOW_DETAILS and processed_files % 100 == 0:
                    print(f"Обработано файлов: {processed_files}...")
                    
            except Exception as e:
                print(f"⚠️ Ошибка при обработке файла {file_path}: {e}")
                continue
    
    # Создаем Excel отчет
    if files_data:
        print("-" * 60)
        print("📊 СОЗДАНИЕ ОТЧЕТА...")
        
        success = create_excel_report(files_data, excel_path)
        
        print("-" * 60)
        print("📈 РЕЗУЛЬТАТЫ АНАЛИЗА:")
        print(f"   Всего файлов в директории: {total_files}")
        print(f"   Успешно обработано: {processed_files}")
        print(f"   Ошибок обработки: {total_files - processed_files}")
        
        if success:
            print(f"   ✅ Excel отчет успешно создан: {excel_path}")
            print(f"   📋 Записей в отчете: {len(files_data)}")
            
            # Выводим краткую статистику по типам файлов
            extensions = {}
            for file_info in files_data:
                ext = file_info[1]  # Тип файла из второго столбца
                extensions[ext] = extensions.get(ext, 0) + 1
            
            print("\n   📊 СТАТИСТИКА ПО ТИПАМ ФАЙЛОВ:")
            for ext, count in sorted(extensions.items(), key=lambda x: x[1], reverse=True)[:10]:
                print(f"      {ext}: {count} файлов")
                
        else:
            print("   ❌ Не удалось создать Excel отчет")
    else:
        print("ℹ️  В указанной директории не найдено файлов.")
    
    print("=" * 60)

if __name__ == "__main__":
    analyze_directory_files()
