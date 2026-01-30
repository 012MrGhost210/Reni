import os
import pandas as pd
from pathlib import Path, PureWindowsPath
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.formatting.rule import FormulaRule
import urllib.parse

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

# Делать ли гиперссылки на файлы
CREATE_HYPERLINKS = True

# Открывать ли Excel файл после создания
OPEN_EXCEL_AFTER_CREATION = True

# ==================== КОНЕЦ НАСТРОЕК ====================

def format_excel_file(worksheet, total_rows):
    """
    Форматирует Excel файл: настраивает ширину столбцов, стили, гиперссылки
    """
    # Устанавливаем ширину столбцов
    column_widths = {
        'A': 40,  # Имя файла
        'B': 20,  # Тип файла
        'C': 25,  # Дата изменения
        'D': 100  # Полный путь (будет скрыт, так как есть гиперссылки)
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
    
    # Настраиваем стиль для гиперссылок
    hyperlink_font = Font(color="0563C1", underline="single")
    
    # Применяем стиль гиперссылок ко всем ячейкам с путями (колонка D)
    for row in range(2, total_rows + 2):  # +2 потому что заголовок в строке 1
        path_cell = worksheet.cell(row=row, column=4)  # Колонка D
        
        if CREATE_HYPERLINKS and path_cell.hyperlink:
            path_cell.font = hyperlink_font
    
    # Делаем автофильтр для заголовков
    worksheet.auto_filter.ref = worksheet.dimensions
    
    # Замораживаем первую строку
    worksheet.freeze_panes = "A2"
    
    # Добавляем условное форматирование для дат
    date_column_letter = 'C'
    date_range = f"{date_column_letter}2:{date_column_letter}{total_rows + 1}"
    
    # Форматирование для сегодняшних файлов
    today_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    today_font = Font(color="006100")
    
    # Ищем файлы, измененные сегодня
    today_rule = FormulaRule(
        formula=[f'AND(${date_column_letter}2>=TODAY(), ${date_column_letter}2<TODAY()+1)'],
        fill=today_fill,
        font=today_font
    )
    worksheet.conditional_formatting.add(date_range, today_rule)
    
    # Форматирование для старых файлов (старше 30 дней)
    old_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    old_font = Font(color="9C0006")
    
    old_rule = FormulaRule(
        formula=[f'${date_column_letter}2<TODAY()-30'],
        fill=old_fill,
        font=old_font
    )
    worksheet.conditional_formatting.add(date_range, old_rule)

def create_file_hyperlink(file_path):
    """
    Создает корректную гиперссылку для файла
    """
    try:
        # Проверяем существование файла
        if not os.path.exists(file_path):
            return None
        
        # Преобразуем путь в формат file://
        # Для Windows путей нужно специальное преобразование
        abs_path = os.path.abspath(file_path)
        
        # Создаем гиперссылку
        # В Excel для Windows лучше использовать обратные слеши
        hyperlink_path = abs_path.replace('/', '\\')
        
        return hyperlink_path
    except:
        return None

def create_excel_report(files_data, output_path):
    """
    Создает Excel файл со списком всех файлов с гиперссылками
    """
    try:
        # Создаем новую рабочую книгу
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Все файлы"
        
        # Добавляем заголовки
        headers = ["Имя файла", "Тип файла", "Дата изменения", "Полный путь"]
        for col, header in enumerate(headers, 1):
            worksheet.cell(row=1, column=col, value=header)
        
        # Заполняем данные
        for row_idx, file_info in enumerate(files_data, 2):  # Начинаем со 2 строки
            filename, file_type, mod_date, full_path = file_info
            
            # Имя файла
            worksheet.cell(row=row_idx, column=1, value=filename)
            
            # Тип файла
            worksheet.cell(row=row_idx, column=2, value=file_type)
            
            # Дата изменения (как дата Excel)
            try:
                date_obj = datetime.strptime(mod_date, '%Y-%m-%d %H:%M:%S')
                worksheet.cell(row=row_idx, column=3, value=date_obj)
                worksheet.cell(row=row_idx, column=3).number_format = 'YYYY-MM-DD HH:MM:SS'
            except:
                worksheet.cell(row=row_idx, column=3, value=mod_date)
            
            # Полный путь с гиперссылкой
            path_cell = worksheet.cell(row=row_idx, column=4, value=full_path)
            
            # Создаем гиперссылку если нужно
            if CREATE_HYPERLINKS:
                hyperlink = create_file_hyperlink(full_path)
                if hyperlink:
                    # Создаем гиперссылку в колонке D
                    path_cell.hyperlink = hyperlink
                    path_cell.value = full_path
                    
                    # Также делаем гиперссылку в имени файла (колонка A)
                    name_cell = worksheet.cell(row=row_idx, column=1)
                    name_cell.hyperlink = hyperlink
        
        # Сортируем по дате (самые новые сверху)
        worksheet.auto_filter.ref = worksheet.dimensions
        
        # Применяем форматирование
        total_rows = len(files_data)
        format_excel_file(worksheet, total_rows)
        
        # Добавляем итоговую строку
        total_row = total_rows + 3
        total_cell = worksheet.cell(row=total_row, column=1, value=f"Всего файлов: {total_rows}")
        total_cell.font = Font(bold=True, color="FF0000", size=12)
        
        # Добавляем инструкцию
        if CREATE_HYPERLINKS:
            instruction_row = total_rows + 4
            worksheet.cell(row=instruction_row, column=1, 
                          value="💡 Щелкните по имени файла или пути, чтобы открыть файл")
            worksheet.cell(row=instruction_row, column=1).font = Font(color="00B050", italic=True)
        
        # Сохраняем файл
        workbook.save(output_path)
        
        return True
        
    except Exception as e:
        print(f"❌ Ошибка при создании Excel отчета: {e}")
        import traceback
        traceback.print_exc()
        return False

def open_excel_file(file_path):
    """
    Открывает Excel файл после создания
    """
    try:
        os.startfile(file_path)
        print(f"📂 Открываю Excel файл...")
        return True
    except Exception as e:
        print(f"⚠️ Не удалось открыть Excel файл: {e}")
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
    
    print("=" * 70)
    print("📁 АНАЛИЗ ФАЙЛОВ В ДИРЕКТОРИИ С ГИПЕРССЫЛКАМИ")
    print("=" * 70)
    print(f"📂 Анализируемая папка: {SOURCE_DIRECTORY}")
    print(f"💾 Отчет будет сохранен: {excel_path}")
    print(f"🔗 Гиперссылки: {'ВКЛЮЧЕНЫ' if CREATE_HYPERLINKS else 'ВЫКЛЮЧЕНЫ'}")
    print("-" * 70)
    
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
                    file_type = file_extension.lstrip('.').upper()
                else:
                    file_type = "БЕЗ РАСШИРЕНИЯ"
                
                # Дата последнего изменения
                try:
                    mod_time = os.path.getmtime(file_path)
                    mod_date = datetime.fromtimestamp(mod_time).strftime('%Y-%m-%d %H:%M:%S')
                except:
                    mod_date = 'НЕДОСТУПНО'
                
                # Полный путь к файлу
                full_path = str(file_path.resolve())
                
                # Добавляем информацию в список
                files_data.append([
                    filename,
                    file_type,
                    mod_date,
                    full_path
                ])
                
                processed_files += 1
                
                if SHOW_DETAILS and processed_files % 100 == 0:
                    print(f"📊 Обработано файлов: {processed_files}...")
                    
            except Exception as e:
                print(f"⚠️ Ошибка при обработке файла {file_path}: {e}")
                continue
    
    # Создаем Excel отчет
    if files_data:
        print("-" * 70)
        print("📈 СОЗДАНИЕ ОТЧЕТА С ГИПЕРССЫЛКАМИ...")
        
        success = create_excel_report(files_data, excel_path)
        
        print("-" * 70)
        print("🎯 РЕЗУЛЬТАТЫ АНАЛИЗА:")
        print(f"   📄 Всего файлов в директории: {total_files}")
        print(f"   ✅ Успешно обработано: {processed_files}")
        print(f"   ⚠️  Ошибок обработки: {total_files - processed_files}")
        
        if success:
            print(f"   💾 Excel отчет успешно создан: {excel_path}")
            print(f"   📋 Записей в отчете: {len(files_data)}")
            
            if CREATE_HYPERLINKS:
                print(f"   🔗 Гиперссылки добавлены к именам файлов и путям")
                print(f"   💡 В Excel: щелкните по имени файла или пути для открытия")
            
            # Выводим краткую статистику по типам файлов
            extensions = {}
            for file_info in files_data:
                ext = file_info[1]  # Тип файла из второго столбца
                extensions[ext] = extensions.get(ext, 0) + 1
            
            print("\n   📊 СТАТИСТИКА ПО ТИПАМ ФАЙЛОВ:")
            top_extensions = sorted(extensions.items(), key=lambda x: x[1], reverse=True)[:10]
            for ext, count in top_extensions:
                percentage = (count / len(files_data)) * 100
                print(f"      {ext:<15} : {count:>5} файлов ({percentage:.1f}%)")
            
            # Открываем Excel файл если нужно
            if OPEN_EXCEL_AFTER_CREATION:
                open_excel_file(excel_path)
                
        else:
            print("   ❌ Не удалось создать Excel отчет")
    else:
        print("ℹ️  В указанной директории не найдено файлов.")
    
    print("=" * 70)

if __name__ == "__main__":
    analyze_directory_files()
