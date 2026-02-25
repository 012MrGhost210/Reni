import os
import re
from pathlib import Path
import csv
from datetime import datetime

try:
    import xlrd
except ImportError:
    print("\n❌ Не установлена библиотека xlrd!")
    print("Установите командой: pip install xlrd")
    input("\nНажмите Enter для выхода...")
    exit()

class ExcelParser:
    def __init__(self, input_folder, output_file):
        self.input_folder = Path(input_folder)
        self.output_file = Path(output_file)
        self.results = []
        
    def extract_date_from_filename(self, filename):
        """Извлекает дату из имени файла"""
        match = re.search(r'(\d{2}\.\d{2}\.\d{4})', filename)
        return match.group(1) if match else None
    
    def debug_find_gazprom(self, sheet):
        """Отладочная функция - ищет все упоминания ГАЗПРОМ"""
        found_rows = []
        
        for row_idx in range(min(sheet.nrows, 50)):  # Проверим первые 50 строк
            row = sheet.row(row_idx)
            row_values = []
            
            for col_idx, cell in enumerate(row[:10]):  # Первые 10 колонок
                cell_value = cell.value
                if cell_value:
                    cell_str = str(cell_value)
                    row_values.append(f"{col_idx}:{cell_str[:30]}")
                    
                    if "ГАЗПРОМ" in cell_str.upper():
                        found_rows.append((row_idx + 1, col_idx + 1, cell_str))
            
            if row_values and row_idx < 20:  # Покажем первые 20 строк для отладки
                print(f"      Строка {row_idx + 1}: {' | '.join(row_values)}")
        
        return found_rows
    
    def find_gazprombank_value(self, sheet):
        """Поиск значения для ГАЗПРОМБАНКА"""
        search_text = "ГАЗПРОМБАНК"
        
        for row_idx in range(sheet.nrows):
            row = sheet.row(row_idx)
            for col_idx, cell in enumerate(row):
                cell_value = cell.value
                if cell_value and search_text in str(cell_value):
                    print(f"      ✅ Найден ГАЗПРОМБАНК в строке {row_idx + 1}, колонке {col_idx + 1}")
                    
                    # Проверим все значения в этой строке
                    print(f"      Все значения в строке {row_idx + 1}:")
                    for c in range(sheet.ncols):
                        val = sheet.cell(row_idx, c).value
                        if val and str(val).strip():
                            print(f"        Колонка {c + 1} ({chr(65 + c)}): {val}")
                    
                    # Ищем числовое значение справа
                    for offset in range(1, 10):
                        target_col = col_idx + offset
                        if target_col < sheet.ncols:
                            val = sheet.cell(row_idx, target_col).value
                            if isinstance(val, (float, int)):
                                print(f"      ✅ Найдено число в колонке {target_col + 1} ({chr(65 + target_col)}): {val}")
                                return val
                    
                    return None
        return None
    
    def process_file(self, file_path):
        """Обрабатывает один Excel файл"""
        print(f"\n📄 Обрабатываю: {file_path.name}")
        
        # Извлекаем дату из имени
        file_date = self.extract_date_from_filename(file_path.name)
        print(f"   Дата из имени: {file_date}")
        
        found_value = None
        
        try:
            # Открываем .xls файл
            wb = xlrd.open_workbook(str(file_path), formatting_info=False)
            sheet = wb.sheet_by_index(0)  # Берем первый лист
            
            print(f"   Размер листа: {sheet.nrows} строк x {sheet.ncols} колонок")
            
            # ОТЛАДКА: покажем структуру первых строк
            print(f"\n   🔍 ОТЛАДКА - первые 10 строк:")
            gazprom_mentions = self.debug_find_gazprom(sheet)
            
            if gazprom_mentions:
                print(f"\n   🔍 Найдены упоминания ГАЗПРОМ:")
                for row, col, text in gazprom_mentions:
                    print(f"      Строка {row}, колонка {col}: {text}")
                
                # Теперь ищем значение
                found_value = self.find_gazprombank_value(sheet)
            else:
                print(f"\n   ⚠️ ГАЗПРОМ не найден в первых 50 строках")
            
            if found_value is not None:
                value_str = f"{found_value:,.0f}".replace(',', ' ')
                print(f"\n   ✅ Найдено значение: {value_str} руб.")
            else:
                print(f"\n   ⚠️ Значение не найдено")
                
        except Exception as e:
            print(f"   ❌ Ошибка: {e}")
            found_value = None
        
        return {
            'Файл': file_path.name,
            'Дата': file_date if file_date else 'Не найдена',
            'Значение': found_value
        }
    
    def run(self):
        """Запускает обработку всех файлов"""
        print("="*80)
        print("ПАРСИНГ EXCEL ФАЙЛОВ (ОТЛАДОЧНЫЙ РЕЖИМ)")
        print("="*80)
        print(f"📂 Папка с файлами: {self.input_folder}")
        print("="*80)
        
        # Получаем все .xls файлы
        excel_files = list(self.input_folder.glob("*.xls"))
        excel_files.sort()  # Сортируем по имени
        
        print(f"\nНайдено .xls файлов: {len(excel_files)}")
        
        if not excel_files:
            print("\n❌ Нет .xls файлов для обработки!")
            return
        
        # Обрабатываем только первые 5 файлов для отладки
        files_to_process = excel_files[:10]
        print(f"\nОбрабатываем первые {len(files_to_process)} файлов для отладки")
        
        for file_path in files_to_process:
            result = self.process_file(file_path)
            self.results.append(result)
            
            input("\nНажмите Enter для продолжения...")
        
        print("\n" + "="*80)
        print("ОТЛАДКА ЗАВЕРШЕНА")
        print("="*80)

def main():
    # Путь к папке с файлами
    input_folder = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\Фонд СЧА"
    
    # Создаем парсер и запускаем
    parser = ExcelParser(input_folder, None)
    parser.run()
    
    print("\n" + "="*80)
    input("Нажмите Enter для выхода...")

if __name__ == "__main__":
    main()
