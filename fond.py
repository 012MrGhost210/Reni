import os
import re
from pathlib import Path

try:
    import xlrd
except ImportError:
    print("\n❌ Не установлена библиотека xlrd!")
    print("Установите командой: pip install xlrd")
    input("\nНажмите Enter для выхода...")
    exit()

class ExcelParser:
    def __init__(self, input_folder):
        self.input_folder = Path(input_folder)
        
    def extract_date_from_filename(self, filename):
        """Извлекает дату из имени файла"""
        match = re.search(r'(\d{2}\.\d{2}\.\d{4})', filename)
        return match.group(1) if match else None
    
    def debug_find_gazprom(self, sheet):
        """Отладочная функция - ищет все упоминания ГАЗПРОМ"""
        found_rows = []
        
        print(f"      Первые 15 строк файла:")
        print(f"      " + "-"*60)
        
        for row_idx in range(min(sheet.nrows, 30)):  # Проверим первые 30 строк
            row = sheet.row(row_idx)
            row_values = []
            
            # Проверяем первые 15 колонок
            for col_idx, cell in enumerate(row[:15]):
                cell_value = cell.value
                if cell_value:
                    cell_str = str(cell_value).strip()
                    if len(cell_str) > 50:
                        cell_str = cell_str[:50] + "..."
                    
                    # Если есть текст, показываем его
                    if cell_str:
                        row_values.append(f"[{col_idx+1}]{cell_str}")
                    
                    # Проверяем на ГАЗПРОМ
                    if "ГАЗПРОМ" in cell_str.upper():
                        found_rows.append((row_idx + 1, col_idx + 1, cell_str))
            
            if row_values:
                print(f"      Строка {row_idx + 1:2d}: {' | '.join(row_values)}")
        
        return found_rows
    
    def find_gazprombank_value(self, sheet):
        """Поиск значения для ГАЗПРОМБАНКА"""
        search_text = "ГАЗПРОМБАНК"
        
        for row_idx in range(sheet.nrows):
            row = sheet.row(row_idx)
            for col_idx, cell in enumerate(row):
                cell_value = cell.value
                if cell_value and search_text in str(cell_value):
                    print(f"\n      🔍 НАЙДЕН ГАЗПРОМБАНК в строке {row_idx + 1}, колонке {col_idx + 1}")
                    
                    # Покажем все значения в этой строке
                    print(f"      Все значения в строке {row_idx + 1}:")
                    for c in range(sheet.ncols):
                        val = sheet.cell(row_idx, c).value
                        if val is not None and str(val).strip():
                            col_letter = chr(65 + c) if c < 26 else f"Column{c+1}"
                            val_str = str(val).strip()
                            if len(val_str) > 50:
                                val_str = val_str[:50] + "..."
                            print(f"        {col_letter}{row_idx + 1}: {val_str}")
                    
                    # Ищем все числа в этой строке
                    numbers = []
                    for c in range(sheet.ncols):
                        val = sheet.cell(row_idx, c).value
                        if isinstance(val, (float, int)):
                            numbers.append((c+1, val))
                    
                    if numbers:
                        print(f"\n      Найдены числа в строке:")
                        for col, num in numbers:
                            print(f"        Колонка {col}: {num:,.0f}".replace(',', ' '))
                        return numbers[0][1] if numbers else None
                    
                    return None
        return None
    
    def process_file(self, file_path):
        """Обрабатывает один Excel файл"""
        print(f"\n{'='*60}")
        print(f"📄 Файл: {file_path.name}")
        print(f"{'='*60}")
        
        # Извлекаем дату из имени
        file_date = self.extract_date_from_filename(file_path.name)
        print(f"📅 Дата из имени: {file_date}")
        
        try:
            # Открываем .xls файл
            wb = xlrd.open_workbook(str(file_path), formatting_info=False)
            sheet = wb.sheet_by_index(0)  # Берем первый лист
            
            print(f"📊 Размер листа: {sheet.nrows} строк x {sheet.ncols} колонок")
            print(f"{'='*60}")
            
            # Покажем содержимое
            gazprom_mentions = self.debug_find_gazprom(sheet)
            
            if gazprom_mentions:
                print(f"\n🔍 Найдены упоминания ГАЗПРОМ:")
                for row, col, text in gazprom_mentions:
                    print(f"   📍 Строка {row}, колонка {col}: {text}")
                
                # Ищем значение
                value = self.find_gazprombank_value(sheet)
                if value:
                    print(f"\n✅ ЗНАЧЕНИЕ НАЙДЕНО: {value:,.0f} руб.".replace(',', ' '))
                else:
                    print(f"\n❌ Значение не найдено в строке с ГАЗПРОМБАНК")
            else:
                print(f"\n⚠️ ГАЗПРОМ не найден в первых 30 строках")
            
            print(f"\n{'-'*60}")
            
        except Exception as e:
            print(f"❌ Ошибка: {e}")
    
    def run(self):
        """Запускает обработку всех файлов"""
        print("="*80)
        print("🔍 ПАРСИНГ EXCEL ФАЙЛОВ (ОТЛАДОЧНЫЙ РЕЖИМ)")
        print("="*80)
        print(f"📂 Папка: {self.input_folder}")
        print("="*80)
        
        # Получаем все .xls файлы
        excel_files = list(self.input_folder.glob("*.xls"))
        excel_files.sort()
        
        print(f"\nНайдено .xls файлов: {len(excel_files)}")
        
        if not excel_files:
            print("\n❌ Нет .xls файлов!")
            return
        
        # Обрабатываем каждый файл по очереди
        for i, file_path in enumerate(excel_files, 1):
            print(f"\nФайл {i} из {len(excel_files)}")
            self.process_file(file_path)
            
            if i < len(excel_files):
                input("\nНажмите Enter для перехода к следующему файлу...")
        
        print("\n" + "="*80)
        print("✅ ОТЛАДКА ЗАВЕРШЕНА")
        print("="*80)

def main():
    # Путь к папке с файлами
    input_folder = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\Фонд СЧА"
    
    # Создаем парсер и запускаем
    parser = ExcelParser(input_folder)
    parser.run()
    
    print("\n" + "="*80)
    input("Нажмите Enter для выхода...")

if __name__ == "__main__":
    main()
