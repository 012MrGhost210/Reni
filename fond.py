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
    
    def debug_search(self, sheet):
        """
        Отладочная функция: выводит ВСЕ ячейки, которые содержат слово "чистых" или "активов"
        """
        print("\n   🔍 ОТЛАДКА: Поиск ключевых слов в файле...")
        print("   " + "-"*60)
        
        found_count = 0
        for row_idx in range(sheet.nrows):
            for col_idx in range(sheet.ncols):
                cell_value = sheet.cell(row_idx, col_idx).value
                if cell_value and isinstance(cell_value, str):
                    cell_lower = cell_value.lower()
                    # Ищем разные варианты ключевых слов
                    if any(word in cell_lower for word in ['чистых', 'активов', 'итого', 'газпромбанк']):
                        found_count += 1
                        print(f"   Строка {row_idx}, столбец {col_idx} ({self.get_column_letter(col_idx)}):")
                        print(f"      Значение: '{cell_value}'")
                        
                        # Показываем соседние ячейки (5 столбцов вправо)
                        print(f"      Соседние ячейки справа:")
                        for offset in range(1, 6):
                            if col_idx + offset < sheet.ncols:
                                neighbor = sheet.cell(row_idx, col_idx + offset).value
                                if neighbor:
                                    print(f"         +{offset}: '{neighbor}'")
                        print()
        
        if found_count == 0:
            print("   ⚠️ Ни одного ключевого слова не найдено!")
            print("   Проверьте, что файл не пустой и содержит текст")
        
        print("   " + "-"*60)
        return found_count
    
    def get_column_letter(self, col_idx):
        """Преобразует индекс столбца в букву (0=A, 1=B, etc.)"""
        result = ""
        while col_idx >= 0:
            result = chr(65 + (col_idx % 26)) + result
            col_idx = col_idx // 26 - 1
        return result
    
    def find_value_by_keyword(self, sheet, search_text, target_col):
        """
        Поиск значения по ключевому слову
        """
        for row_idx in range(sheet.nrows):
            for col_idx in range(sheet.ncols):
                cell_value = sheet.cell(row_idx, col_idx).value
                if cell_value and isinstance(cell_value, str):
                    if search_text.lower() in cell_value.lower():
                        print(f"      ✅ Найдено ключевое слово в строке {row_idx}, столбец {col_idx}")
                        
                        # Берем значение из целевого столбца
                        if target_col < sheet.ncols:
                            value_cell = sheet.cell(row_idx, target_col)
                            print(f"      📍 Значение в столбце {target_col} ({self.get_column_letter(target_col)}): '{value_cell.value}'")
                            return self.parse_number(value_cell.value)
                        else:
                            print(f"      ⚠️ Столбец {target_col} не существует!")
                            return None
        return None
    
    def parse_number(self, value):
        """Преобразует значение в число"""
        if isinstance(value, (float, int)):
            return float(value)
        
        if isinstance(value, str):
            # Убираем пробелы, запятые, слово "руб"
            cleaned = value.replace(' ', '').replace(',', '.').replace('руб', '').replace('р.', '').strip()
            try:
                return float(cleaned)
            except:
                print(f"      ⚠️ Не удалось преобразовать в число: '{value}'")
                return None
        return None
    
    def process_file(self, file_path):
        """Обрабатывает один Excel файл"""
        print(f"\n📄 Обрабатываю: {file_path.name}")
        
        # Извлекаем дату из имени
        file_date = self.extract_date_from_filename(file_path.name)
        if file_date:
            print(f"   Дата из имени: {file_date}")
        
        found_value = None
        
        try:
            # Открываем .xls файл
            wb = xlrd.open_workbook(str(file_path), formatting_info=False)
            sheet = wb.sheet_by_index(0)
            
            print(f"   Размер листа: {sheet.nrows} строк x {sheet.ncols} столбцов")
            
            # 🔍 ОТЛАДКА: показываем все ключевые слова в файле
            self.debug_search(sheet)
            
            # ⚠️ ЗДЕСЬ НУЖНО УКАЗАТЬ КЛЮЧЕВОЕ СЛОВО И СТОЛБЕЦ
            search_text = "Итого стоимость чистых активов"  # ← ИЗМЕНИТЕ ЗДЕСЬ
            target_col = 14  # ← ИЗМЕНИТЕ ЗДЕСЬ (14 = O, 15 = P, 23 = X)
            
            print(f"\n   🔍 Ищу: '{search_text}' в любом столбце")
            print(f"   📍 Беру число из столбца: {self.get_column_letter(target_col)} (индекс {target_col})")
            
            # Ищем значение
            found_value = self.find_value_by_keyword(sheet, search_text, target_col)
            
            if found_value is not None:
                value_str = f"{found_value:,.2f}".replace(',', ' ')
                print(f"   ✅ Найдено: {value_str} руб.")
            else:
                print(f"   ⚠️ Ключевое слово не найдено")
                
        except Exception as e:
            print(f"   ❌ Ошибка: {e}")
            import traceback
            traceback.print_exc()
            found_value = None
        
        return {
            'Файл': file_path.name,
            'Дата': file_date if file_date else 'Не найдена',
            'Значение': found_value
        }
    
    def run(self):
        """Запускает обработку всех файлов"""
        print("="*80)
        print("ОТЛАДОЧНЫЙ ПАРСИНГ - ПОИСК КЛЮЧЕВЫХ СЛОВ")
        print("="*80)
        print(f"📂 Папка: {self.input_folder}")
        print("="*80)
        
        # Получаем все .xls файлы
        excel_files = list(self.input_folder.glob("*.xls"))
        excel_files.sort()
        
        print(f"\nНайдено файлов: {len(excel_files)}")
        
        if not excel_files:
            print("\n❌ Нет .xls файлов!")
            return
        
        # Обрабатываем КАЖДЫЙ файл
        for file_path in excel_files:
            self.process_file(file_path)
            print("\n" + "="*80)
            input("Нажмите Enter для продолжения к следующему файлу...")
        
        print("\n✅ Отладка завершена")

def main():
    # Путь к папке с файлами
    input_folder = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\Фонд СЧА"
    
    # Создаем парсер и запускаем
    parser = ExcelParser(input_folder, "")
    parser.run()

if __name__ == "__main__":
    main()
