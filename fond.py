import os
import re
from pathlib import Path
import csv
import xlrd

class ExcelParser:
    def __init__(self, input_folder, output_file):
        self.input_folder = Path(input_folder)
        self.output_file = Path(output_file)
        self.results = []
        
    def extract_date_from_filename(self, filename):
        """Извлекает дату из имени файла формата [2026_01_12]_29.12.2025_..."""
        match = re.search(r'(\d{2}\.\d{2}\.\d{4})', filename)
        return match.group(1) if match else None
    
    def find_value_by_text(self, sheet, search_text, offset_cols=8):
        """
        Ищет текст и возвращает значение со смещением
        offset_cols: смещение по столбцам (8 - на 8 колонок правее)
        """
        for row_idx in range(sheet.nrows):
            row = sheet.row(row_idx)
            for col_idx, cell in enumerate(row):
                # Проверяем значение ячейки
                cell_value = cell.value
                if cell_value and search_text in str(cell_value):
                    print(f"      Найден текст в строке {row_idx + 1}, колонке {col_idx + 1}")
                    
                    # Берем значение справа через offset_cols
                    target_col = col_idx + offset_cols
                    if target_col < sheet.ncols:
                        value_cell = sheet.cell(row_idx, target_col)
                        value = value_cell.value
                        
                        # Пробуем преобразовать в число если это возможно
                        if isinstance(value, (float, int)):
                            return value
                        elif isinstance(value, str):
                            # Пробуем извлечь число из строки
                            numbers = re.findall(r'-?\d+\.?\d*', value.replace(' ', ''))
                            if numbers:
                                return float(numbers[0])
                        return value
                    else:
                        print(f"      Выход за границы: целевая колонка {target_col + 1}, всего колонок {sheet.ncols}")
                        # Если вышли за границы, ищем последнее число в строке
                        for back_col in range(sheet.ncols - 1, col_idx, -1):
                            val = sheet.cell(row_idx, back_col).value
                            if isinstance(val, (float, int)):
                                return val
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
            
            # Ищем значение
            found_value = self.find_value_by_text(
                sheet, 
                "ГАЗПРОМБАНК", 
                offset_cols=8
            )
            
            if found_value is not None:
                print(f"   ✅ Найдено значение: {found_value}")
            else:
                print(f"   ⚠️ Значение не найдено")
                
        except Exception as e:
            print(f"   ❌ Ошибка: {e}")
            found_value = f"ОШИБКА: {str(e)[:50]}"
        
        return {
            'Файл': file_path.name,
            'Дата': file_date,
            'Значение': found_value
        }
    
    def run(self):
        """Запускает обработку всех файлов"""
        print("="*80)
        print("ПАРСИНГ EXCEL ФАЙЛОВ СЧА Фонд_ПДС")
        print("="*80)
        print(f"📂 Папка: {self.input_folder}")
        print(f"📄 Результат: {self.output_file}")
        print("="*80)
        
        # Получаем все .xls файлы
        excel_files = list(self.input_folder.glob("*.xls"))
        
        print(f"\nНайдено .xls файлов: {len(excel_files)}")
        
        if not excel_files:
            print("❌ Нет .xls файлов для обработки!")
            return
        
        # Обрабатываем каждый файл
        for file_path in excel_files:
            result = self.process_file(file_path)
            self.results.append(result)
        
        # Сохраняем результаты
        self.save_results()
        self.print_summary()
    
    def save_results(self):
        """Сохраняет результаты в CSV"""
        with open(self.output_file, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=['Файл', 'Дата', 'Значение'])
            writer.writeheader()
            
            for row in self.results:
                # Форматируем значение для CSV
                value = row['Значение']
                if isinstance(value, float):
                    row['Значение'] = f"{value:.2f}".replace('.', ',')
                writer.writerow(row)
                
        print(f"\n✅ Результаты сохранены в: {self.output_file}")
    
    def print_summary(self):
        """Выводит краткую статистику"""
        print("\n" + "="*80)
        print("📊 ИТОГИ:")
        print("="*80)
        
        # Считаем статистику
        total = len(self.results)
        found = sum(1 for r in self.results if r['Значение'] and not isinstance(r['Значение'], str) or (isinstance(r['Значение'], str) and not r['Значение'].startswith('ОШИБКА') and r['Значение'] != 'Не найдено'))
        errors = sum(1 for r in self.results if isinstance(r['Значение'], str) and r['Значение'].startswith('ОШИБКА'))
        
        print(f"Всего файлов: {total}")
        print(f"Найдено значений: {found}")
        print(f"Ошибок: {errors}")
        
        if found > 0:
            print("\n📋 Найденные значения:")
            print("-"*60)
            for i, row in enumerate(self.results, 1):
                if row['Значение'] and not isinstance(row['Значение'], str) or (isinstance(row['Значение'], str) and not row['Значение'].startswith('ОШИБКА')):
                    print(f"{i:2d}. {row['Дата']} -> {row['Значение']}")

def main():
    # Пути
    input_folder = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\Фонд СЧА"
    output_file = Path(input_folder) / "!_РЕЗУЛЬТАТЫ_ПАРСИНГА.csv"
    
    print("\n" + "="*80)
    print("УСТАНОВКА НЕОБХОДИМЫХ БИБЛИОТЕК")
    print("="*80)
    print("Выполните в командной строке:")
    print("pip install xlrd")
    print("\nИли нажмите Enter чтобы продолжить (если библиотека уже установлена)")
    input()
    
    # Создаем парсер и запускаем
    parser = ExcelParser(input_folder, output_file)
    parser.run()
    
    print("\n" + "="*80)
    input("Нажмите Enter для выхода...")

if __name__ == "__main__":
    main()
