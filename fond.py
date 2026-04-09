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
    
    def parse_number_from_string(self, value_str):
        """
        Преобразует строку типа '3 123 345 руб.' или '3123345' в число
        """
        if value_str is None:
            return None
            
        if isinstance(value_str, (int, float)):
            return float(value_str)
        
        if not isinstance(value_str, str):
            return None
        
        # Удаляем 'руб.', 'руб', 'р.' и пробелы-разделители тысяч
        cleaned = value_str.replace('руб', '').replace('р.', '').strip()
        cleaned = cleaned.replace(' ', '')  # убираем пробелы тысяч
        cleaned = cleaned.replace(',', '.')  # заменяем запятую на точку (если есть)
        
        try:
            return float(cleaned)
        except ValueError:
            return None
    
    def find_net_asset_value(self, sheet):
        """
        Ищет в столбце A (индекс 0) строку с 'Итого стоимость чистых активов'
        и возвращает число из столбца P (индекс 15) в той же строке
        """
        search_phrase = "итого стоимость чистых активов"
        target_col = 15  # P = 15 (A=0, B=1, ..., P=15)
        
        for row_idx in range(sheet.nrows):
            # Проверяем ячейку в столбце A (индекс 0)
            if sheet.ncols > 0:
                cell_value = str(sheet.cell(row_idx, 0).value).lower().strip()
                
                # Проверяем, содержит ли ячейка искомую фразу
                if search_phrase in cell_value:
                    # Берем значение из столбца P (индекс 15)
                    if target_col < sheet.ncols:
                        value_cell = sheet.cell(row_idx, target_col)
                        parsed_value = self.parse_number_from_string(value_cell.value)
                        
                        if parsed_value is not None:
                            return parsed_value
                        else:
                            # Если не удалось распарсить, выводим отладочную информацию
                            print(f"      Отладка: найдено в строке {row_idx}, значение в P: '{value_cell.value}'")
        
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
            
            # Ищем значение "Итого стоимость чистых активов" в столбце A, число в столбце P
            found_value = self.find_net_asset_value(sheet)
            
            if found_value is not None:
                value_str = f"{found_value:,.2f}".replace(',', ' ')
                print(f"   ✅ Найдено: {value_str} руб.")
            else:
                print(f"   ⚠️ Фраза 'Итого стоимость чистых активов' не найдена")
                
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
        print("ПАРСИНГ EXCEL ФАЙЛОВ - ПОИСК 'ИТОГО СТОИМОСТЬ ЧИСТЫХ АКТИВОВ'")
        print("="*80)
        print(f"📂 Папка: {self.input_folder}")
        print(f"📄 Результат: {self.output_file}")
        print("="*80)
        
        # Получаем все .xls файлы
        excel_files = list(self.input_folder.glob("*.xls"))
        excel_files.sort()
        
        print(f"\nНайдено файлов: {len(excel_files)}")
        
        if not excel_files:
            print("\n❌ Нет .xls файлов!")
            return
        
        print("\n" + "-"*80)
        
        # Обрабатываем каждый файл
        for file_path in excel_files:
            result = self.process_file(file_path)
            self.results.append(result)
        
        # Сохраняем результаты
        self.save_results()
        self.print_summary()
    
    def save_results(self):
        """Сохраняет результаты в CSV"""
        try:
            with open(self.output_file, 'w', encoding='utf-8-sig', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(['Дата', 'Стоимость чистых активов (руб.)', 'Файл'])
                
                # Сортируем по дате
                sorted_results = sorted(self.results, 
                                      key=lambda x: x['Дата'] if x['Дата'] != 'Не найдена' else '')
                
                total_sum = 0
                for row in sorted_results:
                    if row['Значение'] is not None:
                        total_sum += row['Значение']
                        writer.writerow([
                            row['Дата'],
                            f"{row['Значение']:.2f}".replace('.', ','),
                            row['Файл']
                        ])
                    else:
                        writer.writerow([row['Дата'], 'НЕ НАЙДЕНО', row['Файл']])
                
                # Добавляем итоговую строку
                writer.writerow([])
                writer.writerow(['ИТОГО:', f"{total_sum:.2f}".replace('.', ','), ''])
                    
            print(f"\n✅ Результаты сохранены в: {self.output_file}")
            
        except Exception as e:
            print(f"\n❌ Ошибка при сохранении: {e}")
    
    def print_summary(self):
        """Выводит статистику"""
        print("\n" + "="*80)
        print("📊 ИТОГИ:")
        print("="*80)
        
        total = len(self.results)
        found = sum(1 for r in self.results if r['Значение'] is not None)
        total_sum = sum(r['Значение'] for r in self.results if r['Значение'] is not None)
        
        print(f"Всего файлов: {total}")
        print(f"Найдено значений: {found}")
        print(f"Не найдено: {total - found}")
        
        if found > 0:
            print(f"\n💰 Общая стоимость чистых активов: {total_sum:,.2f} руб.".replace(',', ' '))
            
            print("\n📋 Найденные значения:")
            print("-"*60)
            print(f"{'№':<4} {'Дата':<15} {'Значение':>20}") 
            print("-"*60)
            
            sorted_results = sorted([r for r in self.results if r['Значение'] is not None],
                                  key=lambda x: x['Дата'])
            
            for i, row in enumerate(sorted_results, 1):
                value_str = f"{row['Значение']:,.2f}".replace(',', ' ')
                print(f"{i:<4} {row['Дата']:<15} {value_str:>20}")

def main():
    # Путь к папке с файлами
    input_folder = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\Фонд СЧА"
    
    # Файл с результатами
    output_file = Path(input_folder) / f"!_РЕЗУЛЬТАТЫ_СТОИМОСТЬ_ЧИСТЫХ_АКТИВОВ.csv"
    
    # Создаем парсер и запускаем
    parser = ExcelParser(input_folder, output_file)
    parser.run()
    

if __name__ == "__main__":
    main()
