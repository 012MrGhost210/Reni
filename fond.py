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
    
    def clean_number(self, value):
        """
        Очищает число от всего лишнего и преобразует в float
        """
        if value is None:
            return None
        
        # Если уже число
        if isinstance(value, (int, float)):
            return float(value)
        
        # Преобразуем в строку
        value_str = str(value).strip()
        
        # Удаляем слово "руб" и всё после него
        if 'руб' in value_str.lower():
            value_str = value_str.lower().split('руб')[0].strip()
        
        # Удаляем все пробелы (включая неразрывные)
        value_str = re.sub(r'\s+', '', value_str)
        
        # Заменяем запятую на точку (десятичный разделитель)
        value_str = value_str.replace(',', '.')
        
        # Удаляем всё кроме цифр и точки
        value_str = re.sub(r'[^\d.]', '', value_str)
        
        # Если несколько точек, оставляем только последнюю
        if value_str.count('.') > 1:
            parts = value_str.split('.')
            value_str = ''.join(parts[:-1]) + '.' + parts[-1]
        
        try:
            return float(value_str)
        except ValueError:
            return None
    
    def find_net_asset_value(self, sheet):
        """
        Ищет 'Итого стоимость чистых активов' в столбце A
        и возвращает число из столбца P (индекс 15)
        """
        search_text = "Итого стоимость чистых активов"
        target_col = 15  # P = 15
        
        for row_idx in range(sheet.nrows):
            # Проверяем столбец A (индекс 0)
            if sheet.ncols > 0:
                cell_value = sheet.cell(row_idx, 0).value
                if cell_value and isinstance(cell_value, str):
                    if search_text.lower() in cell_value.lower():
                        print(f"      ✅ Найдено в строке {row_idx}")
                        
                        # Берем значение из столбца P
                        if target_col < sheet.ncols:
                            raw_value = sheet.cell(row_idx, target_col).value
                            print(f"      📍 Сырое значение в P: '{raw_value}'")
                            
                            # Очищаем и преобразуем
                            number = self.clean_number(raw_value)
                            if number is not None:
                                print(f"      ✅ Преобразовано в число: {number}")
                                return number
                            else:
                                print(f"      ❌ Не удалось преобразовать: '{raw_value}'")
                        else:
                            print(f"      ❌ Столбца P нет в файле")
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
            
            # Ищем значение
            found_value = self.find_net_asset_value(sheet)
            
            if found_value is not None:
                value_str = f"{found_value:,.2f}".replace(',', ' ')
                print(f"   ✅ Найдено: {value_str} руб.")
            else:
                print(f"   ❌ Значение не найдено")
                
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
        print("ПАРСИНГ - СТОИМОСТЬ ЧИСТЫХ АКТИВОВ")
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
            print(f"\n💰 Общая сумма: {total_sum:,.2f} руб.".replace(',', ' '))
            
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
