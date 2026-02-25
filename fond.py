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
        """Извлекает дату из имени файла формата (2026_01_12)_29.12.2025_..."""
        # Сначала ищем дату в формате ДД.ММ.ГГГГ
        match = re.search(r'(\d{2}\.\d{2}\.\d{4})', filename)
        if match:
            return match.group(1)
        
        # Если не нашли, ищем дату папки в формате ГГГГ_ММ_ДД
        match = re.search(r'\((\d{4}_\d{2}_\d{2})\)', filename)
        if match:
            # Преобразуем 2026_01_12 в 12.01.2026
            date_parts = match.group(1).split('_')
            return f"{date_parts[2]}.{date_parts[1]}.{date_parts[0]}"
        
        return None
    
    def find_gazprombank_value(self, sheet):
        """
        Поиск значения для ГАЗПРОМБАНКА в колонке W (индекс 22)
        """
        search_text = "ГАЗПРОМБАНК"
        
        for row_idx in range(sheet.nrows):
            row = sheet.row(row_idx)
            for col_idx, cell in enumerate(row):
                cell_value = cell.value
                if cell_value and search_text in str(cell_value):
                    # Значение находится в колонке W (индекс 22, так как A=0, B=1, ..., W=22)
                    target_col = 22
                    
                    if target_col < sheet.ncols:
                        value_cell = sheet.cell(row_idx, target_col)
                        value = value_cell.value
                        
                        # Проверяем что это число
                        if isinstance(value, (float, int)):
                            return value
                        elif isinstance(value, str):
                            # Пробуем преобразовать строку в число
                            value = value.replace(' ', '').replace(',', '.')
                            try:
                                return float(value)
                            except:
                                return None
        return None
    
    def process_file(self, file_path):
        """Обрабатывает один Excel файл"""
        print(f"\n📄 Обрабатываю: {file_path.name}")
        
        # Извлекаем дату из имени
        file_date = self.extract_date_from_filename(file_path.name)
        if file_date:
            print(f"   Дата из имени: {file_date}")
        else:
            print(f"   ⚠️ Дата не найдена в имени")
        
        found_value = None
        
        try:
            # Открываем .xls файл
            wb = xlrd.open_workbook(str(file_path), formatting_info=False)
            sheet = wb.sheet_by_index(0)  # Берем первый лист
            
            print(f"   Размер листа: {sheet.nrows} строк x {sheet.ncols} колонок")
            
            # Ищем значение для ГАЗПРОМБАНКА
            found_value = self.find_gazprombank_value(sheet)
            
            if found_value is not None:
                # Форматируем число для красивого вывода
                value_str = f"{found_value:,.0f}".replace(',', ' ')
                print(f"   ✅ Найдено значение: {value_str} руб.")
            else:
                print(f"   ⚠️ Значение не найдено")
                
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
        print("ПАРСИНГ EXCEL ФАЙЛОВ СЧА Фонд_ПДС")
        print("="*80)
        print(f"📂 Папка с файлами: {self.input_folder}")
        print(f"📄 Результат будет сохранен в: {self.output_file}")
        print("="*80)
        
        # Получаем все .xls файлы
        excel_files = list(self.input_folder.glob("*.xls"))
        
        print(f"\nНайдено .xls файлов: {len(excel_files)}")
        
        if not excel_files:
            print("\n❌ Нет .xls файлов для обработки!")
            return
        
        print("\n" + "-"*80)
        print("НАЧАЛО ОБРАБОТКИ")
        print("-"*80)
        
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
                writer.writerow(['Дата', 'Значение (руб.)', 'Файл'])
                
                # Сортируем по дате
                sorted_results = sorted(self.results, 
                                      key=lambda x: x['Дата'] if x['Дата'] != 'Не найдена' else '')
                
                for row in sorted_results:
                    if row['Значение'] is not None:
                        # Форматируем число с разделителями
                        value_str = f"{row['Значение']:,.0f}".replace(',', ' ')
                        writer.writerow([
                            row['Дата'],
                            value_str,
                            row['Файл']
                        ])
                    else:
                        writer.writerow([row['Дата'], 'НЕ НАЙДЕНО', row['Файл']])
                    
            print(f"\n✅ Результаты сохранены в: {self.output_file}")
            
        except Exception as e:
            print(f"\n❌ Ошибка при сохранении результатов: {e}")
    
    def print_summary(self):
        """Выводит краткую статистику"""
        print("\n" + "="*80)
        print("📊 ИТОГИ ОБРАБОТКИ")
        print("="*80)
        
        # Считаем статистику
        total = len(self.results)
        found = sum(1 for r in self.results if r['Значение'] is not None)
        not_found = total - found
        
        print(f"Всего файлов: {total}")
        print(f"✅ Найдено значений: {found}")
        print(f"❌ Не найдено: {not_found}")
        
        if found > 0:
            print("\n📋 НАЙДЕННЫЕ ЗНАЧЕНИЯ:")
            print("-"*80)
            print(f"{'№':<4} {'Дата':<15} {'Значение (руб.)':>30} {'Файл':<30}")
            print("-"*80)
            
            # Сортируем по дате
            sorted_results = sorted([r for r in self.results if r['Значение'] is not None],
                                  key=lambda x: x['Дата'])
            
            for i, row in enumerate(sorted_results, 1):
                value_str = f"{row['Значение']:,.0f}".replace(',', ' ')
                short_name = row['Файl'][:30] + "..." if len(row['Файл']) > 33 else row['Файл']
                print(f"{i:<4} {row['Дата']:<15} {value_str:>30} {short_name:<30}")
            
            # Подсчет общей суммы
            total_sum = sum(r['Значение'] for r in self.results if r['Значение'] is not None)
            print("-"*80)
            print(f"{'ИТОГО:':<20} {total_sum:>30,.0f} руб.".replace(',', ' '))
        
        print("\n" + "="*80)

def main():
    # Путь к папке с файлами
    input_folder = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\Фонд СЧА"
    
    # Файл с результатами
    output_file = Path(input_folder) / f"!_РЕЗУЛЬТАТЫ_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    
    # Создаем парсер и запускаем
    parser = ExcelParser(input_folder, output_file)
    parser.run()
    
    print("\n" + "="*80)
    input("Нажмите Enter для выхода...")

if __name__ == "__main__":
    main()

📄 Обрабатываю: [2026_02_04]_02.02.2026_СЧА Фонд_ПДС.xls
   Дата из имени: 02.02.2026
   Размер листа: 158 строк x 27 колонок
   ⚠️ Значение не найдено

📄 Обрабатываю: [2026_02_05]_03.02.2026_СЧА Фонд_ПДС.xls
   Дата из имени: 03.02.2026
   Размер листа: 158 строк x 27 колонок
   ⚠️ Значение не найдено

📄 Обрабатываю: [2026_02_06]_04.02.2026_СЧА Фонд_ПДС.xls
   Дата из имени: 04.02.2026
   Размер листа: 158 строк x 27 колонок
   ⚠️ Значение не найдено

📄 Обрабатываю: [2026_02_09]_05.02.2026_СЧА Фонд_ПДС.xls
   Дата из имени: 05.02.2026
   Размер листа: 158 строк x 27 колонок
   ⚠️ Значение не найдено

📄 Обрабатываю: [2026_02_10]_06.02.2026_СЧА Фонд_ПДС.xls
   Дата из имени: 06.02.2026
   Размер листа: 158 строк x 27 колонок
   ⚠️ Значение не найдено

📄 Обрабатываю: [2026_02_11]_09.02.2026_СЧА Фонд_ПДС.xls
   Дата из имени: 09.02.2026
   Размер листа: 158 строк x 27 колонок
   ⚠️ Значение не найдено

📄 Обрабатываю: [2026_02_12]_10.02.2026_СЧА Фонд_ПДС.xls
   Дата из имени: 10.02.2026
   Размер листа: 158 строк x 27 колонок
   ⚠️ Значение не найдено

📄 Обрабатываю: [2026_02_13]_11.02.2026_СЧА Фонд_ПДС.xls
   Дата из имени: 11.02.2026
   Размер листа: 158 строк x 27 колонок
   ⚠️ Значение не найдено
