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
    
    def parse_number_from_string(self, value):
        """
        Преобразует строку типа '131 231 руб.' или '131231' в число
        """
        # Если уже число
        if isinstance(value, (float, int)):
            return float(value)
        
        # Если строка
        if isinstance(value, str):
            # Удаляем 'руб.', 'руб', 'р.' и пробелы
            cleaned = value.replace('руб', '').replace('р.', '').strip()
            cleaned = cleaned.replace(' ', '')  # убираем пробелы тысяч
            cleaned = cleaned.replace(',', '.')  # заменяем запятую на точку
            try:
                return float(cleaned)
            except:
                return None
        
        return None
    
    def find_value_by_keyword(self, sheet):
        """
        Поиск значения по ключевому слову.
        Ключевое слово ищется в ЛЮБОМ столбце.
        Число берется из столбца НОМЕР_СТОЛБЦА (укажите нужный)
        """
        # ⚠️ ЗДЕСЬ НУЖНО УКАЗАТЬ КЛЮЧЕВОЕ СЛОВО
        search_text = "ИТОГО ЧИСТЫХ АКТИВОВ"  # ← ИЗМЕНИТЕ НА НУЖНОЕ
        
        # ⚠️ ЗДЕСЬ НУЖНО УКАЗАТЬ СТОЛБЕЦ С ЧИСЛОМ
        # A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7, I=8, J=9, K=10, L=11, M=12, N=13, 
        # O=14, P=15, Q=16, R=17, S=18, T=19, U=20, V=21, W=22, X=23, Y=24, Z=25
        target_col = 15  # ← ИЗМЕНИТЕ НА НУЖНЫЙ СТОЛБЕЦ (15 = P)
        
        for row_idx in range(sheet.nrows):
            row = sheet.row(row_idx)
            for col_idx, cell in enumerate(row):
                cell_value = cell.value
                if cell_value and search_text.lower() in str(cell_value).lower():
                    # Нашли ключевое слово, берем значение из целевого столбца
                    if target_col < sheet.ncols:
                        value_cell = sheet.cell(row_idx, target_col)
                        value = self.parse_number_from_string(value_cell.value)
                        if value is not None:
                            return value
                        else:
                            # Если не число, выводим для отладки
                            print(f"      Отладка: значение в столбце {target_col} = '{value_cell.value}'")
                    else:
                        print(f"      ⚠️ Столбец {target_col} не существует (всего столбцов: {sheet.ncols})")
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
            
            # Ищем значение по ключевому слову
            found_value = self.find_value_by_keyword(sheet)
            
            if found_value is not None:
                value_str = f"{found_value:,.2f}".replace(',', ' ')
                print(f"   ✅ Найдено: {value_str} руб.")
            else:
                print(f"   ⚠️ Ключевое слово не найдено")
                
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
        print("ПАРСИНГ EXCEL ФАЙЛОВ - ПОИСК ПО КЛЮЧЕВОМУ СЛОВУ")
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
                writer.writerow(['Дата', 'Значение (руб.)', 'Файл'])
                
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
    output_file = Path(input_folder) / f"!_РЕЗУЛЬТАТЫ_НОВЫЙ.csv"
    
    # Создаем парсер и запускаем
    parser = ExcelParser(input_folder, output_file)
    parser.run()
    

if __name__ == "__main__":
    main()
