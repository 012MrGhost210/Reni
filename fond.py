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
        print(f"         [clean_number] Вход: '{value}' (тип: {type(value)})")
        
        if value is None:
            print(f"         [clean_number] None, возвращаем None")
            return None
        
        # Если уже число
        if isinstance(value, (int, float)):
            print(f"         [clean_number] Уже число: {value}")
            return float(value)
        
        # Преобразуем в строку
        value_str = str(value).strip()
        print(f"         [clean_number] После str(): '{value_str}'")
        
        # Удаляем слово "руб" и всё после него
        if 'руб' in value_str.lower():
            value_str = value_str.lower().split('руб')[0].strip()
            print(f"         [clean_number] После удаления 'руб': '{value_str}'")
        
        # Удаляем все пробелы (включая неразрывные)
        value_str = re.sub(r'\s+', '', value_str)
        print(f"         [clean_number] После удаления пробелов: '{value_str}'")
        
        # Заменяем запятую на точку (десятичный разделитель)
        value_str = value_str.replace(',', '.')
        print(f"         [clean_number] После замены запятой: '{value_str}'")
        
        # Удаляем всё кроме цифр и точки
        value_str = re.sub(r'[^\d.]', '', value_str)
        print(f"         [clean_number] После удаления лишних символов: '{value_str}'")
        
        # Если несколько точек, оставляем только последнюю
        if value_str.count('.') > 1:
            parts = value_str.split('.')
            value_str = ''.join(parts[:-1]) + '.' + parts[-1]
            print(f"         [clean_number] После обработки нескольких точек: '{value_str}'")
        
        try:
            result = float(value_str)
            print(f"         [clean_number] Успешно преобразовано в: {result}")
            return result
        except ValueError as e:
            print(f"         [clean_number] Ошибка преобразования: {e}")
            return None
    
    def find_net_asset_value(self, sheet):
        """
        Ищет 'Итого стоимость чистых активов' в столбце A
        и возвращает число из столбца P (индекс 15)
        """
        search_text = "Итого стоимость чистых активов"
        target_col = 15  # P = 15
        
        print(f"   Поиск фразы: '{search_text}'")
        
        for row_idx in range(sheet.nrows):
            # Проверяем столбец A (индекс 0)
            if sheet.ncols > 0:
                cell_value = sheet.cell(row_idx, 0).value
                if cell_value and isinstance(cell_value, str):
                    if search_text.lower() in cell_value.lower():
                        print(f"      ✅ Найдено в строке {row_idx}")
                        print(f"      Текст в A: '{cell_value}'")
                        
                        # Берем значение из столбца P
                        if target_col < sheet.ncols:
                            raw_value = sheet.cell(row_idx, target_col).value
                            print(f"      📍 Сырое значение в P (индекс {target_col}): '{raw_value}'")
                            print(f"      📍 Тип raw_value: {type(raw_value)}")
                            
                            # Очищаем и преобразуем
                            number = self.clean_number(raw_value)
                            if number is not None:
                                print(f"      ✅ Преобразовано в число: {number}")
                                return number
                            else:
                                print(f"      ❌ Не удалось преобразовать: '{raw_value}'")
                                
                                # Пробуем посмотреть все столбцы в этой строке
                                print(f"      🔍 Проверяем все столбцы строки {row_idx}:")
                                for col in range(min(sheet.ncols, 30)):
                                    val = sheet.cell(row_idx, col).value
                                    if val and str(val).strip():
                                        col_letter = self.get_column_letter(col)
                                        print(f"         Столбец {col_letter} (инд.{col}): '{val}'")
                                        num = self.clean_number(val)
                                        if num is not None:
                                            print(f"         🎯 Нашли число в столбце {col_letter}: {num}")
                                            return num
                        else:
                            print(f"      ❌ Столбца P (индекс {target_col}) нет в файле (всего столбцов: {sheet.ncols})")
                        return None
        return None
    
    def get_column_letter(self, col_idx):
        """Преобразует индекс столбца в букву"""
        result = ""
        while col_idx >= 0:
            result = chr(65 + (col_idx % 26)) + result
            col_idx = col_idx // 26 - 1
        return result
    
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
            
            # Ищем значение
            found_value = self.find_net_asset_value(sheet)
            
            if found_value is not None:
                value_str = f"{found_value:,.2f}".replace(',', ' ')
                print(f"   ✅ Найдено: {value_str} руб.")
            else:
                print(f"   ❌ Значение не найдено")
                
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
        print("ПАРСИНГ - СТОИМОСТЬ ЧИСТЫХ АКТИВОВ (ОТЛАДКА)")
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
            print("\n" + "="*80)
            input("Нажмите Enter для продолжения...")
        
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
