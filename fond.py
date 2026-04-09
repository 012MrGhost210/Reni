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
    
    def extract_number_from_string(self, text):
        """
        Извлекает число из строки любого формата
        Примеры: '3 451 553,68 руб.', '2662408.68', '2,662,408.68'
        """
        if text is None:
            return None
        
        # Если уже число
        if isinstance(text, (int, float)):
            return float(text)
        
        # Преобразуем в строку
        text = str(text)
        
        # Ищем число в формате: пробелы или запятые как разделители тысяч, запятая или точка как десятичный разделитель
        # Паттерн находит: цифры, пробелы, запятые, точки, затем опционально руб
        match = re.search(r'([\d\s]+[.,]?\d*)\s*руб', text, re.IGNORECASE)
        if not match:
            # Если нет "руб", ищем просто число
            match = re.search(r'[\d\s]+[.,]?\d*', text)
        
        if match:
            number_str = match.group(1) if 'руб' in text.lower() else match.group(0)
            # Убираем все пробелы
            number_str = re.sub(r'\s', '', number_str)
            # Заменяем запятую на точку (десятичный разделитель)
            number_str = number_str.replace(',', '.')
            
            try:
                return float(number_str)
            except:
                pass
        
        return None
    
    def find_net_asset_value(self, sheet):
        """
        Ищет 'Итого стоимость чистых активов' и возвращает число из столбца P (индекс 15)
        """
        search_text = "Итого стоимость чистых активов"
        target_col = 15  # P = 15
        
        for row_idx in range(sheet.nrows):
            # Проверяем ячейку в столбце A (индекс 0) на наличие ключевой фразы
            if sheet.ncols > 0:
                cell_value = sheet.cell(row_idx, 0).value
                if cell_value and isinstance(cell_value, str):
                    if search_text.lower() in cell_value.lower():
                        print(f"      ✅ Найдено ключевое слово в строке {row_idx}, столбец A")
                        
                        # Берем значение из столбца P (индекс 15)
                        if target_col < sheet.ncols:
                            value_cell = sheet.cell(row_idx, target_col)
                            value = value_cell.value
                            print(f"      📍 Значение в столбце P: '{value}'")
                            
                            # Извлекаем число
                            parsed_value = self.extract_number_from_string(value)
                            if parsed_value is not None:
                                print(f"      ✅ Извлечено число: {parsed_value}")
                                return parsed_value
                            else:
                                print(f"      ⚠️ Не удалось извлечь число из '{value}'")
                                
                                # Пробуем посмотреть другие столбцы в этой строке
                                print(f"      🔍 Ищем число в других столбцах строки {row_idx}:")
                                for col in range(sheet.ncols):
                                    val = sheet.cell(row_idx, col).value
                                    num = self.extract_number_from_string(val)
                                    if num is not None:
                                        col_letter = self.get_column_letter(col)
                                        print(f"         Найдено число в столбце {col_letter}: {num}")
                                        return num
                        else:
                            print(f"      ⚠️ Столбца P (индекс 15) нет в файле!")
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
                print(f"   ⚠️ Значение 'Итого стоимость чистых активов' не найдено")
                
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
        print("ПАРСИНГ EXCEL - ПОИСК 'ИТОГО СТОИМОСТЬ ЧИСТЫХ АКТИВОВ'")
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
            print("-"*80)
        
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
