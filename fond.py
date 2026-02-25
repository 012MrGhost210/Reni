import os
import re
from pathlib import Path
import openpyxl
import pandas as pd
from datetime import datetime

# Путь к папке с файлами
input_folder = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\Фонд СЧА"
output_file = Path(input_folder) / "!_РЕЗУЛЬТАТЫ_ПАРСИНГА.csv"

print("="*80)
print("ПАРСИНГ EXCEL ФАЙЛОВ СЧА Фонд_ПДС")
print("="*80)
print(f"📂 Папка с файлами: {input_folder}")
print(f"📄 Результат будет сохранен в: {output_file}")
print("="*80)

# Результаты
results = []

# Получаем все Excel файлы
excel_files = list(Path(input_folder).glob("*.xls*"))
print(f"\nНайдено Excel файлов: {len(excel_files)}")

for file_path in excel_files:
    print(f"\n📄 Обрабатываю: {file_path.name}")
    
    try:
        # 1. Извлекаем дату из имени файла
        # Формат: [2026_01_12]_29.12.2025_СЧА Фонд_ПДС.xls
        date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', file_path.name)
        if date_match:
            file_date = date_match.group(1)
            print(f"   Дата из имени: {file_date}")
        else:
            file_date = "Не найдена"
            print(f"   ⚠️ Дата не найдена в имени")
        
        # 2. Открываем Excel файл
        found_value = None
        
        # Пробуем открыть через openpyxl (для .xlsx)
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            sheet = wb.active
            
            # Ищем строку с ГАЗПРОМБАНК
            search_text = "ГАЗПРОМБАНК"
            
            for row in sheet.iter_rows(values_only=True):
                for cell in row:
                    if cell and search_text in str(cell):
                        # Нашли ячейку с ГАЗПРОМБАНК
                        print(f"   ✅ Найдена строка с ГАЗПРОМБАНК")
                        
                        # Получаем индекс строки и столбца
                        row_idx = row
                        # Ищем значение справа (на 8 позиций)
                        # Это сложно, попробуем найти числовое значение в этой строке
                        numbers_in_row = [c for c in row if isinstance(c, (int, float))]
                        if numbers_in_row:
                            # Берем последнее числовое значение в строке
                            found_value = numbers_in_row[-1]
                            print(f"   💰 Найдено значение: {found_value}")
                        break
                if found_value:
                    break
            
            wb.close()
            
        except Exception as e:
            print(f"   ⚠️ Ошибка при открытии через openpyxl: {e}")
            
            # Пробуем через pandas как запасной вариант
            try:
                df = pd.read_excel(file_path, header=None)
                
                # Ищем строку с ГАЗПРОМБАНК
                for idx, row in df.iterrows():
                    for cell in row:
                        if cell and search_text in str(cell):
                            print(f"   ✅ Найдена строка с ГАЗПРОМБАНК (pandas)")
                            
                            # Ищем числовые значения в этой строке
                            numeric_values = row[pd.to_numeric(row, errors='coerce').notna()]
                            if not numeric_values.empty:
                                found_value = numeric_values.iloc[-1]
                                print(f"   💰 Найдено значение: {found_value}")
                            break
                    if found_value:
                        break
                        
            except Exception as e2:
                print(f"   ❌ Ошибка и с pandas: {e2}")
        
        # Сохраняем результат
        results.append({
            'Файл': file_path.name,
            'Дата_из_имени': file_date,
            'Найдено_значение': found_value if found_value else "Не найдено"
        })
        
    except Exception as e:
        print(f"   ❌ Ошибка обработки файла: {e}")
        results.append({
            'Файл': file_path.name,
            'Дата_из_имени': "Ошибка",
            'Найдено_значение': f"Ошибка: {str(e)[:50]}"
        })

# Сохраняем результаты в CSV
import csv

with open(output_file, 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=['Файл', 'Дата_из_имени', 'Найдено_значение'])
    writer.writeheader()
    writer.writerows(results)

print("\n" + "="*80)
print("ГОТОВО!")
print("="*80)
print(f"Обработано файлов: {len(results)}")
print(f"Результаты сохранены в: {output_file}")
print("\nПервые 10 результатов:")
print("-"*40)

for i, row in enumerate(results[:10], 1):
    print(f"{i:2d}. {row['Дата_из_имени']} - {row['Найдено_значение']}")

print("="*80)
input("\nНажмите Enter для выхода...")
