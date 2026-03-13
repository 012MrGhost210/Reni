import pandas as pd
import re
import os
from datetime import datetime
from openpyxl.styles import numbers

# Пути к файлам
input_path = r'C:\Users\ytggf\OneDrive\Документы\renlife\Сводные ааааа\ff\Test\Аааааыывыв.xlsx'
output_path = r'C:\Users\ytggf\OneDrive\Документы\renlife\Сводные ааааа\ff\Test\us\ОСВ_плоская.xlsx'

# Создаем папку us, если её нет
os.makedirs(os.path.dirname(output_path), exist_ok=True)

print(f"📂 Читаем файл: {input_path}")
print(f"📁 Сохраняем в: {output_path}")

# Читаем файл Excel
df = pd.read_excel(input_path, sheet_name='Лист_1', header=None)
print(f"✅ Файл загружен, строк: {len(df)}")

# Инициализируем переменные
current_data = {
    'подразделение': None,
    'банковский_счет': None,
    'статья_ддс': None
}

result_rows = []
row_idx = 0
total_rows = len(df)

# Функция для проверки, является ли строка статьей ДДС
def is_article(text):
    if not text or len(text) < 3:
        return False
    
    skip_words = ['Обороты за', 'Итого', 'Всего', 'Банковские счета', 
                  'Статьи движения', 'Сальдо', 'Выводимые данные']
    
    if any(skip in text for skip in skip_words):
        return False
    
    if text[0].isdigit():
        return False
    
    if 'Основное подразделение' in text:
        return False
    
    if ',' in text and sum(c.isdigit() for c in text) > 10:
        return False
    
    if text.startswith('Обороты за'):
        return False
    
    return True

while row_idx < total_rows:
    row = df.iloc[row_idx]
    first_col = str(row[0]) if pd.notna(row[0]) else ""
    first_col = first_col.strip()
    
    if not first_col or first_col == "nan" or row_idx < 10:
        row_idx += 1
        continue
    
    # ПОИСК ПОДРАЗДЕЛЕНИЙ
    if 'Основное подразделение' in first_col:
        current_data['подразделение'] = first_col
        current_data['банковский_счет'] = None
        current_data['статья_ддс'] = None
        
        if row_idx + 1 < total_rows:
            next_row = df.iloc[row_idx + 1]
            next_first = str(next_row[0]) if pd.notna(next_row[0]) else ""
            
            if next_first and any(c.isdigit() for c in next_first) and ',' in next_first:
                current_data['банковский_счет'] = next_first
                row_idx += 1
        
        row_idx += 1
        continue
    
    # ПОИСК СТАТЕЙ ДДС
    if is_article(first_col):
        current_data['статья_ддс'] = first_col
        row_idx += 1
        continue
    
    # ПОИСК ОБОРОТОВ ПО ДАТАМ
    if 'Обороты за' in first_col:
        date_match = re.search(r'(\d{2}\.\d{2}\.\d{2})', first_col)
        if date_match:
            date_str = date_match.group(1)
            
            # Преобразуем в формат с полным годом
            try:
                day, month, year = date_str.split('.')
                full_year = f"20{year}"
                date_with_year = f"{day}.{month}.{full_year}"
                excel_date = datetime.strptime(date_with_year, '%d.%m.%Y')
            except:
                excel_date = None
            
            # Получаем дебет и кредит
            debit = row[3] if pd.notna(row[3]) else 0
            credit = row[4] if pd.notna(row[4]) else 0
            
            try:
                debit = float(str(debit).replace(',', '.')) if debit != 0 else 0
                credit = float(str(credit).replace(',', '.')) if credit != 0 else 0
            except:
                debit = 0
                credit = 0
            
            if debit != 0 or credit != 0:
                result_rows.append({
                    'Подразделение': current_data.get('подразделение', ''),
                    'Банковский_счет': current_data.get('банковский_счет', ''),
                    'Статья_ДДС': current_data.get('статья_ддс', ''),
                    'Дата': excel_date,
                    'Дебет': debit,
                    'Кредит': credit
                })
    
    row_idx += 1

# Создаем DataFrame
if result_rows:
    result_df = pd.DataFrame(result_rows)
    
    # Удаляем строки с нулями
    result_df = result_df[(result_df['Дебет'] != 0) | (result_df['Кредит'] != 0)]
    
    # Заменяем None на пустую строку
    text_columns = ['Подразделение', 'Банковский_счет', 'Статья_ДДС']
    for col in text_columns:
        result_df[col] = result_df[col].fillna('')
    
    # Убираем дубликаты
    result_df = result_df.drop_duplicates()
    
    # Сортируем по дате
    result_df = result_df.sort_values('Дата')
    
    # Определяем, какие столбцы оставить
    columns_to_keep = []
    for col in result_df.columns:
        if col in ['Дебет', 'Кредит', 'Дата']:
            columns_to_keep.append(col)
        else:
            non_empty = result_df[col].astype(str).str.strip().str.len() > 0
            if non_empty.any():
                columns_to_keep.append(col)
    
    # Оставляем только нужные столбцы
    result_df = result_df[columns_to_keep]
    
    # Сохраняем с форматированием
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        result_df.to_excel(writer, sheet_name='ОСВ_плоская', index=False, startcol=0)
        worksheet = writer.sheets['ОСВ_плоская']
        
        # Настраиваем формат даты
        if 'Дата' in columns_to_keep:
            date_col_idx = columns_to_keep.index('Дата') + 1
            date_col_letter = chr(64 + date_col_idx)
            for row in range(2, len(result_df) + 2):
                cell = worksheet[f'{date_col_letter}{row}']
                cell.number_format = 'DD.MM.YYYY'
        
        # Настраиваем ширину колонок и формат чисел
        for idx, col in enumerate(columns_to_keep, 1):
            col_letter = chr(64 + idx)
            
            if col == 'Подразделение':
                worksheet.column_dimensions[col_letter].width = 40
            elif col == 'Банковский_счет':
                worksheet.column_dimensions[col_letter].width = 70
            elif col == 'Статья_ДДС':
                worksheet.column_dimensions[col_letter].width = 70
            elif col == 'Дата':
                worksheet.column_dimensions[col_letter].width = 15
            elif col in ['Дебет', 'Кредит']:
                worksheet.column_dimensions[col_letter].width = 20
                for row in range(2, len(result_df) + 2):
                    cell = worksheet[f'{col_letter}{row}']
                    cell.number_format = '#,##0.00'
    
    # ТОЛЬКО ФИНАЛЬНАЯ СТАТИСТИКА
    print("\n" + "="*70)
    print(f"✅ УСПЕШНО! Создано {len(result_df)} строк")
    print(f"📁 Файл сохранен: {output_path}")
    print("="*70)
    print(f"\n📊 СТАТИСТИКА:")
    print(f"   Всего операций: {len(result_df)}")
    print(f"   Сумма по дебету: {result_df['Дебет'].sum():,.2f}")
    print(f"   Сумма по кредиту: {result_df['Кредит'].sum():,.2f}")
    
    # Сохраняем CSV
    csv_path = output_path.replace('.xlsx', '.csv')
    csv_df = result_df.copy()
    if 'Дата' in csv_df.columns:
        csv_df['Дата'] = csv_df['Дата'].dt.strftime('%d.%m.%Y')
    csv_df.to_csv(csv_path, index=False, encoding='utf-8-sig')
    print(f"\n📁 Также сохранено в CSV: {csv_path}")
    
else:
    print("❌ Не найдено операций!")

print("\n🎯 Готово!")
