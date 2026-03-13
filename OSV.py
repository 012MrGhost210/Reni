import pandas as pd
import re
import os
from datetime import datetime

# Пути к файлам
input_path = r'C:\Users\ytggf\OneDrive\Документы\renlife\Сводные ааааа\ff\Test\Аааааыывыв.xlsx'
output_path = r'C:\Users\ytggf\OneDrive\Документы\renlife\Сводные ааааа\ff\Test\us\ОСВ_плоская.xlsx'

# Создаем папку us, если её нет
os.makedirs(os.path.dirname(output_path), exist_ok=True)

print(f"📂 Читаем файл: {input_path}")
print(f"📁 Сохраняем в: {output_path}")

# Читаем файл Excel, пропускаем первые колонки (A, B, C - они пустые)
# Нам нужны колонки D (дебет) и E (кредит) - это индексы 3 и 4
# Но нам также нужна колонка A для анализа структуры
df = pd.read_excel(input_path, sheet_name='Лист_1', header=None)
print(f"✅ Файл загружен, строк: {len(df)}")

# Инициализируем переменные
current_data = {
    'счет': '20501',
    'подразделение': None,
    'банковский_счет': None,
    'статья_ддс': None
}

result_rows = []
row_idx = 0
total_rows = len(df)

print("🔄 Начинаем обработку...")

# Функция для проверки, является ли строка статьей ДДС
def is_article(text):
    if not text or len(text) < 3:
        return False
    
    # Ключевые слова, которые НЕ являются статьями ДДС
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
        print(f"\n📍 Найдено подразделение [{row_idx}]: {first_col}")
        
        if row_idx + 1 < total_rows:
            next_row = df.iloc[row_idx + 1]
            next_first = str(next_row[0]) if pd.notna(next_row[0]) else ""
            
            if next_first and any(c.isdigit() for c in next_first) and ',' in next_first:
                current_data['банковский_счет'] = next_first
                print(f"🏦 Найден банковский счет [{row_idx + 1}]: {next_first[:70]}...")
                row_idx += 1
        
        row_idx += 1
        continue
    
    # ПОИСК СТАТЕЙ ДДС
    if is_article(first_col):
        current_data['статья_ддс'] = first_col
        print(f"📝 Найдена статья [{row_idx}]: {first_col}")
        row_idx += 1
        continue
    
    # ПОИСК ОБОРОТОВ ПО ДАТАМ
    if 'Обороты за' in first_col:
        date_match = re.search(r'(\d{2}\.\d{2}\.\d{2})', first_col)
        if date_match:
            date_str = date_match.group(1)  # Например: "13.02.26"
            
            # ПРЕОБРАЗУЕМ В ФОРМАТ ДД.ММ.ГГГГ (с 20 в начале года)
            try:
                # Преобразуем "13.02.26" в "13.02.2026"
                day, month, year = date_str.split('.')
                full_year = f"20{year}"  # 26 -> 2026
                date_with_year = f"{day}.{month}.{full_year}"
                
                # Создаем объект datetime для Excel
                excel_date = datetime.strptime(date_with_year, '%d.%m.%Y')
                display_date = date_with_year
            except:
                excel_date = None
                display_date = date_str
            
            # Получаем дебет и кредит из колонок D и E (индексы 3 и 4)
            debit = row[3] if pd.notna(row[3]) else 0
            credit = row[4] if pd.notna(row[4]) else 0
            
            # Преобразуем в числа
            try:
                debit = float(str(debit).replace(',', '.')) if debit != 0 else 0
                credit = float(str(credit).replace(',', '.')) if credit != 0 else 0
            except:
                debit = 0
                credit = 0
            
            # Добавляем запись, если есть движение
            if debit != 0 or credit != 0:
                result_rows.append({
                    'Счет': current_data['счет'],
                    'Подразделение': current_data.get('подразделение', ''),
                    'Банковский_счет': current_data.get('банковский_счет', ''),
                    'Статья_ДДС': current_data.get('статья_ддс', ''),
                    'Дата': excel_date,  # Сохраняем как datetime для Excel
                    'Дебет': debit,
                    'Кредит': credit
                })
                
                article_display = current_data.get('статья_ддс', '')
                if article_display is None:
                    article_display = ''
                print(f"  ➕ Операция: {display_date} | Д:{debit:>15,.0f} | К:{credit:>15,.0f} | Ст: {str(article_display)[:40]}")
    
    row_idx += 1

# Создаем DataFrame
if result_rows:
    result_df = pd.DataFrame(result_rows)
    
    # Удаляем строки с нулями
    result_df = result_df[(result_df['Дебет'] != 0) | (result_df['Кредит'] != 0)]
    
    # Заменяем None на пустую строку
    text_columns = ['Счет', 'Подразделение', 'Банковский_счет', 'Статья_ДДС']
    for col in text_columns:
        result_df[col] = result_df[col].fillna('')
    
    # Убираем дубликаты
    result_df = result_df.drop_duplicates()
    
    # Сортируем по дате
    result_df = result_df.sort_values('Дата')
    
    # Переупорядочиваем колонки - теперь у нас только нужные колонки (без A,B,C)
    columns = ['Счет', 'Подразделение', 'Банковский_счет', 'Статья_ДДС', 'Дата', 'Дебет', 'Кредит']
    result_df = result_df[columns]
    
    # Сохраняем с правильным форматированием
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Преобразуем даты в формат с полным годом для отображения
        result_df.to_excel(writer, sheet_name='ОСВ_плоская', index=False, startcol=0)
        
        # Получаем лист для форматирования
        worksheet = writer.sheets['ОСВ_плоская']
        
        # Настраиваем формат даты (ДД.ММ.ГГГГ)
        from openpyxl.styles import numbers
        date_column = 'E'  # Колонка с датой (после удаления A,B,C это будет колонка E)
        for row in range(2, len(result_df) + 2):  # +2 из-за заголовка и индексации с 1
            cell = worksheet[f'{date_column}{row}']
            cell.number_format = 'DD.MM.YYYY'  # Формат даты для Excel
        
        # Настраиваем ширину колонок
        column_widths = {
            'A': 10,  # Счет
            'B': 40,  # Подразделение
            'C': 70,  # Банковский счет
            'D': 70,  # Статья ДДС
            'E': 15,  # Дата
            'F': 20,  # Дебет
            'G': 20   # Кредит
        }
        
        for col, width in column_widths.items():
            worksheet.column_dimensions[col].width = width
        
        # Форматируем числа с разделителями
        for col in ['F', 'G']:  # Дебет и Кредит
            for row in range(2, len(result_df) + 2):
                cell = worksheet[f'{col}{row}']
                cell.number_format = '#,##0.00'  # Формат чисел с разделителями
    
    print("\n" + "="*70)
    print(f"✅ УСПЕШНО! Создано {len(result_df)} строк")
    print(f"📁 Файл сохранен: {output_path}")
    print("="*70)
    
    # Статистика
    print(f"\n📊 СТАТИСТИКА:")
    print(f"   Всего операций: {len(result_df)}")
    print(f"   Сумма дебетов: {result_df['Дебет'].sum():,.2f}")
    print(f"   Сумма кредитов: {result_df['Кредит'].sum():,.2f}")
    
    # Диапазон дат
    if not result_df['Дата'].isna().all():
        min_date = result_df['Дата'].min()
        max_date = result_df['Дата'].max()
        print(f"   Период: {min_date.strftime('%d.%m.%Y')} - {max_date.strftime('%d.%m.%Y')}")
    
    # Статистика по подразделениям
    print(f"\n🏢 ПОДРАЗДЕЛЕНИЯ:")
    for podr in result_df['Подразделение'].unique():
        if podr:
            count = len(result_df[result_df['Подразделение'] == podr])
            print(f"   📍 {podr}: {count} операций")
    
    # Статистика по статьям ДДС
    print(f"\n📋 СТАТЬИ ДДС (топ-10):")
    article_stats = result_df['Статья_ДДС'].value_counts().head(10)
    for article, count in article_stats.items():
        if article:
            print(f"   • {article[:70]}: {count} операций")
    
    # Сохраняем CSV для совместимости
    csv_path = output_path.replace('.xlsx', '.csv')
    csv_df = result_df.copy()
    csv_df['Дата'] = csv_df['Дата'].dt.strftime('%d.%m.%Y')  # В CSV тоже полный год
    csv_df.to_csv(csv_path, index=False, encoding='utf-8-sig')
    print(f"\n📁 Также сохранено в CSV: {csv_path}")
    
    # Показываем пример данных
    print("\n📋 ПРИМЕР ДАННЫХ (первые 5 строк):")
    example = result_df.head(5).copy()
    example['Дата'] = example['Дата'].dt.strftime('%d.%m.%Y')
    print(example.to_string())
    
else:
    print("❌ Не найдено операций!")

print("\n🎯 Готово!")
