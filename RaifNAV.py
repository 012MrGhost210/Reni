import pandas as pd
import os
import glob
import re

def natural_sort_key(sheet_name):
    """Функция для естественной сортировки (256, 257 и т.д.)"""
    return [int(text) if text.isdigit() else text.lower() 
            for text in re.split('([0-9]+)', sheet_name)]

# Путь к папке Documents
docs_path = r'C:\Users\ytggf\OneDrive\Документы'

# Ищем файл
search_pattern = os.path.join(docs_path, '**', '*Отчет по СЧА*.xlsx')
found_files = glob.glob(search_pattern, recursive=True)

if found_files:
    print(f"Найдено {len(found_files)} файлов:")
    for i, file in enumerate(found_files):
        print(f"{i+1}. {file}")
    
    source_file = found_files[0]
    print(f"\nОбрабатываем файл: {source_file}")
    
    try:
        # Читаем все листы
        excel_file = pd.ExcelFile(source_file)
        sheet_names = excel_file.sheet_names
        print(f"Найдены листы: {sheet_names}")
        
        # Сортируем листы в естественном порядке
        sheet_names.sort(key=natural_sort_key)
        print(f"Листы после сортировки: {sheet_names}")
        
        # Словари для хранения данных
        scha_data = {}  # для СЧА
        inout_data = {}  # для вводов-выводов
        
        # Проходим по каждому листу
        for sheet in sheet_names:
            try:
                # Читаем данные, пропуская первые 6 строк (шапка)
                df = pd.read_excel(source_file, sheet_name=sheet, skiprows=6, header=None)
                
                # Убираем полностью пустые колонки
                df = df.dropna(axis=1, how='all')
                
                # В этом файле всегда 5 колонок (№, Дата, Вводы, Выводы, СЧА)
                if len(df.columns) == 5:
                    df.columns = ['№', 'Date', 'Вводы', 'Выводы', 'СЧА']
                else:
                    print(f"  Лист '{sheet}' пропущен: неожиданное кол-во колонок {len(df.columns)}")
                    continue
                
                # Удаляем пустые строки
                df = df.dropna(subset=['Date'])
                
                # Фильтруем только строки с датами (не итоговые)
                df = df[df['Date'].astype(str).str.contains('Суммарная|Количество|Средняя|№ п/п', na=False) == False]
                
                # Нормализуем даты
                df['Date'] = pd.to_datetime(df['Date'], format='%d.%m.%Y', errors='coerce').dt.date
                
                # Удаляем строки с некорректными датами
                df = df.dropna(subset=['Date'])
                
                if len(df) > 0:
                    # --- Обработка СЧА ---
                    scha_temp = df[['Date', 'СЧА']].copy()
                    
                    # Пробуем вычислить формулы Excel
                    # Сначала пробуем преобразовать как есть
                    scha_temp['СЧА'] = pd.to_numeric(scha_temp['СЧА'], errors='coerce')
                    
                    # Если есть пропуски (из-за формул), пробуем вычислить последовательно
                    if scha_temp['СЧА'].isna().any():
                        # Находим первое числовое значение
                        first_valid = scha_temp['СЧА'].first_valid_index()
                        if first_valid is not None:
                            base_value = scha_temp.loc[first_valid, 'СЧА']
                            # Заполняем остальные, увеличивая на 1
                            for idx in scha_temp.index:
                                if pd.isna(scha_temp.loc[idx, 'СЧА']):
                                    days_diff = (pd.to_datetime(scha_temp.loc[idx, 'Date']) - 
                                                pd.to_datetime(scha_temp.loc[first_valid, 'Date'])).days
                                    scha_temp.loc[idx, 'СЧА'] = base_value + days_diff
                    
                    scha_temp = scha_temp.dropna(subset=['СЧА'])
                    scha_temp = scha_temp.rename(columns={'СЧА': sheet})
                    scha_data[sheet] = scha_temp
                    
                    # --- Обработка InOut (Вводы - Выводы) ---
                    inout_temp = df[['Date', 'Вводы', 'Выводы']].copy()
                    
                    # Преобразуем в числа
                    inout_temp['Вводы'] = pd.to_numeric(inout_temp['Вводы'], errors='coerce').fillna(0)
                    inout_temp['Выводы'] = pd.to_numeric(inout_temp['Выводы'], errors='coerce').fillna(0)
                    
                    # Считаем чистое движение: Вводы - Выводы
                    inout_temp[sheet] = inout_temp['Вводы'] - inout_temp['Выводы']
                    
                    inout_temp = inout_temp[['Date', sheet]].copy()
                    inout_temp = inout_temp.dropna(subset=[sheet])
                    
                    inout_data[sheet] = inout_temp
                    
                    print(f"✓ Лист '{sheet}': СЧА: {len(scha_temp)} дат, InOut: {len(inout_temp)} дат")
                
            except Exception as e:
                print(f"✗ Ошибка листа '{sheet}': {e}")
        
        # Проверяем, есть ли данные
        if not scha_data and not inout_data:
            print("❌ Не удалось собрать данные ни с одного листа")
        else:
            # Функция для объединения данных
            def combine_data(data_dict):
                if data_dict:
                    # Получаем отсортированный список ключей
                    sheets_list = sorted(data_dict.keys(), key=natural_sort_key)
                    result = data_dict[sheets_list[0]]
                    
                    for sheet in sheets_list[1:]:
                        result = pd.merge(result, data_dict[sheet], on='Date', how='outer')
                    
                    result = result.sort_values('Date')
                    result = result.drop_duplicates(subset=['Date'])
                    result['Date'] = pd.to_datetime(result['Date'])
                    
                    return result
                return None
            
            # Объединяем данные
            scha_result = combine_data(scha_data)
            inout_result = combine_data(inout_data)
            
            # Сохраняем
            output_path = os.path.join(os.path.dirname(source_file), "Сводная_СЧА_InOut.xlsx")
            
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                if scha_result is not None:
                    scha_result.to_excel(writer, sheet_name='СЧА', index=False)
                    print(f"\n✅ Лист СЧА: {len(scha_result)} строк")
                    print(f"   Колонки: {list(scha_result.columns)}")
                
                if inout_result is not None:
                    inout_result.to_excel(writer, sheet_name='InOut', index=False)
                    print(f"✅ Лист InOut: {len(inout_result)} строк")
                    print(f"   Колонки: {list(inout_result.columns)}")
            
            print(f"\n✅ Готово! Файл сохранен: {output_path}")
            
            # Показываем примеры
            if scha_result is not None:
                print("\n--- Первые 5 строк СЧА ---")
                print(scha_result.head(5))
            
            if inout_result is not None:
                print("\n--- Первые 5 строк InOut ---")
                print(inout_result.head(5))
            
            # Проверка дубликатов
            if scha_result is not None:
                scha_dupl = scha_result[scha_result.duplicated(subset=['Date'], keep=False)]
                print(f"\nДубликаты дат в СЧА: {len(scha_dupl)}")
            
            if inout_result is not None:
                inout_dupl = inout_result[inout_result.duplicated(subset=['Date'], keep=False)]
                print(f"Дубликаты дат в InOut: {len(inout_dupl)}")
            
    except Exception as e:
        print(f"❌ Ошибка при обработке файла: {e}")
        
else:
    print("❌ Файлы 'Отчет по СЧА' не найдены")
    
    # Поищем в других местах
    other_paths = [
        r'C:\Users\ytggf\Downloads',
        r'C:\Users\ytggf\Desktop',
        r'C:\Users\ytggf\OneDrive\Рабочий стол'
    ]
    
    print("\nПоиск в других местах...")
    for path in other_paths:
        if os.path.exists(path):
            search_pattern = os.path.join(path, '*Отчет по СЧА*.xlsx')
            files = glob.glob(search_pattern, recursive=False)
            for file in files:
                print(f"Найден: {file}")
