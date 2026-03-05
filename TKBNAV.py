import pandas as pd
import os
import glob

# Путь к папке Documents
docs_path = r'C:\Users\ytggf\OneDrive\Документы'

# Ищем файл
search_pattern = os.path.join(docs_path, '**', '*Сводная РСА-СЧА*.xlsx')
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
                
                # Смотрим сколько колонок осталось
                print(f"\nЛист '{sheet}' - колонок после очистки: {len(df.columns)}")
                
                # Если осталось 7 колонок - используем наши названия
                if len(df.columns) == 7:
                    df.columns = ['№', 'Date', 'Вводы', 'Выводы', 'РСА', 'СЧА', 'Пусто']
                    df = df.drop(columns=['Пусто'])
                # Если осталось 6 колонок
                elif len(df.columns) == 6:
                    df.columns = ['№', 'Date', 'Вводы', 'Выводы', 'РСА', 'СЧА']
                else:
                    print(f"  Неожиданное количество колонок: {len(df.columns)}")
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
                    scha_temp['СЧА'] = pd.to_numeric(scha_temp['СЧА'], errors='coerce')
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
                    sheets_list = list(data_dict.keys())
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
                
                if inout_result is not None:
                    inout_result.to_excel(writer, sheet_name='InOut', index=False)
                    print(f"✅ Лист InOut: {len(inout_result)} строк")
            
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
    print("❌ Файлы 'Сводная РСА-СЧА' не найдены")
    
    # Поищем в других местах
    other_paths = [
        r'C:\Users\ytggf\Downloads',
        r'C:\Users\ytggf\Desktop',
        r'C:\Users\ytggf\OneDrive\Рабочий стол'
    ]
    
    print("\nПоиск в других местах...")
    for path in other_paths:
        if os.path.exists(path):
            search_pattern = os.path.join(path, '*Сводная РСА-СЧА*.xlsx')
            files = glob.glob(search_pattern, recursive=False)
            for file in files:
                print(f"Найден: {file}")
