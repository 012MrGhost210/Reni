import pandas as pd
import os
import glob

# Путь к папке Documents
docs_path = r'C:\Users\ytggf\OneDrive\Документы'

# Ищем все файлы .xls и .xlsx с ключевым словом "Вознаграждение"
search_pattern = os.path.join(docs_path, '**', '*Вознаграждение*.xls*')
found_files = glob.glob(search_pattern, recursive=True)

if found_files:
    print(f"Найдено {len(found_files)} файлов:")
    for i, file in enumerate(found_files):
        print(f"{i+1}. {file}")
    
    # Берем первый найденный файл
    source_file = found_files[0]
    print(f"\nОбрабатываем файл: {source_file}")
    
    try:
        # Читаем все листы
        excel_file = pd.ExcelFile(source_file)
        sheet_names = excel_file.sheet_names
        print(f"Найдены листы: {sheet_names}")
        
        # Словари для хранения данных
        nav_data = {}
        inout_data = {}
        
        # Проходим по каждому листу
        for sheet in sheet_names:
            # Пропускаем лист "ИТОГО"
            if sheet == "ИТОГО":
                continue
            
            try:
                # Читаем данные
                df = pd.read_excel(source_file, sheet_name=sheet)
                
                # Проверяем наличие колонок
                if 'Date' in df.columns:
                    # --- Обработка NAV ---
                    if 'NAV' in df.columns:
                        nav_df = df[['Date', 'NAV']].copy()
                        
                        # Нормализуем даты
                        nav_df['Date'] = pd.to_datetime(nav_df['Date']).dt.date
                        
                        # Очищаем от пустых значений в NAV
                        nav_df = nav_df.dropna(subset=['NAV'])
                        
                        # Удаляем строки, где NAV - текст (не число)
                        nav_df = nav_df[pd.to_numeric(nav_df['NAV'], errors='coerce').notna()]
                        
                        # Группируем по дате (берем первое значение)
                        nav_df = nav_df.groupby('Date').first().reset_index()
                        
                        # Переименовываем колонку NAV в название листа
                        nav_df = nav_df.rename(columns={'NAV': sheet})
                        
                        nav_data[sheet] = nav_df
                    
                    # --- Обработка InOut ---
                    if 'InOut' in df.columns:
                        inout_df = df[['Date', 'InOut']].copy()
                        
                        # Нормализуем даты
                        inout_df['Date'] = pd.to_datetime(inout_df['Date']).dt.date
                        
                        # Очищаем от пустых значений в InOut
                        inout_df = inout_df.dropna(subset=['InOut'])
                        
                        # Удаляем строки, где InOut - текст (не число)
                        inout_df = inout_df[pd.to_numeric(inout_df['InOut'], errors='coerce').notna()]
                        
                        # Группируем по дате (берем первое значение)
                        inout_df = inout_df.groupby('Date').first().reset_index()
                        
                        # Переименовываем колонку InOut в название листа
                        inout_df = inout_df.rename(columns={'InOut': sheet})
                        
                        inout_data[sheet] = inout_df
                    
                    print(f"✓ Лист '{sheet}': NAV: {len(nav_data.get(sheet, []))} дат, InOut: {len(inout_data.get(sheet, []))} дат")
                else:
                    print(f"✗ Лист '{sheet}': нет колонки Date")
                    
            except Exception as e:
                print(f"✗ Ошибка листа '{sheet}': {e}")
        
        # Функция для объединения данных
        def combine_data(data_dict, value_name):
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
        
        # Объединяем NAV данные
        nav_result = combine_data(nav_data, 'NAV')
        
        # Объединяем InOut данные
        inout_result = combine_data(inout_data, 'InOut')
        
        # Сохраняем в одной книге Excel на разных листах
        output_path = os.path.join(os.path.dirname(source_file), "Сводный_файл_СК.xlsx")
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            if nav_result is not None:
                nav_result.to_excel(writer, sheet_name='NAV', index=False)
                print(f"\n✅ Лист NAV: {len(nav_result)} строк")
            
            if inout_result is not None:
                inout_result.to_excel(writer, sheet_name='InOut', index=False)
                print(f"✅ Лист InOut: {len(inout_result)} строк")
        
        print(f"\n✅ Готово! Файл сохранен: {output_path}")
        
        # Показываем первые несколько строк каждого листа
        if nav_result is not None:
            print("\n--- Первые 5 строк NAV ---")
            print(nav_result.head(5))
        
        if inout_result is not None:
            print("\n--- Первые 5 строк InOut ---")
            print(inout_result.head(5))
        
        # Проверяем дубликаты
        if nav_result is not None:
            nav_duplicates = nav_result[nav_result.duplicated(subset=['Date'], keep=False)]
            print(f"\nДубликаты дат в NAV: {len(nav_duplicates)}")
        
        if inout_result is not None:
            inout_duplicates = inout_result[inout_result.duplicated(subset=['Date'], keep=False)]
            print(f"Дубликаты дат в InOut: {len(inout_duplicates)}")
            
    except Exception as e:
        print(f"❌ Ошибка при обработке файла: {e}")
        
else:
    print("❌ Файлы с ключевым словом 'Вознаграждение' не найдены")
    
    # Поищем в других типичных местах
    other_paths = [
        r'C:\Users\ytggf\Downloads',
        r'C:\Users\ytggf\Desktop',
        r'C:\Users\ytggf\OneDrive\Рабочий стол'
    ]
    
    print("\nПоиск в других местах...")
    for path in other_paths:
        if os.path.exists(path):
            search_pattern = os.path.join(path, '*Вознаграждение*.xls*')
            files = glob.glob(search_pattern, recursive=False)
            for file in files:
                print(f"Найден: {file}")
