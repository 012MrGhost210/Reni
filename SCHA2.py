import pandas as pd
import os
import glob
import re
import shutil
from datetime import datetime
import win32com.client as win32

# --- НАСТРОЙКИ ---
# Путь к папке с уже загруженными отчетами
source_folder = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\25.Автоматизация\NAV for DI"  # Здесь уже лежат готовые файлы

# Папка для результатов
output_folder = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\25.Автоматизация\NaVi"

# Настройки email
EMAIL_RECIPIENTS = "Stepan.Koltsov@renlife.com; Ulyana.Pankratova@renlife.com; Mihail.Uvarkin@renlife.com"

# --- ФУНКЦИИ ДЛЯ ОБРАБОТКИ ---
def natural_sort_key(sheet_name):
    """Натуральная сортировка для имен листов"""
    return [int(text) if text.isdigit() else text.lower() 
            for text in re.split('([0-9]+)', sheet_name)]

def find_latest_file(folder_path, pattern):
    """Находит самый свежий файл по шаблону в папке"""
    files = glob.glob(os.path.join(folder_path, pattern), recursive=False)
    files = [f for f in files if not os.path.basename(f).startswith('~$')]
    
    if not files:
        return None
    
    # Сортируем по дате модификации (самый свежий последний)
    files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    return files[0]

def process_sputnik(folder_path, output_folder):
    """Обработка Спутник (файлы Вознаграждение)"""
    try:
        # Ищем файл
        file_pattern = '*Вознаграждение*.xls*'
        file_path = find_latest_file(folder_path, file_pattern)
        
        if not file_path:
            return "Спутник: файлы не найдены"
        
        print(f"  Обработка файла: {os.path.basename(file_path)}")
        
        # Определяем движок для чтения
        engine = 'openpyxl' if file_path.endswith('.xlsx') else None
        
        with pd.ExcelFile(file_path, engine=engine) as excel_file:
            nav_data, inout_data = {}, {}
            
            for sheet in excel_file.sheet_names:
                if sheet == "ИТОГО":
                    continue
                
                df = pd.read_excel(file_path, sheet_name=sheet, engine=engine)
                
                if 'Date' not in df.columns:
                    continue
                
                # Обработка NAV
                if 'NAV' in df.columns:
                    nav = df[['Date', 'NAV']].copy()
                    nav['Date'] = pd.to_datetime(nav['Date']).dt.date
                    nav = nav.dropna(subset=['NAV'])
                    nav = nav[pd.to_numeric(nav['NAV'], errors='coerce').notna()]
                    nav = nav.groupby('Date').first().reset_index().rename(columns={'NAV': sheet})
                    nav_data[sheet] = nav
                
                # Обработка InOut
                if 'InOut' in df.columns:
                    inout = df[['Date', 'InOut']].copy()
                    inout['Date'] = pd.to_datetime(inout['Date']).dt.date
                    inout = inout.dropna(subset=['InOut'])
                    inout = inout[pd.to_numeric(inout['InOut'], errors='coerce').notna()]
                    inout = inout.groupby('Date').first().reset_index().rename(columns={'InOut': sheet})
                    inout_data[sheet] = inout
            
            # Объединение данных NAV
            if nav_data:
                nav_result = nav_data[list(nav_data.keys())[0]]
                for sheet in list(nav_data.keys())[1:]:
                    nav_result = pd.merge(nav_result, nav_data[sheet], on='Date', how='outer')
                nav_result = nav_result.sort_values('Date').drop_duplicates(subset=['Date'])
                nav_result['Date'] = pd.to_datetime(nav_result['Date'])
            else:
                nav_result = None
            
            # Объединение данных InOut
            if inout_data:
                inout_result = inout_data[list(inout_data.keys())[0]]
                for sheet in list(inout_data.keys())[1:]:
                    inout_result = pd.merge(inout_result, inout_data[sheet], on='Date', how='outer')
                inout_result = inout_result.sort_values('Date').drop_duplicates(subset=['Date'])
                inout_result['Date'] = pd.to_datetime(inout_result['Date'])
            else:
                inout_result = None
        
        # Сохраняем результат
        output_path = os.path.join(output_folder, 'NaViСпутник_СЧА.xlsx')
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            if nav_result is not None and not nav_result.empty:
                nav_result.to_excel(writer, sheet_name='NAV', index=False)
            if inout_result is not None and not inout_result.empty:
                inout_result.to_excel(writer, sheet_name='InOut', index=False)
        
        return "Спутник: успешно"
        
    except Exception as e:
        return f"Спутник: {str(e)[:150]}"

def process_tkb(folder_path, output_folder):
    """Обработка ТКБ (Сводная РСА-СЧА)"""
    try:
        file_pattern = '*Сводная РСА-СЧА*.xlsx'
        file_path = find_latest_file(folder_path, file_pattern)
        
        if not file_path:
            return "ТКБ: файлы не найдены"
        
        print(f"  Обработка файла: {os.path.basename(file_path)}")
        
        with pd.ExcelFile(file_path, engine='openpyxl') as excel_file:
            sheets = sorted(excel_file.sheet_names, key=natural_sort_key)
            scha_data, inout_data = {}, {}
            
            for sheet in sheets:
                df = pd.read_excel(file_path, sheet_name=sheet, skiprows=6, header=None, engine='openpyxl')
                df = df.dropna(axis=1, how='all')
                
                if len(df.columns) == 7:
                    df.columns = ['№', 'Date', 'Вводы', 'Выводы', 'РСА', 'СЧА', 'Пусто']
                    df = df.drop(columns=['Пусто'])
                elif len(df.columns) == 6:
                    df.columns = ['№', 'Date', 'Вводы', 'Выводы', 'РСА', 'СЧА']
                else:
                    continue
                
                df = df.dropna(subset=['Date'])
                df = df[~df['Date'].astype(str).str.contains('Суммарная|Количество|Средняя|№ п/п', na=False)]
                df['Date'] = pd.to_datetime(df['Date'], format='%d.%m.%Y', errors='coerce').dt.date
                df = df.dropna(subset=['Date'])
                
                if len(df) > 0:
                    # СЧА
                    scha = df[['Date', 'СЧА']].copy()
                    scha['СЧА'] = pd.to_numeric(scha['СЧА'], errors='coerce')
                    scha_data[sheet] = scha.dropna().rename(columns={'СЧА': sheet})
                    
                    # InOut
                    inout = df[['Date', 'Вводы', 'Выводы']].copy()
                    inout['Вводы'] = pd.to_numeric(inout['Вводы'], errors='coerce').fillna(0)
                    inout['Выводы'] = pd.to_numeric(inout['Выводы'], errors='coerce').fillna(0)
                    inout[sheet] = inout['Вводы'] - inout['Выводы']
                    inout_data[sheet] = inout[['Date', sheet]].dropna()
            
            # Объединение данных
            if scha_data:
                sorted_keys = sorted(scha_data.keys(), key=natural_sort_key)
                scha_result = scha_data[sorted_keys[0]]
                for sheet in sorted_keys[1:]:
                    scha_result = pd.merge(scha_result, scha_data[sheet], on='Date', how='outer')
                scha_result = scha_result.sort_values('Date').drop_duplicates(subset=['Date'])
                scha_result['Date'] = pd.to_datetime(scha_result['Date'])
            else:
                scha_result = None
            
            if inout_data:
                sorted_keys = sorted(inout_data.keys(), key=natural_sort_key)
                inout_result = inout_data[sorted_keys[0]]
                for sheet in sorted_keys[1:]:
                    inout_result = pd.merge(inout_result, inout_data[sheet], on='Date', how='outer')
                inout_result = inout_result.sort_values('Date').drop_duplicates(subset=['Date'])
                inout_result['Date'] = pd.to_datetime(inout_result['Date'])
            else:
                inout_result = None
        
        output_path = os.path.join(output_folder, 'NaViТКБ_СЧА.xlsx')
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            if scha_result is not None and not scha_result.empty:
                scha_result.to_excel(writer, sheet_name='СЧА', index=False)
            if inout_result is not None and not inout_result.empty:
                inout_result.to_excel(writer, sheet_name='InOut', index=False)
        
        return "ТКБ: успешно"
        
    except Exception as e:
        return f"ТКБ: {str(e)[:150]}"

def process_raif(folder_path, output_folder):
    """Обработка Райффайзен (Отчет по СЧА)"""
    try:
        file_pattern = '*Отчет по СЧА*.xlsx'
        file_path = find_latest_file(folder_path, file_pattern)
        
        if not file_path:
            return "Райф: файлы не найдены"
        
        print(f"  Обработка файла: {os.path.basename(file_path)}")
        
        with pd.ExcelFile(file_path, engine='openpyxl') as excel_file:
            sheets = sorted(excel_file.sheet_names, key=natural_sort_key)
            scha_data, inout_data = {}, {}
            
            for sheet in sheets:
                df = pd.read_excel(file_path, sheet_name=sheet, skiprows=6, header=None, engine='openpyxl')
                df = df.dropna(axis=1, how='all')
                
                if len(df.columns) != 5:
                    continue
                df.columns = ['№', 'Date', 'Вводы', 'Выводы', 'СЧА']
                
                df = df.dropna(subset=['Date'])
                df = df[~df['Date'].astype(str).str.contains('Суммарная|Количество|Средняя|№ п/п', na=False)]
                df['Date'] = pd.to_datetime(df['Date'], format='%d.%m.%Y', errors='coerce').dt.date
                df = df.dropna(subset=['Date'])
                
                if len(df) > 0:
                    # СЧА с заполнением пропусков
                    scha = df[['Date', 'СЧА']].copy()
                    scha['СЧА'] = pd.to_numeric(scha['СЧА'], errors='coerce')
                    
                    if scha['СЧА'].isna().any():
                        first_valid = scha['СЧА'].first_valid_index()
                        if first_valid is not None:
                            base = scha.loc[first_valid, 'СЧА']
                            base_date = pd.to_datetime(scha.loc[first_valid, 'Date'])
                            for idx in scha.index:
                                if pd.isna(scha.loc[idx, 'СЧА']):
                                    days = (pd.to_datetime(scha.loc[idx, 'Date']) - base_date).days
                                    scha.loc[idx, 'СЧА'] = base + days
                    
                    scha_data[sheet] = scha.dropna().rename(columns={'СЧА': sheet})
                    
                    # InOut
                    inout = df[['Date', 'Вводы', 'Выводы']].copy()
                    inout['Вводы'] = pd.to_numeric(inout['Вводы'], errors='coerce').fillna(0)
                    inout['Выводы'] = pd.to_numeric(inout['Выводы'], errors='coerce').fillna(0)
                    inout[sheet] = inout['Вводы'] - inout['Выводы']
                    inout_data[sheet] = inout[['Date', sheet]].dropna()
            
            # Объединение данных
            if scha_data:
                sorted_keys = sorted(scha_data.keys(), key=natural_sort_key)
                scha_result = scha_data[sorted_keys[0]]
                for sheet in sorted_keys[1:]:
                    scha_result = pd.merge(scha_result, scha_data[sheet], on='Date', how='outer')
                scha_result = scha_result.sort_values('Date').drop_duplicates(subset=['Date'])
                scha_result['Date'] = pd.to_datetime(scha_result['Date'])
            else:
                scha_result = None
            
            if inout_data:
                sorted_keys = sorted(inout_data.keys(), key=natural_sort_key)
                inout_result = inout_data[sorted_keys[0]]
                for sheet in sorted_keys[1:]:
                    inout_result = pd.merge(inout_result, inout_data[sheet], on='Date', how='outer')
                inout_result = inout_result.sort_values('Date').drop_duplicates(subset=['Date'])
                inout_result['Date'] = pd.to_datetime(inout_result['Date'])
            else:
                inout_result = None
        
        output_path = os.path.join(output_folder, 'NaViРайф_СЧА.xlsx')
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            if scha_result is not None and not scha_result.empty:
                scha_result.to_excel(writer, sheet_name='СЧА', index=False)
            if inout_result is not None and not inout_result.empty:
                inout_result.to_excel(writer, sheet_name='InOut', index=False)
        
        return "Райф: успешно"
        
    except Exception as e:
        return f"Райф: {str(e)[:150]}"

def process_first(folder_path, output_folder):
    """Обработка УК Первая (Сводная СЧА)"""
    try:
        file_pattern = '*Сводная СЧА*.xlsx'
        file_path = find_latest_file(folder_path, file_pattern)
        
        if not file_path:
            return "Первая: файлы не найдены"
        
        print(f"  Обработка файла: {os.path.basename(file_path)}")
        
        with pd.ExcelFile(file_path, engine='openpyxl') as excel_file:
            sheets = sorted(excel_file.sheet_names, key=natural_sort_key)
            scha_data = {}
            
            for sheet in sheets:
                if sheet in ["ИТОГО", "Total", "Сводка"]:
                    continue
                
                df = pd.read_excel(file_path, sheet_name=sheet, header=None, engine='openpyxl')
                
                # Поиск заголовков
                header_row = None
                data_start_row = None
                
                for idx, row in df.iterrows():
                    row_str = ' '.join(str(v) for v in row.values if pd.notna(v))
                    if 'п/п' in row_str and 'День' in row_str and 'СЧА' in row_str:
                        header_row = idx
                        data_start_row = idx + 1
                        break
                
                if header_row is None:
                    continue
                
                headers = df.iloc[header_row].values
                col_map = {}
                for i, h in enumerate(headers):
                    if pd.isna(h):
                        continue
                    h_str = str(h).strip()
                    
                    if h_str == 'День':
                        col_map['date'] = i
                    elif h_str == 'СЧА Методика':
                        col_map['scha_method'] = i
                    elif h_str == 'СЧА Баланс':
                        col_map['scha_balance'] = i
                    elif h_str == 'СЧА из П2':
                        col_map['scha_p2'] = i
                
                if 'date' not in col_map:
                    continue
                
                # Сбор данных
                data_rows = []
                for idx in range(data_start_row, len(df)):
                    row = df.iloc[idx]
                    
                    if all(pd.isna(v) for v in row.values):
                        break
                    
                    date_val = row[col_map['date']] if col_map['date'] < len(row) else None
                    if pd.isna(date_val):
                        continue
                    
                    try:
                        if isinstance(date_val, (int, float)) and date_val > 40000:
                            date_obj = pd.Timestamp.fromordinal(int(date_val) - 693594).date()
                        else:
                            date_obj = pd.to_datetime(date_val, format='%d.%m.%Y', errors='coerce').date()
                        
                        if date_obj is None or pd.isna(date_obj):
                            continue
                    except:
                        continue
                    
                    # Поиск значения СЧА (приоритет: Методика > Баланс > П2)
                    scha_value = None
                    if 'scha_method' in col_map and col_map['scha_method'] < len(row):
                        val = row[col_map['scha_method']]
                        if pd.notna(val) and isinstance(val, (int, float)):
                            scha_value = val
                    
                    if scha_value is None and 'scha_balance' in col_map and col_map['scha_balance'] < len(row):
                        val = row[col_map['scha_balance']]
                        if pd.notna(val) and isinstance(val, (int, float)):
                            scha_value = val
                    
                    if scha_value is None and 'scha_p2' in col_map and col_map['scha_p2'] < len(row):
                        val = row[col_map['scha_p2']]
                        if pd.notna(val) and isinstance(val, (int, float)):
                            scha_value = val
                    
                    if scha_value is not None:
                        data_rows.append({
                            'Date': date_obj,
                            sheet: scha_value
                        })
                
                if data_rows:
                    df_sheet = pd.DataFrame(data_rows)
                    scha_data[sheet] = df_sheet
            
            if not scha_data:
                return "Первая: нет данных СЧА ни в одном листе"
            
            # Объединение данных
            sorted_keys = sorted(scha_data.keys(), key=natural_sort_key)
            scha_result = scha_data[sorted_keys[0]]
            for sheet in sorted_keys[1:]:
                scha_result = pd.merge(scha_result, scha_data[sheet], on='Date', how='outer')
            
            scha_result = scha_result.sort_values('Date').drop_duplicates(subset=['Date'])
            scha_result['Date'] = pd.to_datetime(scha_result['Date'])
        
        output_path = os.path.join(output_folder, 'NaViПервая_СЧА.xlsx')
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            scha_result.to_excel(writer, sheet_name='СЧА', index=False)
        
        return "Первая: успешно"
        
    except Exception as e:
        return f"Первая: {str(e)[:150]}"

def send_email(processing_results):
    """
    Отправляет email с результатами выполнения
    """
    try:
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        
        mail.To = EMAIL_RECIPIENTS
        mail.Subject = f"Отчет по обработке СЧА от {datetime.now().strftime('%d.%m.%Y')}"
        
        body = f"Дата и время выполнения: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n\n"
        body += "ОБРАБОТКА КОМПАНИЙ\n"
        body += "-" * 40 + "\n"
        
        processing_error = False
        for res in processing_results:
            if "успешно" in res:
                body += f" ✓ {res}\n"
            else:
                body += f" ✗ {res}\n"
                processing_error = True
        
        if processing_error:
            body += "\nВНИМАНИЕ: Обработка завершена с ошибкой. Просьба проверить отчеты СЧА вручную.\n"
        else:
            body += "\nВсе отчеты СЧА успешно обработаны!\n"
        
        mail.Body = body
        
        mail.Send()
        print("\n Email успешно отправлен")
        return True
    except Exception as e:
        print(f"\n Ошибка при отправке email: {e}")
        return False

def check_folder(folder_path):
    """Проверяет наличие файлов в папке"""
    if not os.path.exists(folder_path):
        print(f"❌ Папка не существует: {folder_path}")
        return False
    
    files = os.listdir(folder_path)
    if not files:
        print(f"⚠️ Папка пуста: {folder_path}")
        return False
    
    print(f"✅ Папка найдена, файлов: {len(files)}")
    return True

# --- ОСНОВНАЯ ПРОГРАММА ---
def main():
    print("=" * 60)
    print("ОБРАБОТКА ОТЧЕТОВ СЧА (без копирования)")
    print("=" * 60)
    print()
    
    # Проверка папки с исходными файлами
    print(f"📁 Проверка папки: {source_folder}")
    if not check_folder(source_folder):
        print("\n❌ Работа прервана: папка с файлами не найдена или пуста")
        return
    
    print("\n" + "=" * 60)
    print("НАЧАЛО ОБРАБОТКИ")
    print("=" * 60)
    print()
    
    processing_results = []
    
    # Обработка каждой компании
    print("📊 Обработка Спутник...")
    result = process_sputnik(source_folder, output_folder)
    print(f"  {result}")
    processing_results.append(result)
    
    print("📊 Обработка ТКБ...")
    result = process_tkb(source_folder, output_folder)
    print(f"  {result}")
    processing_results.append(result)
    
    print("📊 Обработка Райффайзен...")
    result = process_raif(source_folder, output_folder)
    print(f"  {result}")
    processing_results.append(result)
    
    print("📊 Обработка Первая...")
    result = process_first(source_folder, output_folder)
    print(f"  {result}")
    processing_results.append(result)
    
    print("\n" + "=" * 60)
    print("ОТПРАВКА ОТЧЕТА ПО EMAIL")
    print("=" * 60)
    
    send_email(processing_results)
    
    print("\n" + "=" * 60)
    print("✅ ВСЕ ЭТАПЫ ЗАВЕРШЕНЫ")
    print("=" * 60)

if __name__ == "__main__":
    main()
