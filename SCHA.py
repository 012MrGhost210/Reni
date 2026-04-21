import pandas as pd
import os
import glob
import re
import shutil
from datetime import datetime
import win32com.client as win32

# --- НАСТРОЙКИ ДЛЯ ОБНОВЛЕНИЯ ---
source_root = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\diadoc_connector\Документооборот завершён"
destination_folder = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\25.Автоматизация\NAV for DI"

# Соответствие папок: (путь, что_ищем, короткое_имя_для_вывода)
TARGETS = [
    {
        "short_name": "Спутник",
        "path": "7744000951-АО -УК -СПУТНИК - УПРАВЛЕНИЕ КАПИТАЛОМ-",
        "search_pattern": "Вознаграждение"
    },
    {
        "short_name": "Райф",
        "path": "7702358512-ООО -УК Райффайзен-",
        "search_pattern": "Отчет по СЧА"
    },
    {
        "short_name": "ТКБ",
        "path": "7825489723-ТКБ Инвестмент Партнерс (АО)",
        "search_pattern": "Сводная РСА-СЧА"
    }
]

# Настройки email
EMAIL_RECIPIENTS = "Stepan.Koltsov@renlife.com; Ulyana.Pankratova@renlife.com; Aleksandra.Belousova@renlife.com"

# --- ФУНКЦИИ ДЛЯ ОБНОВЛЕНИЯ ---
def remove_text_in_braces(filename):
    """Удаляет из имени файла все, что в фигурных скобках."""
    new_name = re.sub(r'\{.*?\}', '', filename)
    new_name = re.sub(r'\s+', ' ', new_name).strip()
    return new_name

def is_today(date):
    """Проверяет, является ли дата сегодняшней."""
    today = datetime.now().date()
    return date.date() == today

def clear_folder(folder_path):
    """Очищает папку: удаляет все файлы и подпапки."""
    if os.path.exists(folder_path):
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print(f"  Ошибка при очистке {file_path}: {e}")

def process_folder(target, source_root, destination_folder):
    """Ищет САМЫЙ СВЕЖИЙ файл в конкретной подпапке, проверяет дату и копирует."""
    folder_path = os.path.join(source_root, target["path"])
    short_name = target["short_name"]
    search_pattern = target["search_pattern"]
    
    matching_files = []

    if not os.path.exists(folder_path):
        return f"{short_name}: папка не найдена"

    for filename in os.listdir(folder_path):
        if search_pattern.lower() in filename.lower():
            file_path = os.path.join(folder_path, filename)
            if os.path.isfile(file_path):
                mod_time = os.path.getmtime(file_path)
                mod_date = datetime.fromtimestamp(mod_time)
                matching_files.append((filename, mod_date, file_path))

    if not matching_files:
        return f"{short_name}: файлы не найдены"

    matching_files.sort(key=lambda x: x[1], reverse=True)
    latest_file, latest_date, latest_path = matching_files[0]

    if not is_today(latest_date):
        return f"{short_name}: самый свежий файл от {latest_date.date()}"

    new_filename = remove_text_in_braces(latest_file)
    destination_file = os.path.join(destination_folder, new_filename)

    base, ext = os.path.splitext(new_filename)
    counter = 1
    while os.path.exists(destination_file):
        destination_file = os.path.join(destination_folder, f"{base}_{counter}{ext}")
        counter += 1

    try:
        shutil.copy2(latest_path, destination_file)
        return f"{short_name}: скопирован"
    except Exception as e:
        return f"{short_name}: ошибка копирования"

# --- ФУНКЦИИ ДЛЯ ОСНОВНОЙ ОБРАБОТКИ ---
def natural_sort_key(sheet_name):
    return [int(text) if text.isdigit() else text.lower() 
            for text in re.split('([0-9]+)', sheet_name)]

def process_sputnik():
    """Обработка Спутник (файлы Вознаграждение)"""
    files = glob.glob(os.path.join(destination_folder, '**', '*Вознаграждение*.xls*'), recursive=True)
    if not files:
        return " Спутник: файлы не найдены"
    
    try:
        excel_file = pd.ExcelFile(files[0])
        nav_data, inout_data = {}, {}
        
        for sheet in excel_file.sheet_names:
            if sheet == "ИТОГО":
                continue
            
            df = pd.read_excel(files[0], sheet_name=sheet)
            if 'Date' not in df.columns:
                continue
            
            if 'NAV' in df.columns:
                nav = df[['Date', 'NAV']].copy()
                nav['Date'] = pd.to_datetime(nav['Date']).dt.date
                nav = nav.dropna(subset=['NAV'])
                nav = nav[pd.to_numeric(nav['NAV'], errors='coerce').notna()]
                nav = nav.groupby('Date').first().reset_index().rename(columns={'NAV': sheet})
                nav_data[sheet] = nav
            
            if 'InOut' in df.columns:
                inout = df[['Date', 'InOut']].copy()
                inout['Date'] = pd.to_datetime(inout['Date']).dt.date
                inout = inout.dropna(subset=['InOut'])
                inout = inout[pd.to_numeric(inout['InOut'], errors='coerce').notna()]
                inout = inout.groupby('Date').first().reset_index().rename(columns={'InOut': sheet})
                inout_data[sheet] = inout
        
        if nav_data:
            nav_result = nav_data[list(nav_data.keys())[0]]
            for sheet in list(nav_data.keys())[1:]:
                nav_result = pd.merge(nav_result, nav_data[sheet], on='Date', how='outer')
            nav_result = nav_result.sort_values('Date').drop_duplicates(subset=['Date'])
            nav_result['Date'] = pd.to_datetime(nav_result['Date'])
        else:
            nav_result = None
        
        if inout_data:
            inout_result = inout_data[list(inout_data.keys())[0]]
            for sheet in list(inout_data.keys())[1:]:
                inout_result = pd.merge(inout_result, inout_data[sheet], on='Date', how='outer')
            inout_result = inout_result.sort_values('Date').drop_duplicates(subset=['Date'])
            inout_result['Date'] = pd.to_datetime(inout_result['Date'])
        else:
            inout_result = None
        
        output = r'\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\25.Автоматизация\NaVi\NaViСпутник_СЧА.xlsx'
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            if nav_result is not None:
                nav_result.to_excel(writer, sheet_name='NAV', index=False)
            if inout_result is not None:
                inout_result.to_excel(writer, sheet_name='InOut', index=False)
        
        return " Спутник: успешно"
    except Exception as e:
        return f" Спутник: {str(e)[:150]}"

def process_tkb():
    """Обработка ТКБ (Сводная РСА-СЧА)"""
    files = glob.glob(os.path.join(destination_folder, '**', '*Сводная РСА-СЧА*.xlsx'), recursive=True)
    if not files:
        return " ТКБ: файлы не найдены"
    
    try:
        excel_file = pd.ExcelFile(files[0])
        sheets = sorted(excel_file.sheet_names, key=natural_sort_key)
        scha_data, inout_data = {}, {}
        
        for sheet in sheets:
            df = pd.read_excel(files[0], sheet_name=sheet, skiprows=6, header=None)
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
                scha = df[['Date', 'СЧА']].copy()
                scha['СЧА'] = pd.to_numeric(scha['СЧА'], errors='coerce')
                scha_data[sheet] = scha.dropna().rename(columns={'СЧА': sheet})
                
                inout = df[['Date', 'Вводы', 'Выводы']].copy()
                inout['Вводы'] = pd.to_numeric(inout['Вводы'], errors='coerce').fillna(0)
                inout['Выводы'] = pd.to_numeric(inout['Выводы'], errors='coerce').fillna(0)
                inout[sheet] = inout['Вводы'] - inout['Выводы']
                inout_data[sheet] = inout[['Date', sheet]].dropna()
        
        if scha_data:
            scha_result = scha_data[sorted(scha_data.keys(), key=natural_sort_key)[0]]
            for sheet in sorted(scha_data.keys(), key=natural_sort_key)[1:]:
                scha_result = pd.merge(scha_result, scha_data[sheet], on='Date', how='outer')
            scha_result = scha_result.sort_values('Date').drop_duplicates(subset=['Date'])
            scha_result['Date'] = pd.to_datetime(scha_result['Date'])
        else:
            scha_result = None
        
        if inout_data:
            inout_result = inout_data[sorted(inout_data.keys(), key=natural_sort_key)[0]]
            for sheet in sorted(inout_data.keys(), key=natural_sort_key)[1:]:
                inout_result = pd.merge(inout_result, inout_data[sheet], on='Date', how='outer')
            inout_result = inout_result.sort_values('Date').drop_duplicates(subset=['Date'])
            inout_result['Date'] = pd.to_datetime(inout_result['Date'])
        else:
            inout_result = None
        
        output = r'\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\25.Автоматизация\NaVi\NaViТКБ_СЧА.xlsx'
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            if scha_result is not None:
                scha_result.to_excel(writer, sheet_name='СЧА', index=False)
            if inout_result is not None:
                inout_result.to_excel(writer, sheet_name='InOut', index=False)
        
        return " ТКБ: успешно"
    except Exception as e:
        return f" ТКБ: {str(e)[:150]}"

def process_raif():
    """Обработка Райффайзен (Отчет по СЧА)"""
    files = glob.glob(os.path.join(destination_folder, '**', '*Отчет по СЧА*.xlsx'), recursive=True)
    if not files:
        return " Райф: файлы не найдены"
    
    try:
        excel_file = pd.ExcelFile(files[0])
        sheets = sorted(excel_file.sheet_names, key=natural_sort_key)
        scha_data, inout_data = {}, {}
        
        for sheet in sheets:
            df = pd.read_excel(files[0], sheet_name=sheet, skiprows=6, header=None)
            df = df.dropna(axis=1, how='all')
            
            if len(df.columns) != 5:
                continue
            df.columns = ['№', 'Date', 'Вводы', 'Выводы', 'СЧА']
            
            df = df.dropna(subset=['Date'])
            df = df[~df['Date'].astype(str).str.contains('Суммарная|Количество|Средняя|№ п/п', na=False)]
            df['Date'] = pd.to_datetime(df['Date'], format='%d.%m.%Y', errors='coerce').dt.date
            df = df.dropna(subset=['Date'])
            
            if len(df) > 0:
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
                
                inout = df[['Date', 'Вводы', 'Выводы']].copy()
                inout['Вводы'] = pd.to_numeric(inout['Вводы'], errors='coerce').fillna(0)
                inout['Выводы'] = pd.to_numeric(inout['Выводы'], errors='coerce').fillna(0)
                inout[sheet] = inout['Вводы'] - inout['Выводы']
                inout_data[sheet] = inout[['Date', sheet]].dropna()
        
        if scha_data:
            scha_result = scha_data[sorted(scha_data.keys(), key=natural_sort_key)[0]]
            for sheet in sorted(scha_data.keys(), key=natural_sort_key)[1:]:
                scha_result = pd.merge(scha_result, scha_data[sheet], on='Date', how='outer')
            scha_result = scha_result.sort_values('Date').drop_duplicates(subset=['Date'])
            scha_result['Date'] = pd.to_datetime(scha_result['Date'])
        else:
            scha_result = None
        
        if inout_data:
            inout_result = inout_data[sorted(inout_data.keys(), key=natural_sort_key)[0]]
            for sheet in sorted(inout_data.keys(), key=natural_sort_key)[1:]:
                inout_result = pd.merge(inout_result, inout_data[sheet], on='Date', how='outer')
            inout_result = inout_result.sort_values('Date').drop_duplicates(subset=['Date'])
            inout_result['Date'] = pd.to_datetime(inout_result['Date'])
        else:
            inout_result = None
        
        output = r'\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\25.Автоматизация\NaVi\NaViРайф_СЧА.xlsx'
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            if scha_result is not None:
                scha_result.to_excel(writer, sheet_name='СЧА', index=False)
            if inout_result is not None:
                inout_result.to_excel(writer, sheet_name='InOut', index=False)
        
        return " Райф: успешно"
    except Exception as e:
        return f" Райф: {str(e)[:150]}"

def send_email(update_results, processing_results):
    """
    Отправляет email с результатами выполнения
    """
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    
    mail.To = EMAIL_RECIPIENTS
    mail.Subject = f"Отчет по обработке СЧА от {datetime.now().strftime('%d.%m.%Y')}"
    
    # Формируем тело письма
    body = f"Дата и время выполнения: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n\n"
    
    # Этап 1: Обновление файлов
    body += "ЭТАП 1: Обновление исходных файлов\n"
    body += "-" * 40 + "\n"
    
    all_ok = True
    failed_updates = []
    
    for res in update_results:
        if "скопирован" in res:
            body += f" {res}\n"
        else:
            body += f" {res}\n"
            all_ok = False
            failed_updates.append(res.split(":")[0])
    
    if all_ok:
        body += "\n Все отчеты СЧА были обработаны (скопированы свежие файлы)\n"
    else:
        body += f"\n Не удалось найти отчеты: {', '.join(failed_updates)}\n"
    
    body += "\n"
    
    # Этап 2: Обработка компаний
    body += "ЭТАП 2: Обработка компаний\n"
    body += "-" * 40 + "\n"
    
    processing_error = False
    for res in processing_results:
        if "успешно" in res:
            body += f" {res}\n"
        else:
            body += f" {res}\n"
            processing_error = True
    
    if processing_error:
        body += "\n ВНИМАНИЕ: Второй этап завершён с ошибкой. Просьба проверить отчеты СЧА вручную.\n"
    
    mail.Body = body
    
    try:
        mail.Send()
        print("\n Email успешно отправлен")
    except Exception as e:
        print(f"\n Ошибка при отправке email: {e}")

# --- ОСНОВНАЯ ПРОГРАММА ---
def main():
    print("="*50)
    print("ЭТАП 1: ОБНОВЛЕНИЕ ИСХОДНЫХ ФАЙЛОВ")
    print("="*50)
    
    # Очищаем папку назначения
    print("Очистка папки NAV for DI...")
    clear_folder(destination_folder)
    
    # Создаем папку назначения, если её нет
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)
    
    print("\nКопирование свежих файлов:")
    update_results = []
    for target in TARGETS:
        result = process_folder(target, source_root, destination_folder)
        print(f"  {result}")
        update_results.append(result)
    
    print("\n" + "="*50)
    print("ЭТАП 2: ОБРАБОТКА КОМПАНИЙ")
    print("="*50)
    
    processing_results = [
        process_sputnik(),
        process_tkb(),
        process_raif()
    ]
    
    for res in processing_results:
        print(f"  {res}")
    
    print("\n" + "="*50)
    print("ЭТАП 3: ОТПРАВКА EMAIL")
    print("="*50)
    
    send_email(update_results, processing_results)
    
    print("\n" + "="*50)
    print(" ВСЕ ЭТАПЫ ЗАВЕРШЕНЫ")
    print("="*50)

if __name__ == "__main__":
    main()
