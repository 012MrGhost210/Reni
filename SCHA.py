import pandas as pd
import os
import glob
import re
from pathlib import Path

def natural_sort_key(sheet_name):
    """Естественная сортировка листов"""
    return [int(text) if text.isdigit() else text.lower() 
            for text in re.split('([0-9]+)', sheet_name)]

def combine_data(data_dict):
    """Объединение данных по датам"""
    if not data_dict:
        return None
    sheets_list = sorted(data_dict.keys(), key=natural_sort_key)
    result = data_dict[sheets_list[0]]
    for sheet in sheets_list[1:]:
        result = pd.merge(result, data_dict[sheet], on='Date', how='outer')
    return result.sort_values('Date').drop_duplicates(subset=['Date'])

def process_sputnik():
    """Обработка Спутник (Вознаграждение)"""
    pattern = r'\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\25.Автоматизация\NAV for DI\**\*Вознаграждение*.xls*'
    files = glob.glob(pattern, recursive=True)
    
    if not files:
        return "❌ Спутник: файлы 'Вознаграждение' не найдены"
    
    try:
        excel_file = pd.ExcelFile(files[0])
        nav_data, inout_data = {}, {}
        
        for sheet in excel_file.sheet_names:
            if sheet == "ИТОГО":
                continue
            df = pd.read_excel(files[0], sheet_name=sheet)
            if 'Date' not in df.columns:
                continue
                
            df['Date'] = pd.to_datetime(df['Date']).dt.date
            
            if 'NAV' in df.columns:
                nav = df[['Date', 'NAV']].dropna()
                nav = nav[pd.to_numeric(nav['NAV'], errors='coerce').notna()]
                nav_data[sheet] = nav.groupby('Date').first().rename(columns={'NAV': sheet})
            
            if 'InOut' in df.columns:
                inout = df[['Date', 'InOut']].dropna()
                inout = inout[pd.to_numeric(inout['InOut'], errors='coerce').notna()]
                inout_data[sheet] = inout.groupby('Date').first().rename(columns={'InOut': sheet})
        
        output = r'\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\25.Автоматизация\NaVi\NaViСпутник_СЧА.xlsx'
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            if nav_data:
                combine_data(nav_data).to_excel(writer, sheet_name='NAV', index=False)
            if inout_data:
                combine_data(inout_data).to_excel(writer, sheet_name='InOut', index=False)
        return "✅ Спутник: успешно"
    except Exception as e:
        return f"❌ Спутник: {str(e)[:100]}"

def process_tkb():
    """Обработка ТКБ (Сводная РСА-СЧА)"""
    pattern = r'\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\25.Автоматизация\NAV for DI\**\*Сводная РСА-СЧА*.xlsx'
    files = glob.glob(pattern, recursive=True)
    
    if not files:
        return "❌ ТКБ: файлы 'Сводная РСА-СЧА' не найдены"
    
    try:
        excel_file = pd.ExcelFile(files[0])
        sheets = sorted(excel_file.sheet_names, key=natural_sort_key)
        scha_data, inout_data = {}, {}
        
        for sheet in sheets:
            df = pd.read_excel(files[0], sheet_name=sheet, skiprows=6, header=None)
            df = df.dropna(axis=1, how='all')
            
            cols = 7 if len(df.columns) == 7 else 6
            if cols == 7:
                df.columns = ['№', 'Date', 'Вводы', 'Выводы', 'РСА', 'СЧА', 'Пусто']
                df = df.drop(columns=['Пусто'])
            elif cols == 6:
                df.columns = ['№', 'Date', 'Вводы', 'Выводы', 'РСА', 'СЧА']
            else:
                continue
            
            df = df.dropna(subset=['Date'])
            df = df[~df['Date'].astype(str).str.contains('Суммарная|Количество|Средняя|№ п/п', na=False)]
            df['Date'] = pd.to_datetime(df['Date'], format='%d.%m.%Y', errors='coerce').dt.date
            df = df.dropna(subset=['Date'])
            
            if len(df) == 0:
                continue
            
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
        
        output = r'\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\25.Автоматизация\NaVi\NaViТКБ_СЧА.xlsx'
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            if scha_data:
                combine_data(scha_data).to_excel(writer, sheet_name='СЧА', index=False)
            if inout_data:
                combine_data(inout_data).to_excel(writer, sheet_name='InOut', index=False)
        return "✅ ТКБ: успешно"
    except Exception as e:
        return f"❌ ТКБ: {str(e)[:100]}"

def process_raif():
    """Обработка Райффайзен (Отчет по СЧА)"""
    pattern = r'\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\25.Автоматизация\NAV for DI\**\*Отчет по СЧА*.xlsx'
    files = glob.glob(pattern, recursive=True)
    
    if not files:
        return "❌ Райф: файлы 'Отчет по СЧА' не найдены"
    
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
            
            if len(df) == 0:
                continue
            
            # СЧА с обработкой формул
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
        
        output = r'\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\25.Автоматизация\NaVi\NaViРайф_СЧА.xlsx'
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            if scha_data:
                combine_data(scha_data).to_excel(writer, sheet_name='СЧА', index=False)
            if inout_data:
                combine_data(inout_data).to_excel(writer, sheet_name='InOut', index=False)
        return "✅ Райф: успешно"
    except Exception as e:
        return f"❌ Райф: {str(e)[:100]}"

def main():
    print("="*50)
    print("Запуск обработки компаний")
    print("="*50)
    
    results = [
        process_sputnik(),
        process_tkb(),
        process_raif()
    ]
    
    print("\n" + "="*50)
    print("ИТОГИ:")
    for res in results:
        print(res)

if __name__ == "__main__":
    main()
