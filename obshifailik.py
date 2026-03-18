import pandas as pd
import os
import glob
import re

def natural_sort_key(sheet_name):
    return [int(text) if text.isdigit() else text.lower() 
            for text in re.split('([0-9]+)', sheet_name)]

docs_path = r'\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\25.Автоматизация\NAV for DI'

def process_sputnik():
    """Обработка Спутник (файлы Вознаграждение)"""
    files = glob.glob(os.path.join(docs_path, '**', '*Вознаграждение*.xls*'), recursive=True)
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
    files = glob.glob(os.path.join(docs_path, '**', '*Сводная РСА-СЧА*.xlsx'), recursive=True)
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
    files = glob.glob(os.path.join(docs_path, '**', '*Отчет по СЧА*.xlsx'), recursive=True)
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

def main():
    print("="*40)
    print("ОБРАБОТКА КОМПАНИЙ")
    print("="*40)
    
    results = [
        process_sputnik(),
        process_tkb(),
        process_raif()
    ]
    
    print("\n" + "="*40)
    print("РЕЗУЛЬТАТЫ:")
    for res in results:
        print(res)

if __name__ == "__main__":
    main()
