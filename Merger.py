import pandas as pd
import os
from datetime import datetime
import glob

def parse_date_from_cell(a2_value):
    """Извлечь дату из ячейки A2"""
    try:
        # Пример: "на дату 1.10.2025 0:00:00"
        date_str = a2_value.replace('на дату ', '').split(' ')[0]
        return datetime.strptime(date_str, '%d.%m.%Y').date()
    except:
        return None

def read_excel_with_date(file_path):
    """Чтение Excel файла с извлечением даты"""
    try:
        # Читаем ячейку A2 для получения даты
        df_meta = pd.read_excel(file_path, header=None, nrows=2)
        report_date = parse_date_from_cell(df_meta.iloc[1, 0])
        
        # Читаем основную таблицу (начиная с строки 4, где 0-based indexing)
        df = pd.read_excel(file_path, header=4)
        
        # Добавляем столбец с датой отчета
        df['Дата отчета'] = report_date
        
        return df
    except Exception as e:
        print(f"Ошибка чтения файла {file_path}: {e}")
        return None

def merge_excel_files(folder_path, output_file):
    """Объединение всех Excel файлов в папке"""
    all_data = []
    
    # Ищем все Excel файлы в папке
    excel_files = glob.glob(os.path.join(folder_path, "*.xlsx"))
    
    for file_path in excel_files:
        print(f"Обрабатываю файл: {os.path.basename(file_path)}")
        df = read_excel_with_date(file_path)
        if df is not None and not df.empty:
            all_data.append(df)
    
    if all_data:
        # Объединяем все DataFrame
        merged_df = pd.concat(all_data, ignore_index=True)
        
        # Сохраняем результат
        merged_df.to_excel(output_file, index=False)
        print(f"Объединенный файл сохранен как: {output_file}")
        print(f"Объединено файлов: {len(all_data)}")
        print(f"Общее количество строк: {len(merged_df)}")
    else:
        print("Не найдено данных для объединения")

# Использование
folder_path = "путь_к_папке_с_файлами"  # Укажи путь к папке с файлами
output_file = "объединенный_отчет.xlsx"

merge_excel_files(folder_path, output_file)
