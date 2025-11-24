import pandas as pd
import numpy as np

# Маппинг портфелей из твоего сообщения
portfolio_mapping = {
    '020611/1': 'ДУ «Спутник-УК» 020611/1 SPURZ 1',
    '020611/2': 'ДУ «Спутник-УК» 020611/2 SPURZ 2', 
    '020611/3': 'ДУ «Спутник-УК» 020611/3 SPURZ 3',
    '081121/1': 'ДУ «Спутник-УК» 081121/1 SPURZ 11',
    '081121/2': 'ДУ «Спутник-УК» 081121/2 SPURZ 12',
    '141111/1': 'ДУ «Спутник-УК» 141111/1 SPURZ 4',
    '190221/1': 'ДУ «Спутник-УК» 190221/1 SPURZ 10',
    '220223/1': 'ДУ «Спутник-УК» 220223/1 SPURZ 13',
    '220223/2': 'ДУ «Спутник-УК» 220223/2 SPURZ 14',
    '260716/1': 'ДУ «Спутник-УК» 260716/1 SPURZ 5',
    '271210/2': 'ДУ «Спутник-УК» 271210/2 SPURZ',
    '050925/1': 'ДУ «Спутник-УК» 050925/1 SPURZ 15'
}

def process_portfolio_data(file_path):
    """Обработка данных портфелей из Excel файла"""
    
    # Читаем файл
    df = pd.read_excel(file_path, header=4)
    
    # Фильтруем только строки с данными (убираем пустые)
    df = df[df['Портфель'].notna()]
    
    # Определяем числовые колонки для группировки
    numeric_columns = ['Стоимость', 'НКД,начисленные %', 'Дебеторская/ Кредиторская задолженности']
    
    # Конвертируем числовые колонки
    for col in numeric_columns:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    # Группируем по портфелю и суммируем числовые колонки
    grouped_df = df.groupby('Портфель')[numeric_columns].sum().reset_index()
    
    # Добавляем полное название портфеля из маппинга
    def get_full_portfolio_name(portfolio):
        # Ищем совпадение в маппинге
        for key, value in portfolio_mapping.items():
            if key in portfolio:
                return value
        return portfolio  # Если не нашли, оставляем оригинальное название
    
    grouped_df['Полное название портфеля'] = grouped_df['Портфель'].apply(get_full_portfolio_name)
    
    # Переупорядочиваем колонки
    result_df = grouped_df[['Портфель', 'Полное название портфеля'] + numeric_columns]
    
    return result_df

def merge_multiple_portfolio_files(file_paths, output_file):
    """Объединение нескольких файлов с портфелями"""
    
    all_data = []
    
    for file_path in file_paths:
        print(f"Обрабатываю файл: {file_path}")
        
        # Извлекаем дату из файла
        df_meta = pd.read_excel(file_path, header=None, nrows=2)
        date_str = df_meta.iloc[1, 0].replace('на дату ', '').split(' ')[0]
        
        # Обрабатываем данные портфелей
        portfolio_df = process_portfolio_data(file_path)
        
        # Добавляем дату отчета
        portfolio_df['Дата отчета'] = date_str
        
        all_data.append(portfolio_df)
    
    if all_data:
        # Объединяем все данные
        final_df = pd.concat(all_data, ignore_index=True)
        
        # Сохраняем результат
        final_df.to_excel(output_file, index=False)
        print(f"Файл сохранен: {output_file}")
        
        # Выводим сводку
        print(f"\nСводка:")
        print(f"Всего портфелей: {len(final_df)}")
        print(f"Диапазон дат: {final_df['Дата отчета'].min()} - {final_df['Дата отчета'].max()}")
        
        return final_df
    else:
        print("Нет данных для обработки")
        return None

# Альтернативная функция для более точного маппинга
def enhanced_portfolio_mapping(portfolio_name):
    """Улучшенный маппинг портфелей"""
    
    mapping_rules = [
        ('271210/2', 'ДУ «Спутник-УК» 271210/2 SPURZ'),
        ('020611/1 агресс', 'ДУ «Спутник-УК» 020611/1 SPURZ 1'),
        ('020611/2 консерв', 'ДУ «Спутник-УК» 020611/2 SPURZ 2'),
        ('020611/3 сбаланс', 'ДУ «Спутник-УК» 020611/3 SPURZ 3'),
        ('141111/1 агресс', 'ДУ «Спутник-УК» 141111/1 SPURZ 4'),
        ('260716/1 индивид', 'ДУ «Спутник-УК» 260716/1 SPURZ 5'),
        ('190221/1', 'ДУ «Спутник-УК» 190221/1 SPURZ 10'),
        ('081121/1', 'ДУ «Спутник-УК» 081121/1 SPURZ 11'),
        ('081121/2', 'ДУ «Спутник-УК» 081121/2 SPURZ 12'),
        ('220223/1', 'ДУ «Спутник-УК» 220223/1 SPURZ 13'),
        ('220223/2', 'ДУ «Спутник-УК» 220223/2 SPURZ 14'),
        ('050925/1', 'ДУ «Спутник-УК» 050925/1 SPURZ 15'),
    ]
    
    for key, value in mapping_rules:
        if key in portfolio_name:
            return value
    
    return portfolio_name

# Использование
if __name__ == "__main__":
    # Список файлов для обработки
    files_to_process = ["тест ккуу.xlsx"]  # Добавь нужные файлы
    
    # Обработка и объединение
    result = merge_multiple_portfolio_files(files_to_process, "обработанные_портфели.xlsx")
    
    if result is not None:
        print("\nПервые 5 строк результата:")
        print(result.head())
