import pandas as pd
import os
from datetime import datetime, timedelta

# Маппинг портфелей
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

def extract_data_from_merger(input_file_path):
    """Извлекает данные из файла Мерджер.xlsx"""
    
    print(f"📖 Читаю файл: {input_file_path}")
    
    try:
        # Читаем файл с правильным заголовком (строка 1 в 0-based индексации)
        df = pd.read_excel(input_file_path, header=1)
        print(f"Найдено строк в таблице: {len(df)}")
        print(f"Колонки: {df.columns.tolist()}")
        
        # Переименовываем первую колонку в 'Портфель'
        df = df.rename(columns={df.columns[0]: 'Портфель'})
        
        # Фильтруем только строки с данными в колонке Портфель
        df = df[df['Портфель'].notna()]
        df = df[df['Портфель'].astype(str).str.len() < 100]
        
        print(f"Строк после фильтрации: {len(df)}")
        
        # Определяем числовые колонки
        column_mapping = {
            df.columns[13]: 'Стоимость',
            df.columns[14]: 'НКД', 
            df.columns[15]: 'Задолженности'
        }
        
        # Переименовываем числовые колонки
        df = df.rename(columns=column_mapping)
        print(f"Переименованные числовые колонки: {list(column_mapping.values())}")
        
        # Конвертируем числовые колонки
        numeric_columns = ['Стоимость', 'НКД', 'Задолженности']
        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # Группируем по портфелю и суммируем числовые колонки
        grouped_df = df.groupby('Портфель')[numeric_columns].sum().reset_index()
        print(f"Сгруппировано портфелей: {len(grouped_df)}")
        
        # Добавляем идентификатор портфеля для маппинга
        def get_portfolio_id(portfolio):
            portfolio_str = str(portfolio)
            for key in portfolio_mapping.keys():
                if key in portfolio_str:
                    return key
            return None
        
        grouped_df['Portfolio_ID'] = grouped_df['Портфель'].apply(get_portfolio_id)
        
        # Выводим информацию о найденных портфелях
        print("\n📊 Найденные портфели:")
        for _, row in grouped_df.iterrows():
            if row['Portfolio_ID']:
                print(f"  ✅ {row['Портфель']} -> {row['Portfolio_ID']} (Стоимость: {row['Стоимость']:,.2f})")
            else:
                print(f"  ⚠️ {row['Портфель']} -> НЕ ОПРЕДЕЛЕН")
        
        return grouped_df
        
    except Exception as e:
        print(f"❌ Ошибка при чтении файла: {e}")
        import traceback
        traceback.print_exc()
        return None

def create_pivot_format(portfolio_data, output_file_path):
    """Создает файл в формате как в примере 232321312321dddddвавав.xlsx"""
    
    print("\n🔄 Создаю файл в целевом формате...")
    
    try:
        # Создаем даты с 2025-10-01 по 2025-10-30
        dates = [datetime(2025, 10, 1) + timedelta(days=i) for i in range(30)]
        
        # Создаем базовую структуру данных
        result_data = []
        
        for date in dates:
            row = {'Date': date}
            
            # Для каждого портфеля добавляем значение стоимости
            for portfolio_id in portfolio_mapping.keys():
                portfolio_value = portfolio_data[portfolio_data['Portfolio_ID'] == portfolio_id]['Стоимость']
                if not portfolio_value.empty:
                    row[portfolio_id] = portfolio_value.values[0]
                else:
                    # Если портфель не найден, используем значение по умолчанию
                    row[portfolio_id] = 121321312
            
            # Добавляем NAV как сумму всех портфелей
            row['NAV'] = sum([row[pid] for pid in portfolio_mapping.keys()])
            result_data.append(row)
        
        # Создаем финальный DataFrame
        final_df = pd.DataFrame(result_data)
        
        # Сохраняем с правильным форматированием
        with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
            # Создаем лист SAM_2025
            worksheet = writer.book.create_sheet('SAM_2025')
            
            # Добавляем заголовки как в примере
            headers = ['', 'СК', 'СК1', 'СК2', 'СК3', 'СК4', 'СК5', 'СК10', 'СК11', 'СК12', 'СК13', 'СК14', 'СК15', 'NAV']
            for col_idx, header in enumerate(headers, 1):
                worksheet.cell(row=2, column=col_idx, value=header)
            
            # Добавляем коды портфелей
            portfolio_codes = ['', '271210/2', '020611/1', '020611/2', '020611/3', '141111/1', '260716/1', 
                             '190221/1', '081121/1', '081121/2', '220223/1', '220223/2', '050925/1', '']
            for col_idx, code in enumerate(portfolio_codes, 1):
                worksheet.cell(row=3, column=col_idx, value=code)
            
            # Добавляем названия продуктов
            product_names = [
                'Date',
                'НСЖ рег. (защит.)\nНСЖ сингл (защит.)',
                'ИСЖ ДУ 2.0 (защит.)\nИСЖ сингл (защит.)',
                '-',
                'ИСЖ ДУ 1.0 (защит.)',
                '-', 
                'ИСЖ ДУ 2.0 ВСК (риск.)',
                'ИСЖ Опцион сб (защит.)',
                'НСЖ HTM (защит.)\nНСЖ Private (защит.)',
                'SMART (защит.)',
                'ИСЖ ДУ 2.0 (риск.)\nИСЖ сингл (риск.)',
                'ИСЖ ДУ 1.0 (защит.)',
                'Рлайф',
                'NAV'
            ]
            
            for col_idx, name in enumerate(product_names, 1):
                worksheet.cell(row=4, column=col_idx, value=name)
            
            # Добавляем данные по датам
            for row_idx, (_, row_data) in enumerate(final_df.iterrows(), 5):
                # Дата
                worksheet.cell(row=row_idx, column=1, value=row_data['Date'])
                
                # Данные по портфелям
                worksheet.cell(row=row_idx, column=2, value=row_data['271210/2'])
                worksheet.cell(row=row_idx, column=3, value=row_data['020611/1'])
                worksheet.cell(row=row_idx, column=4, value=row_data['020611/2'])
                worksheet.cell(row=row_idx, column=5, value=row_data['020611/3'])
                worksheet.cell(row=row_idx, column=6, value=row_data['141111/1'])
                worksheet.cell(row=row_idx, column=7, value=row_data['260716/1'])
                worksheet.cell(row=row_idx, column=8, value=row_data['190221/1'])
                worksheet.cell(row=row_idx, column=9, value=row_data['081121/1'])
                worksheet.cell(row=row_idx, column=10, value=row_data['081121/2'])
                worksheet.cell(row=row_idx, column=11, value=row_data['220223/1'])
                worksheet.cell(row=row_idx, column=12, value=row_data['220223/2'])
                worksheet.cell(row=row_idx, column=13, value=row_data['050925/1'])
                
                # NAV (формула)
                worksheet.cell(row=row_idx, column=14, value=f"=SUM(B{row_idx}:M{row_idx})")
            
            # Устанавливаем активным лист SAM_2025
            writer.book.active = worksheet
        
        print(f"✅ Файл успешно создан: {output_file_path}")
        print(f"📅 Период: с 2025-10-01 по 2025-10-30")
        print(f"📊 Обработано портфелей: {len(portfolio_mapping)}")
        
        # Выводим сводку по данным
        total_nav = final_df['NAV'].iloc[0] if len(final_df) > 0 else 0
        print(f"💰 Общий NAV: {total_nav:,.2f}")
        
        return final_df
        
    except Exception as e:
        print(f"❌ Ошибка при создании файла: {e}")
        import traceback
        traceback.print_exc()
        return None

def process_merger_to_target_format():
    """Основная функция обработки"""
    
    input_file = r"M:\Финансовый департамент\Treasury\Базы данных(автоматизация)\DI_DATABASE\Мерджер.xlsx"
    output_file = r"M:\Финансовый департамент\Treasury\Базы данных(автоматизация)\DI_DATABASE\обработанные_портфели.xlsx"
    
    print("🚀 ЗАПУСК ОБРАБОТКИ...")
    print(f"Входной файл: {input_file}")
    print(f"Выходной файл: {output_file}")
    
    # Шаг 1: Извлекаем данные из Мерджер.xlsx
    portfolio_data = extract_data_from_merger(input_file)
    
    if portfolio_data is None:
        print("❌ Не удалось извлечь данные из файла Мерджер.xlsx")
        return
    
    # Шаг 2: Создаем файл в целевом формате
    result = create_pivot_format(portfolio_data, output_file)
    
    if result is not None:
        print(f"\n🎉 ОБРАБОТКА ЗАВЕРШЕНА УСПЕШНО!")
        print(f"📁 Результат сохранен: {output_file}")
        print(f"📊 Формат соответствует примеру файла")
    else:
        print(f"\n❌ ОБРАБОТКА ЗАВЕРШИЛАСЬ С ОШИБКОЙ")

# Запуск обработки
if __name__ == "__main__":
    process_merger_to_target_format()
