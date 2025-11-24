import pandas as pd
import os

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

def extract_date_from_file(df):
    """Извлечение даты из файла (последняя колонка, последняя строка с данными)"""
    try:
        # Ищем столбец с датой (обычно последний)
        date_columns = [col for col in df.columns if 'дата' in str(col).lower() or 'Дата отчета' in str(col)]
        
        if date_columns:
            date_col = date_columns[-1]  # берем последний подходящий столбец
            # Ищем последнюю непустую строку в этом столбце
            date_values = df[date_col].dropna()
            if len(date_values) > 0:
                date_value = date_values.iloc[-1]
                if hasattr(date_value, 'strftime'):
                    return date_value.strftime('%d.%m.%Y')
                else:
                    return str(date_value)
        
        # Альтернативный поиск - в последней строке данных
        non_empty_rows = df[df['Портфель'].notna()]
        if len(non_empty_rows) > 0:
            last_row = non_empty_rows.iloc[-1]
            for col in df.columns:
                if 'дата' in str(col).lower():
                    date_value = last_row[col]
                    if pd.notna(date_value):
                        if hasattr(date_value, 'strftime'):
                            return date_value.strftime('%d.%m.%Y')
                        else:
                            return str(date_value)
        
        return "01.10.2025"  # дата по умолчанию из файла
    
    except Exception as e:
        print(f"Ошибка при извлечении даты: {e}")
        return "01.10.2025"

def process_merger_file(input_file_path, output_file_path):
    """Обработка файла Мерджер.xlsx"""
    
    print(f"Читаю файл: {input_file_path}")
    
    try:
        # Читаем файл начиная с заголовка (строка 1 в 0-based индексации)
        df = pd.read_excel(input_file_path, header=1)
        print(f"Найдено строк в таблице: {len(df)}")
        print(f"Колонки: {df.columns.tolist()}")
        
        # Извлекаем дату
        date_str = extract_date_from_file(df)
        print(f"Дата отчета: {date_str}")
        
        # Фильтруем только строки с данными в колонке Портфель
        df = df[df['Портфель'].notna()]
        
        # Убираем строки, где Портфель слишком длинный (возможно заголовки)
        df = df[df['Портфель'].str.len() < 100]
        
        # Убираем полностью пустые строки
        df = df[df.iloc[:, 1:].notna().any(axis=1)]
        
        print(f"Строк после фильтрации: {len(df)}")
        
        # Определяем числовые колонки для группировки
        # В вашем файле нужные колонки: Стоимость (столбец N), НКД (столбец O), Задолженности (столбец P)
        numeric_columns = ['Стоимость', 'НКД,начисленные %', 'Дебеторская/ Кредиторская задолженности']
        
        # Проверяем какие колонки действительно есть в файле
        available_numeric_cols = [col for col in numeric_columns if col in df.columns]
        print(f"Доступные числовые колонки: {available_numeric_cols}")
        
        # Конвертируем числовые колонки
        for col in available_numeric_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # Группируем по портфелю и суммируем числовые колонки
        if available_numeric_cols:
            grouped_df = df.groupby('Портфель')[available_numeric_cols].sum().reset_index()
        else:
            # Если числовых колонок нет, просто группируем по портфелю
            grouped_df = df.groupby('Портфель').size().reset_index(name='Количество записей')
        
        print(f"Сгруппировано портфелей: {len(grouped_df)}")
        
        # Добавляем полное название портфеля из маппинга
        def get_full_portfolio_name(portfolio):
            portfolio_str = str(portfolio)
            for key, value in portfolio_mapping.items():
                if key in portfolio_str:
                    return value
            return portfolio_str
        
        grouped_df['Полное название портфеля'] = grouped_df['Портфель'].apply(get_full_portfolio_name)
        
        # Добавляем дату отчета
        grouped_df['Дата отчета'] = date_str
        
        # Формируем итоговый DataFrame
        base_columns = ['Портфель', 'Полное название портфеля', 'Дата отчета']
        if available_numeric_cols:
            result_columns = base_columns + available_numeric_cols
        else:
            result_columns = base_columns + ['Количество записей']
        
        result_df = grouped_df[result_columns]
        
        # Сохраняем результат
        result_df.to_excel(output_file_path, index=False)
        print(f"✅ Результат сохранен: {output_file_path}")
        
        # Выводим информацию о результате
        print(f"\n📊 Сводка результата:")
        print(f"Обработано портфелей: {len(result_df)}")
        
        if 'Стоимость' in result_df.columns:
            print(f"Общая стоимость: {result_df['Стоимость'].sum():,.2f}")
        if 'НКД,начисленные %' in result_df.columns:
            print(f"Общий НКД: {result_df['НКД,начисленные %'].sum():,.2f}")
        
        # Показываем какие портфели были обработаны
        print("\nОбработанные портфели:")
        for _, row in result_df.iterrows():
            portfolio_info = f"  - {row['Портфель']} -> {row['Полное название портфеля']}"
            if 'Стоимость' in row:
                portfolio_info += f" (Стоимость: {row['Стоимость']:,.2f})"
            print(portfolio_info)
        
        return result_df
        
    except Exception as e:
        print(f"❌ Ошибка при обработке файла: {e}")
        import traceback
        traceback.print_exc()
        return None

def debug_file_structure(input_file_path):
    """Функция для отладки структуры файла"""
    print(f"\n🔍 АНАЛИЗ СТРУКТУРЫ ФАЙЛА: {input_file_path}")
    
    try:
        # Читаем первые 10 строк для анализа
        df_debug = pd.read_excel(input_file_path, header=None, nrows=10)
        
        print("Первые 10 строк файла:")
        for i in range(len(df_debug)):
            non_empty_cells = df_debug.iloc[i].dropna()
            if len(non_empty_cells) > 0:
                print(f"Строка {i}: {list(non_empty_cells.values)}")
        
        # Пробуем найти заголовок
        for i in range(len(df_debug)):
            row_values = df_debug.iloc[i].dropna().values
            if len(row_values) > 0 and 'Портфель' in str(row_values):
                print(f"✅ Заголовок найден в строке {i}")
                break
        else:
            print("❌ Заголовок 'Портфель' не найден в первых 10 строках")
            
    except Exception as e:
        print(f"Ошибка при анализе структуры: {e}")

# Использование
if __name__ == "__main__":
    # Укажи здесь пути к своим файлам
    input_file = "Мерджер.xlsx"  # Ваш файл
    output_file = "обработанные_портфели.xlsx"  # Результат
    
    # Сначала анализируем структуру файла
    debug_file_structure(input_file)
    
    # Затем обрабатываем
    print(f"\n🚀 ЗАПУСК ОБРАБОТКИ...")
    result = process_merger_file(input_file, output_file)
    
    if result is None:
        print("\n❌ Обработка не удалась")
    else:
        print(f"\n✅ ОБРАБОТКА ЗАВЕРШЕНА УСПЕШНО!")
        print(f"Результат сохранен в: {output_file}")
