import pandas as pd

def transform_to_wide_format_simple(input_file_path, output_file_path):
    """Преобразует данные в широкий формат (простая версия)"""
    
    print(f"🔄 ПРЕОБРАЗОВАНИЕ В ШИРОКИЙ ФОРМАТ...")
    
    try:
        # Читаем файл
        df = pd.read_excel(input_file_path, header=0)
        df = df.rename(columns={df.columns[0]: 'Портфель'})
        
        # Определяем нужные колонки
        money_columns = [
            'Стоимость',
            'НКД,\nначисленные %', 
            'Дебеторская/ Кредиторская задолженности'
        ]
        
        # Находим колонку с датой отчета
        date_column = None
        for col in df.columns:
            if 'дата' in str(col).lower() and 'отчет' in str(col).lower():
                date_column = col
                break
        
        if date_column is None:
            print("❌ Не найдена колонка с датой отчета")
            return None
        
        print(f"Колонка с датой: '{date_column}'")
        
        # Конвертируем дату и числовые колонки
        df[date_column] = pd.to_datetime(df[date_column], errors='coerce')
        
        for col in money_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # Фильтруем валидные данные
        df = df[df['Портфель'].notna()]
        df = df[df[date_column].notna()]
        
        # Суммируем нужные колонки
        df['Общая_сумма'] = 0
        for col in money_columns:
            if col in df.columns:
                df['Общая_сумма'] += df[col]
        
        # Группируем по дате и портфелю
        grouped = df.groupby([date_column, 'Портфель'])['Общая_сумма'].sum().reset_index()
        
        # Преобразуем в широкий формат (pivot)
        wide_df = grouped.pivot_table(
            index=date_column,
            columns='Портфель',
            values='Общая_сумма',
            aggfunc='sum'
        ).reset_index()
        
        # Заполняем пропущенные значения нулями
        wide_df = wide_df.fillna(0)
        
        # Переименовываем колонку с датой и форматируем даты
        wide_df = wide_df.rename(columns={date_column: 'Date'})
        wide_df['Date'] = wide_df['Date'].dt.strftime('%d.%m.%Y')
        
        print(f"✅ Преобразовано в широкий формат:")
        print(f"   - Дат: {len(wide_df)}")
        print(f"   - Портфелей: {len(wide_df.columns) - 1}")
        
        # Показываем список портфелей
        portfolio_columns = [col for col in wide_df.columns if col != 'Date']
        print(f"   - Список портфелей: {len(portfolio_columns)} шт")
        
        # Сохраняем первоначальный широкий формат
        wide_df.to_excel(output_file_path, index=False)
        print(f"💾 Широкий формат сохранен: {output_file_path}")
        
        return wide_df
        
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()
        return None

def rename_portfolio_columns(wide_df, output_file_path):
    """Переименовывает колонки с портфелями по маппингу"""
    
    print(f"\n🔄 ПЕРЕИМЕНОВАНИЕ ПОРТФЕЛЕЙ...")
    
    try:
        # Маппинг для переименования портфелей
        portfolio_mapping = {
            '020611/1 агресс. от 02.06.2011': '020611/1',
            '020611/2 агресс. от 02.06.2011': '020611/2', 
            '020611/3 агресс. от 02.06.2011': '020611/3',
            '081121/1 агресс. от 08.11.2021': '081121/1',
            '081121/2 агресс. от 08.11.2021': '081121/2',
            '141111/1 агресс. от 14.11.2011': '141111/1',
            '190221/1 агресс. от 19.02.2021': '190221/1',
            '220223/1 агресс. от 22.02.2023': '220223/1',
            '220223/2 агресс. от 22.02.2023': '220223/2',
            '260716/1 агресс. от 26.07.2016': '260716/1',
            '271210/2 агресс. от 27.12.2010': '271210/2',
            '050925/1 агресс. от 05.09.2025': '050925/1'
        }
        
        # Создаем словарь для переименования колонок
        column_rename = {'Date': 'Date'}
        
        # Для каждой колонки в данных
        for col in wide_df.columns:
            if col != 'Date':
                # Ищем соответствие в маппинге
                new_name = None
                for old_name, new_name_val in portfolio_mapping.items():
                    if old_name in col:
                        new_name = new_name_val
                        break
                
                if new_name:
                    column_rename[col] = new_name
                    print(f"   ✅ {col} -> {new_name}")
                else:
                    # Оставляем оригинальное название если нет в маппинге
                    column_rename[col] = col
                    print(f"   ⚠️ {col} -> оставлено без изменений")
        
        # Применяем переименование
        renamed_df = wide_df.rename(columns=column_rename)
        
        # Сохраняем результат
        renamed_df.to_excel(output_file_path, index=False)
        print(f"💾 Файл с переименованными портфелями сохранен: {output_file_path}")
        
        return renamed_df
        
    except Exception as e:
        print(f"❌ Ошибка при переименовании: {e}")
        return None

# Запускаем обработку
if __name__ == "__main__":
    input_file = r"M:\Финансовый департамент\Treasury\Базы данных(автоматизация)\DI_DATABASE\Мерджер.xlsx"
    wide_output = r"M:\Финансовый департамент\Treasury\Базы данных(автоматизация)\DI_DATABASE\широкий_формат.xlsx"
    final_output = r"M:\Финансовый департамент\Treasury\Базы данных(автоматизация)\DI_DATABASE\финальный_формат.xlsx"
    
    print("🚀 ЗАПУСК ПРЕОБРАЗОВАНИЯ...")
    
    # Шаг 1: Создаем широкий формат
    wide_data = transform_to_wide_format_simple(input_file, wide_output)
    
    if wide_data is not None:
        # Шаг 2: Переименовываем портфели
        final_data = rename_portfolio_columns(wide_data, final_output)
        
        if final_data is not None:
            print(f"\n🎉 ПРЕОБРАЗОВАНИЕ ЗАВЕРШЕНО!")
            print(f"📊 ИТОГОВАЯ СТАТИСТИКА:")
            print(f"   - Дат: {len(final_data)}")
            print(f"   - Портфелей: {len(final_data.columns) - 1}")
            print(f"   - Диапазон дат: {final_data['Date'].min()} - {final_data['Date'].max()}")
            
            # Показываем список портфелей в финальном файле
            portfolio_cols = [col for col in final_data.columns if col != 'Date']
            print(f"   - Портфели в файле: {portfolio_cols}")
        else:
            print("❌ Не удалось переименовать портфели")
    else:
        print("❌ Не удалось создать широкий формат")
