import pandas as pd

def transform_to_wide_format(input_file_path, output_file_path):
    """Преобразует данные в широкий формат (даты по строкам, портфели по столбцам)"""
    
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
        
        # Переименовываем колонку с датой
        wide_df = wide_df.rename(columns={date_column: 'Date'})
        
        print(f"✅ Преобразовано в широкий формат:")
        print(f"   - Дат: {len(wide_df)}")
        print(f"   - Портфелей: {len(wide_df.columns) - 1}")  # минус колонка Date
        print(f"   - Общая структура: {wide_df.shape}")
        
        # Показываем первые несколько строк
        print(f"\n📊 ПРЕВЬЮ ДАННЫХ:")
        print(wide_df.head())
        
        # Сохраняем в Excel
        wide_df.to_excel(output_file_path, index=False)
        print(f"\n💾 Файл сохранен: {output_file_path}")
        
        return wide_df
        
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()
        return None

def create_final_format_with_nav(wide_df, output_file_path):
    """Создает финальный файл с NAV и правильными названиями портфелей"""
    
    print(f"\n🎯 СОЗДАНИЕ ФИНАЛЬНОГО ФОРМАТА...")
    
    try:
        # Маппинг для сокращенных названий портфелей
        portfolio_mapping = {
            '020611/1': '020611/1 агресс. от 02.06.2011',
            '020611/2': '020611/2 агресс. от 02.06.2011', 
            '020611/3': '020611/3 агресс. от 02.06.2011',
            '081121/1': '081121/1 агресс. от 08.11.2021',
            '081121/2': '081121/2 агресс. от 08.11.2021',
            '141111/1': '141111/1 агресс. от 14.11.2011',
            '190221/1': '190221/1 агресс. от 19.02.2021',
            '220223/1': '220223/1 агресс. от 22.02.2023',
            '220223/2': '220223/2 агресс. от 22.02.2023',
            '260716/1': '260716/1 агресс. от 26.07.2016',
            '271210/2': '271210/2 агресс. от 27.12.2010',
            '050925/1': '050925/1 агресс. от 05.09.2025'
        }
        
        # Переименовываем колонки в сокращенные названия
        column_mapping = {'Date': 'Date'}
        for short_name, full_name in portfolio_mapping.items():
            # Ищем колонку с полным названием
            for col in wide_df.columns:
                if col != 'Date' and full_name in col:
                    column_mapping[col] = short_name
                    break
        
        # Применяем переименование
        final_df = wide_df.rename(columns=column_mapping)
        
        # Оставляем только нужные портфели
        needed_columns = ['Date'] + list(portfolio_mapping.keys())
        final_df = final_df[[col for col in needed_columns if col in final_df.columns]]
        
        # Добавляем NAV как сумму всех портфелей
        portfolio_cols = [col for col in final_df.columns if col != 'Date']
        final_df['NAV'] = final_df[portfolio_cols].sum(axis=1)
        
        print(f"✅ Финальный формат создан:")
        print(f"   - Дат: {len(final_df)}")
        print(f"   - Портфелей: {len(portfolio_cols)}")
        print(f"   - NAV рассчитан")
        
        # Сохраняем финальный файл
        final_df.to_excel(output_file_path, index=False)
        print(f"💾 Финальный файл сохранен: {output_file_path}")
        
        return final_df
        
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        return None

# Запускаем обработку
if __name__ == "__main__":
    input_file = r"M:\Финансовый департамент\Treasury\Базы данных(автоматизация)\DI_DATABASE\Мерджер.xlsx"
    wide_output = r"M:\Финансовый департамент\Treasury\Базы данных(автоматизация)\DI_DATABASE\широкий_формат.xlsx"
    final_output = r"M:\Финансовый департамент\Treasury\Базы данных(автоматизация)\DI_DATABASE\финальный_формат.xlsx"
    
    print("🚀 ЗАПУСК ПРЕОБРАЗОВАНИЯ...")
    
    # Шаг 1: Преобразуем в широкий формат
    wide_data = transform_to_wide_format(input_file, wide_output)
    
    if wide_data is not None:
        # Шаг 2: Создаем финальный формат с NAV
        final_data = create_final_format_with_nav(wide_data, final_output)
        
        if final_data is not None:
            print(f"\n🎉 ПРЕОБРАЗОВАНИЕ ЗАВЕРШЕНО!")
            print(f"📊 ИТОГОВАЯ СТАТИСТИКА:")
            print(f"   - Диапазон дат: {final_data['Date'].min()} - {final_data['Date'].max()}")
            print(f"   - Всего записей: {len(final_data)}")
            print(f"   - Портфелей: {len(final_data.columns) - 2}")  # минус Date и NAV
            print(f"   - Средний NAV: {final_data['NAV'].mean():,.2f}")
        else:
            print("❌ Не удалось создать финальный формат")
    else:
        print("❌ Не удалось преобразовать в широкий формат")
