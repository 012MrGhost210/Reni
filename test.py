import pandas as pd

def get_daily_portfolio_totals(input_file_path):
    """Получает сумму по каждому портфелю за каждую дату"""
    
    print(f"📊 ПОЛУЧЕНИЕ ДАННЫХ ПО ДАТАМ И ПОРТФЕЛЯМ...")
    
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
                print(f"💰 Обработана колонка: {col}")
            else:
                print(f"⚠️ Колонка '{col}' не найдена")
        
        # Фильтруем валидные данные
        df = df[df['Портфель'].notna()]
        df = df[df[date_column].notna()]
        
        print(f"📅 Уникальные даты в файле: {df[date_column].dt.date.unique()}")
        print(f"🎯 Уникальные портфели: {df['Портфель'].nunique()}")
        
        # Суммируем нужные колонки
        df['Общая_сумма'] = 0
        for col in money_columns:
            if col in df.columns:
                df['Общая_сумма'] += df[col]
        
        # Группируем по дате и портфелю
        result = df.groupby([date_column, 'Портфель'])['Общая_сумма'].sum().reset_index()
        
        print(f"\n📈 РЕЗУЛЬТАТ - СУММЫ ПО ДАТАМ И ПОРТФЕЛЯМ:")
        print(f"Всего записей: {len(result)}")
        
        # Показываем данные по датам
        dates = result[date_column].dt.date.unique()
        for date in sorted(dates):
            date_data = result[result[date_column].dt.date == date]
            print(f"\n📅 {date}:")
            print(f"   Всего портфелей: {len(date_data)}")
            print(f"   Общая сумма за день: {date_data['Общая_сумма'].sum():,.2f}")
            
            # Показываем топ-5 портфелей за эту дату
            top_portfolios = date_data.nlargest(5, 'Общая_сумма')
            for _, row in top_portfolios.iterrows():
                print(f"   - {row['Портфель']}: {row['Общая_сумма']:,.2f}")
        
        return result
        
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()
        return None

def save_daily_totals_to_excel(data, output_file_path):
    """Сохраняет результаты в Excel"""
    
    if data is None:
        return
    
    try:
        # Сохраняем в Excel
        with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
            # Основная таблица
            data.to_excel(writer, sheet_name='Суммы_по_датам', index=False)
            
            # Сводка по датам
            summary_by_date = data.groupby(data.iloc[:, 0].dt.date)['Общая_сумма'].agg(['sum', 'count']).reset_index()
            summary_by_date.columns = ['Дата', 'Общая_сумма', 'Количество_портфелей']
            summary_by_date.to_excel(writer, sheet_name='Сводка_по_датам', index=False)
            
            # Сводка по портфелям
            summary_by_portfolio = data.groupby('Портфель')['Общая_сумма'].agg(['sum', 'count']).reset_index()
            summary_by_portfolio.columns = ['Портфель', 'Общая_сумма', 'Количество_дней']
            summary_by_portfolio.to_excel(writer, sheet_name='Сводка_по_портфелям', index=False)
        
        print(f"\n💾 Результаты сохранены в: {output_file_path}")
        
    except Exception as e:
        print(f"❌ Ошибка при сохранении: {e}")

# Запускаем обработку
if __name__ == "__main__":
    input_file = r"M:\Финансовый департамент\Treasury\Базы данных(автоматизация)\DI_DATABASE\Мерджер.xlsx"
    output_file = r"M:\Финансовый департамент\Treasury\Базы данных(автоматизация)\DI_DATABASE\ежедневные_суммы.xlsx"
    
    print("🚀 ЗАПУСК РАСЧЕТА ЕЖЕДНЕВНЫХ СУММ...")
    
    # Получаем данные
    daily_totals = get_daily_portfolio_totals(input_file)
    
    if daily_totals is not None:
        # Сохраняем в Excel
        save_daily_totals_to_excel(daily_totals, output_file)
        
        print(f"\n✅ РАСЧЕТ ЗАВЕРШЕН!")
        print(f"📊 Получено {len(daily_totals)} записей")
        print(f"📅 Охвачено дат: {daily_totals.iloc[:, 0].nunique()}")
        print(f"🎯 Охвачено портфелей: {daily_totals['Портфель'].nunique()}")
    else:
        print("❌ Не удалось получить данные")
