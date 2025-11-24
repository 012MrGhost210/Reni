import pandas as pd

def debug_calculations(input_file_path):
    """Детальный дебаг расчета сумм"""
    
    print(f"🔍 ДЕТАЛЬНЫЙ ДЕБАГ РАСЧЕТОВ...")
    
    try:
        # Читаем файл
        df = pd.read_excel(input_file_path, header=0)
        df = df.rename(columns={df.columns[0]: 'Портфель'})
        
        # Фильтруем валидные строки
        df = df[df['Портфель'].notna()]
        df = df[~df['Портфель'].astype(str).str.contains('итог', case=False, na=False)]
        df = df[df['Портфель'].astype(str).str.len() < 100]
        
        print(f"Всего строк: {len(df)}")
        
        # Смотрим на конкретные колонки
        target_columns = ['Стоимость', 'НКД,\nначисленные %', 'Дебеторская/ Кредиторская задолженности']
        
        print(f"\n📊 АНАЛИЗ КОЛОНОК:")
        for col in target_columns:
            if col in df.columns:
                # Конвертируем в числа
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                
                # Статистика по колонке
                total = df[col].sum()
                avg = df[col].mean()
                max_val = df[col].max()
                min_val = df[col].min()
                
                print(f"\n{col}:")
                print(f"  Сумма: {total:,.2f}")
                print(f"  Среднее: {avg:,.2f}")
                print(f"  Максимум: {max_val:,.2f}")
                print(f"  Минимум: {min_val:,.2f}")
                print(f"  Не нулевых значений: {(df[col] != 0).sum()}")
                
                # Покажем первые 10 ненулевых значений
                non_zero = df[df[col] != 0][['Портфель', col]].head(10)
                if len(non_zero) > 0:
                    print(f"  Примеры ненулевых значений:")
                    for _, row in non_zero.iterrows():
                        print(f"    {row['Портфель']}: {row[col]:,.2f}")
            else:
                print(f"❌ Колонка '{col}' не найдена")
        
        # Теперь посмотрим на один конкретный портфель
        print(f"\n🎯 АНАЛИЗ КОНКРЕТНОГО ПОРТФЕЛЯ:")
        sample_portfolio = df[df['Портфель'].str.contains('020611/1', na=False)].head(1)
        if len(sample_portfolio) > 0:
            portfolio_name = sample_portfolio['Портфель'].iloc[0]
            print(f"Портфель: {portfolio_name}")
            
            for col in target_columns:
                if col in sample_portfolio.columns:
                    value = sample_portfolio[col].iloc[0]
                    print(f"  {col}: {value:,.2f}")
        
        # Суммируем только по нужным колонкам
        print(f"\n🧮 ПРАВИЛЬНЫЙ РАСЧЕТ:")
        df['Итог'] = 0
        for col in target_columns:
            if col in df.columns:
                df['Итог'] += df[col]
        
        # Группируем по портфелям
        portfolio_totals = df.groupby('Портфель')['Итог'].sum().reset_index()
        
        # Покажем топ-10 портфелей по сумме
        print(f"\n📈 ТОП-10 ПОРТФЕЛЕЙ ПО СУММЕ:")
        top_portfolios = portfolio_totals.nlargest(10, 'Итог')
        for _, row in top_portfolios.iterrows():
            print(f"  {row['Портфель']}: {row['Итог']:,.2f}")
        
        # Общая сумма
        total_sum = portfolio_totals['Итог'].sum()
        print(f"\n💰 ОБЩАЯ СУММА ВСЕХ ПОРТФЕЛЕЙ: {total_sum:,.2f}")
        
        # Проверим, может быть я неправильно понимаю валюту?
        print(f"\n💱 ПРОВЕРКА ВАЛЮТЫ:")
        if 'Валюта котировки' in df.columns:
            currencies = df['Валюта котировки'].value_counts()
            print("Распределение по валютам:")
            for currency, count in currencies.items():
                print(f"  {currency}: {count} записей")
        
        return portfolio_totals
        
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()
        return None

# Запускаем дебаг
if __name__ == "__main__":
    input_file = r"M:\Финансовый департамент\Treasury\Базы данных(автоматизация)\DI_DATABASE\Мерджер.xlsx"
    debug_calculations(input_file)
