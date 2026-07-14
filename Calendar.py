import pandas as pd
from datetime import datetime
import os

# ============================================================
# НАСТРАИВАЕМЫЕ ПАРАМЕТРЫ
# ============================================================
CALENDAR_FILE_PATH = r"M:\Финансовый департамент\Treasury\3. ЗАКРЫТИЕ\Отчеты УК\Сводные отчеты УК\Календарь.xlsx"
PORTFOLIO_FILE_PATH = r"M:\Финансовый департамент\Treasury\3. ЗАКРЫТИЕ\Отчеты УК\Сводные отчеты УК\NAV.xlsx"
OUTPUT_FILE_PATH = r"M:\Финансовый департамент\Treasury\3. ЗАКРЫТИЕ\Отчеты УК\Сводные отчеты УК\Coupon_events.xlsx"
# ============================================================

def load_calendar(file_path):
    """Загружает календарь купонов"""
    try:
        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            return pd.DataFrame()
        
        df = pd.read_excel(file_path, skiprows=3)
        df.columns = ['ISIN', 'NAME', 'VOLUME', 'DATE', 'NOMINAL', 'CURRENCY', 
                     'OUTSTANDING_NOMINAL', 'COUPON_RATE', 'PAYMENT', 'PAYMENT_RUB']
        
        # Преобразуем дату
        df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce')
        df = df.dropna(subset=['DATE'])
        
        # Очищаем ISIN
        df['ISIN'] = df['ISIN'].astype(str).str.strip()
        df = df[df['ISIN'] != 'null']
        df = df[df['ISIN'] != '']
        
        # Фильтруем только те строки, где есть выплата
        df['PAYMENT_RUB'] = pd.to_numeric(df['PAYMENT_RUB'], errors='coerce')
        df = df[df['PAYMENT_RUB'].notna()]
        df = df[df['PAYMENT_RUB'] > 0]
        
        print(f"Loaded {len(df)} coupon events from calendar")
        return df
    except Exception as e:
        print(f"Error loading calendar: {e}")
        return pd.DataFrame()

def load_portfolio(file_path):
    """Загружает портфель"""
    try:
        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            return pd.DataFrame()
        
        df = pd.read_excel(file_path)
        
        # Проверяем нужные колонки
        required = ['NAV_DATE', 'PORTFOLIO', 'MANAGEMENT_COMPANY', 'ASSET', 'ISIN']
        if all(col in df.columns for col in required):
            df['NAV_DATE'] = pd.to_datetime(df['NAV_DATE'], errors='coerce')
            df['ISIN'] = df['ISIN'].astype(str).str.strip()
            print(f"Loaded {len(df)} portfolio records")
            return df
        else:
            missing = [col for col in required if col not in df.columns]
            print(f"Missing columns: {missing}")
            return pd.DataFrame()
    except Exception as e:
        print(f"Error loading portfolio: {e}")
        return pd.DataFrame()

def merge_data(calendar_df, portfolio_df):
    """Объединяет календарь и портфель"""
    if calendar_df.empty:
        print("Calendar is empty")
        return pd.DataFrame()
    
    if portfolio_df.empty:
        print("Portfolio is empty")
        return pd.DataFrame()
    
    # Создаем словарь для быстрого поиска по ISIN
    portfolio_lookup = {}
    for _, row in portfolio_df.iterrows():
        isin = row['ISIN']
        if isin not in portfolio_lookup:
            portfolio_lookup[isin] = []
        portfolio_lookup[isin].append({
            'PORTFOLIO': row.get('PORTFOLIO', ''),
            'MANAGEMENT_COMPANY': row.get('MANAGEMENT_COMPANY', ''),
            'ASSET': row.get('ASSET', '')
        })
    
    # Собираем результаты
    results = []
    
    for _, event in calendar_df.iterrows():
        isin = event['ISIN']
        event_date = event['DATE'].strftime('%d.%m.%Y')
        payment = event['PAYMENT_RUB']
        
        # Ищем ISIN в портфеле
        if isin in portfolio_lookup:
            for portfolio_record in portfolio_lookup[isin]:
                results.append({
                    'DATE': event_date,
                    'ISIN': isin,
                    'NAME': event.get('NAME', ''),
                    'ASSET': portfolio_record['ASSET'],
                    'PORTFOLIO': portfolio_record['PORTFOLIO'],
                    'MANAGEMENT_COMPANY': portfolio_record['MANAGEMENT_COMPANY'],
                    'PAYMENT_RUB': payment
                })
        else:
            # Если ISIN не найден в портфеле - пропускаем
            print(f"ISIN not found in portfolio: {isin}")
    
    df_result = pd.DataFrame(results)
    print(f"Created {len(df_result)} merged records")
    return df_result

def main():
    print("=" * 60)
    print("Starting data merge process...")
    print("=" * 60)
    
    # Загружаем данные
    print("\n1. Loading calendar...")
    calendar_df = load_calendar(CALENDAR_FILE_PATH)
    
    print("\n2. Loading portfolio...")
    portfolio_df = load_portfolio(PORTFOLIO_FILE_PATH)
    
    if calendar_df.empty:
        print("ERROR: Calendar data is empty. Stopping.")
        return
    
    if portfolio_df.empty:
        print("ERROR: Portfolio data is empty. Stopping.")
        return
    
    # Объединяем данные
    print("\n3. Merging data...")
    result_df = merge_data(calendar_df, portfolio_df)
    
    if result_df.empty:
        print("ERROR: No data after merge. Stopping.")
        return
    
    # Сортируем по дате
    result_df['DATE_SORT'] = pd.to_datetime(result_df['DATE'], format='%d.%m.%Y')
    result_df = result_df.sort_values('DATE_SORT')
    result_df = result_df.drop('DATE_SORT', axis=1)
    
    # Сохраняем результат
    print(f"\n4. Saving result to: {OUTPUT_FILE_PATH}")
    result_df.to_excel(OUTPUT_FILE_PATH, index=False)
    
    # Выводим статистику
    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)
    print(f"Total coupon events processed: {len(calendar_df)}")
    print(f"Total merged records: {len(result_df)}")
    print(f"Unique ISINs: {result_df['ISIN'].nunique()}")
    print(f"Unique Portfolios: {result_df['PORTFOLIO'].nunique()}")
    print(f"Unique Management Companies: {result_df['MANAGEMENT_COMPANY'].nunique()}")
    
    # Показываем пример
    print("\nSample data:")
    print(result_df.head(10).to_string())
    
    print(f"\n✅ Done! File saved to: {OUTPUT_FILE_PATH}")

if __name__ == "__main__":
    main()
