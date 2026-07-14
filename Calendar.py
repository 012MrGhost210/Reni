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
        
        # Приводим PAYMENT_RUB к числовому формату
        df['PAYMENT_RUB'] = pd.to_numeric(df['PAYMENT_RUB'], errors='coerce')
        
        print(f"Calendar loaded: {len(df)} rows")
        return df
    except Exception as e:
        print(f"Error loading calendar: {e}")
        return pd.DataFrame()

def load_portfolio(file_path):
    """Загружает портфель NAV"""
    try:
        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            return pd.DataFrame()
        
        df = pd.read_excel(file_path)
        
        # Проверяем нужные колонки
        required = ['NAV_DATE', 'PORTFOLIO', 'MANAGEMENT_COMPANY', 'ASSET', 'ISIN']
        if all(col in df.columns for col in required):
            df['ISIN'] = df['ISIN'].astype(str).str.strip()
            print(f"Portfolio loaded: {len(df)} rows")
            return df
        else:
            missing = [col for col in required if col not in df.columns]
            print(f"Missing columns: {missing}")
            return pd.DataFrame()
    except Exception as e:
        print(f"Error loading portfolio: {e}")
        return pd.DataFrame()

def merge_data(calendar_df, portfolio_df):
    """Объединяет данные: берем ISIN из портфеля и ищем в календаре"""
    if portfolio_df.empty:
        print("Portfolio is empty")
        return pd.DataFrame()
    
    if calendar_df.empty:
        print("Calendar is empty")
        return pd.DataFrame()
    
    # Создаем словарь для быстрого поиска по ISIN в календаре
    calendar_dict = {}
    for _, row in calendar_df.iterrows():
        isin = row['ISIN']
        if isin not in calendar_dict:
            calendar_dict[isin] = []
        calendar_dict[isin].append({
            'DATE': row['DATE'],
            'NAME': row.get('NAME', ''),
            'PAYMENT_RUB': row.get('PAYMENT_RUB', 0)
        })
    
    # Собираем результаты
    results = []
    found_count = 0
    not_found_count = 0
    
    for _, nav_row in portfolio_df.iterrows():
        isin = nav_row['ISIN']
        portfolio = nav_row.get('PORTFOLIO', '')
        asset = nav_row.get('ASSET', '')
        management = nav_row.get('MANAGEMENT_COMPANY', '')
        
        # Ищем ISIN в календаре
        if isin in calendar_dict:
            for event in calendar_dict[isin]:
                # Пропускаем, если выплата = 0 или NaN
                if pd.isna(event['PAYMENT_RUB']) or event['PAYMENT_RUB'] <= 0:
                    continue
                    
                results.append({
                    'DATE': event['DATE'].strftime('%d.%m.%Y'),
                    'ISIN': isin,
                    'NAME': event['NAME'],
                    'ASSET': asset,
                    'PORTFOLIO': portfolio,
                    'MANAGEMENT_COMPANY': management,
                    'PAYMENT_RUB': event['PAYMENT_RUB']
                })
            found_count += 1
        else:
            not_found_count += 1
    
    df_result = pd.DataFrame(results)
    print(f"\nISINs found in calendar: {found_count}")
    print(f"ISINs NOT found in calendar: {not_found_count}")
    print(f"Total merged records: {len(df_result)}")
    
    return df_result

def main():
    print("=" * 60)
    print("Coupon Events Generator")
    print("=" * 60)
    
    print("\n1. Loading calendar...")
    calendar_df = load_calendar(CALENDAR_FILE_PATH)
    
    print("\n2. Loading portfolio (NAV)...")
    portfolio_df = load_portfolio(PORTFOLIO_FILE_PATH)
    
    if portfolio_df.empty:
        print("ERROR: Portfolio is empty. Stopping.")
        return
    
    if calendar_df.empty:
        print("ERROR: Calendar is empty. Stopping.")
        return
    
    print(f"\n3. Merging data (matching ISINs from portfolio to calendar)...")
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
    print(f"Total ISINs in portfolio: {len(portfolio_df)}")
    print(f"Total coupon events found: {len(result_df)}")
    print(f"Unique ISINs with coupons: {result_df['ISIN'].nunique()}")
    print(f"Unique Portfolios: {result_df['PORTFOLIO'].nunique()}")
    print(f"Unique Management Companies: {result_df['MANAGEMENT_COMPANY'].nunique()}")
    
    # Показываем пример
    print("\nSample data (first 10 rows):")
    print(result_df.head(10).to_string())
    
    print(f"\n✅ Done! File saved to: {OUTPUT_FILE_PATH}")

if __name__ == "__main__":
    main()
