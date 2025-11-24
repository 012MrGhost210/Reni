import pandas as pd

def analyze_merger_structure(input_file_path):
    """Тщательно анализирует структуру файла Мерджер.xlsx"""
    
    print(f"🔍 АНАЛИЗ СТРУКТУРЫ ФАЙЛА: {input_file_path}")
    
    try:
        # Читаем первые строки чтобы понять структуру
        df_raw = pd.read_excel(input_file_path, header=None, nrows=10)
        print("Первые 10 строк файла:")
        print(df_raw)
        print("\n" + "="*50)
        
        # Пробуем найти заголовок
        for i in range(5):
            df_test = pd.read_excel(input_file_path, header=i)
            print(f"Заголовок в строке {i}: {df_test.columns.tolist()[:10]}...")
            
            # Проверяем есть ли колонка с портфелями
            first_col = df_test.columns[0]
            if 'портф' in str(first_col).lower() or 'portfolio' in str(first_col).lower():
                print(f"✅ Найден заголовок портфелей в строке {i}")
                header_row = i
                break
        else:
            print("❌ Не найден заголовок портфелей")
            return None
        
        # Читаем файл с правильным заголовком
        df = pd.read_excel(input_file_path, header=header_row)
        print(f"\n📊 СТРУКТУРА ДАННЫХ:")
        print(f"Всего колонок: {len(df.columns)}")
        print(f"Всего строк: {len(df)}")
        
        # Показываем все названия колонок
        print("\n📋 ВСЕ КОЛОНКИ:")
        for i, col in enumerate(df.columns):
            print(f"{i:2d}. {col}")
        
        # Анализируем первую колонку (портфели)
        print(f"\n🎯 ПЕРВАЯ КОЛОНКА (портфели):")
        print(f"Название: '{df.columns[0]}'")
        print(f"Уникальных значений: {df.iloc[:, 0].nunique()}")
        print(f"Примеры значений:")
        print(df.iloc[:, 0].dropna().head(10).tolist())
        
        # Ищем числовые колонки которые нужно суммировать
        numeric_columns = []
        money_indicators = ['стоимость', 'нкд', 'начислен', 'дебитор', 'кредитор', 'задолженност']
        
        for i, col_name in enumerate(df.columns):
            col_str = str(col_name).lower()
            if any(indicator in col_str for indicator in money_indicators):
                print(f"💰 Найдена денежная колонка [{i}]: {col_name}")
                numeric_columns.append((i, col_name))
        
        return df, numeric_columns, header_row
        
    except Exception as e:
        print(f"❌ Ошибка при анализе: {e}")
        import traceback
        traceback.print_exc()
        return None

def calculate_correct_totals(input_file_path):
    """Правильно рассчитывает итоги по портфелям"""
    
    print(f"\n🧮 ПРАВИЛЬНЫЙ РАСЧЕТ ИТОГОВ...")
    
    try:
        # Сначала анализируем структуру
        analysis_result = analyze_merger_structure(input_file_path)
        if analysis_result is None:
            return None
            
        df, numeric_columns, header_row = analysis_result
        
        # Читаем заново с правильным заголовком
        df = pd.read_excel(input_file_path, header=header_row)
        
        # Переименовываем первую колонку
        df = df.rename(columns={df.columns[0]: 'Портфель'})
        
        # Фильтруем валидные строки с портфелями
        df = df[df['Портфель'].notna()]
        df = df[~df['Портфель'].astype(str).str.contains('итог', case=False, na=False)]
        df = df[df['Портфель'].astype(str).str.len() < 100]
        
        print(f"📊 Валидных строк с портфелями: {len(df)}")
        
        # Конвертируем числовые колонки
        for col_idx, col_name in numeric_columns:
            df[col_name] = pd.to_numeric(df[col_name], errors='coerce').fillna(0)
            print(f"Конвертирована {col_name}: сумма = {df[col_name].sum():,.2f}")
        
        # Суммируем все денежные колонки для каждого портфеля
        df['Итог_портфеля'] = 0
        for col_idx, col_name in numeric_columns:
            df['Итог_портфеля'] += df[col_name]
        
        # Группируем по портфелям
        portfolio_totals = df.groupby('Портфель')['Итог_портфеля'].sum().reset_index()
        
        # Добавляем идентификатор портфеля
        def get_portfolio_id(portfolio):
            portfolio_str = str(portfolio)
            for key in portfolio_mapping.keys():
                if key in portfolio_str:
                    return key
            return None
        
        portfolio_totals['Portfolio_ID'] = portfolio_totals['Портфель'].apply(get_portfolio_id)
        
        print(f"\n📈 РЕЗУЛЬТАТЫ РАСЧЕТА:")
        total_sum = 0
        for _, row in portfolio_totals.iterrows():
            if row['Portfolio_ID']:
                print(f"  ✅ {row['Portfolio_ID']}: {row['Итог_портфеля']:,.2f}")
                total_sum += row['Итог_портфеля']
            else:
                print(f"  ⚠️ {row['Портфель']}: {row['Итог_портфеля']:,.2f} (не распознан)")
        
        print(f"💰 ОБЩАЯ СУММА: {total_sum:,.2f}")
        
        # Создаем словарь с правильными значениями
        correct_portfolio_values = {}
        for _, row in portfolio_totals.iterrows():
            if row['Portfolio_ID']:
                correct_portfolio_values[row['Portfolio_ID']] = row['Итог_портфеля']
        
        return correct_portfolio_values
        
    except Exception as e:
        print(f"❌ Ошибка при расчете: {e}")
        import traceback
        traceback.print_exc()
        return None

# Запускаем анализ
if __name__ == "__main__":
    input_file = r"M:\Финансовый департамент\Treasury\Базы данных(автоматизация)\DI_DATABASE\Мерджер.xlsx"
    
    print("🔍 ДЕТАЛЬНЫЙ АНАЛИЗ ФАЙЛА МЕРДЖЕР...")
    correct_values = calculate_correct_totals(input_file)
    
    if correct_values:
        print(f"\n🎯 ПРАВИЛЬНЫЕ ЗНАЧЕНИЯ ПОРТФЕЛЕЙ:")
        for portfolio_id, value in correct_values.items():
            print(f"  {portfolio_id}: {value:,.2f}")
    else:
        print("❌ Не удалось проанализировать файл")
