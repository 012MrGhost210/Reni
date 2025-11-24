def check_small_sample(input_file_path):
    """Проверяем небольшую выборку вручную"""
    
    print(f"\n🔎 РУЧНАЯ ПРОВЕРКА ВЫБОРКИ...")
    
    try:
        df = pd.read_excel(input_file_path, header=0)
        df = df.rename(columns={df.columns[0]: 'Портфель'})
        
        # Берем первые 20 строк
        sample = df.head(20)
        
        print("Первые 20 строк (только нужные колонки):")
        columns_to_show = ['Портфель', 'Стоимость', 'НКД,\nначисленные %', 'Дебеторская/ Кредиторская задолженности']
        
        for col in columns_to_show:
            if col in sample.columns:
                sample[col] = pd.to_numeric(sample[col], errors='coerce').fillna(0)
        
        for _, row in sample.iterrows():
            portfolio = row['Портфель']
            cost = row.get('Стоимость', 0)
            nkd = row.get('НКД,\nначисленные %', 0)
            debt = row.get('Дебеторская/ Кредиторская задолженности', 0)
            total = cost + nkd + debt
            
            print(f"{portfolio[:30]}... | Стоимость: {cost:12.2f} | НКД: {nkd:8.2f} | Задолж: {debt:8.2f} | Итого: {total:12.2f}")
    
    except Exception as e:
        print(f"Ошибка: {e}")

# Запускаем оба анализа
input_file = r"M:\Финансовый департамент\Treasury\Базы данных(автоматизация)\DI_DATABASE\Мерджер.xlsx"
debug_calculations(input_file)
check_small_sample(input_file)
