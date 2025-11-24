import pandas as pd

def create_final_file_improved(input_file_path, output_file_path):
    """Создает финальный файл с частичным переименованием и фильтрацией"""
    
    print(f"🚀 СОЗДАНИЕ ФИНАЛЬНОГО ФАЙЛА...")
    
    try:
        # Читаем исходный файл
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
        
        print(f"📅 Колонка с датой: '{date_column}'")
        
        # Конвертируем дату и числовые колонки
        df[date_column] = pd.to_datetime(df[date_column], errors='coerce')
        
        for col in money_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # Фильтруем валидные данные и убираем портфели с REZHS
        df = df[df['Портфель'].notna()]
        df = df[df[date_column].notna()]
        df = df[~df['Портфель'].astype(str).str.contains('REZHS', case=False, na=False)]
        
        print(f"✅ После фильтрации REZHS осталось строк: {len(df)}")
        
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
        
        print(f"✅ Широкий формат создан:")
        print(f"   - Дат: {len(wide_df)}")
        print(f"   - Портфелей: {len(wide_df.columns) - 1}")
        
        # Функция для извлечения короткого названия портфеля
        def extract_short_name(full_name):
            # Ищем паттерн "XXX/XXX" в названии портфеля
            import re
            match = re.search(r'(\d{6}/\d{1,2})', str(full_name))
            if match:
                return match.group(1)
            return full_name
        
        # Переименовываем колонки по частичному совпадению
        column_rename = {'Date': 'Date'}
        
        print(f"\n🔄 ПЕРЕИМЕНОВАНИЕ ПОРТФЕЛЕЙ:")
        for col in wide_df.columns:
            if col == 'Date':
                continue
                
            # Извлекаем короткое название
            short_name = extract_short_name(col)
            
            if short_name != col:
                column_rename[col] = short_name
                print(f"   ✅ '{col}' -> '{short_name}'")
            else:
                column_rename[col] = col
                print(f"   ⚠️ '{col}' -> оставлено без изменений")
        
        # Применяем переименование
        final_df = wide_df.rename(columns=column_rename)
        
        # Сохраняем финальный файл
        final_df.to_excel(output_file_path, index=False)
        print(f"\n💾 Финальный файл сохранен: {output_file_path}")
        
        # Статистика
        print(f"\n📊 ИТОГОВАЯ СТАТИСТИКА:")
        print(f"   - Дат: {len(final_df)}")
        print(f"   - Портфелей: {len(final_df.columns) - 1}")
        print(f"   - Диапазон дат: {final_df['Date'].min()} - {final_df['Date'].max()}")
        
        # Показываем список портфелей в финальном файле
        portfolio_cols = [col for col in final_df.columns if col != 'Date']
        print(f"   - Портфели в файле: {portfolio_cols}")
        
        return final_df
        
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()
        return None

# Запускаем обработку
if __name__ == "__main__":
    input_file = r"M:\Финансовый департамент\Treasury\Базы данных(автоматизация)\DI_DATABASE\Мерджер.xlsx"
    output_file = r"M:\Финансовый департамент\Treasury\Базы данных(автоматизация)\DI_DATABASE\финальный_формат.xlsx"
    
    print("🚀 ЗАПУСК СОЗДАНИЯ ФИНАЛЬНОГО ФАЙЛА...")
    
    result = create_final_file_improved(input_file, output_file)
    
    if result is not None:
        print(f"\n🎉 ФАЙЛ УСПЕШНО СОЗДАН!")
        print(f"📁 Расположение: {output_file}")
    else:
        print("❌ Не удалось создать файл")
