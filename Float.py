import requests
import zipfile
import io
import pandas as pd
from datetime import datetime, timedelta
import os

login = '-----'
password = '---'

# Пути к файлам
source_zip_path = r'M:\Финансовый департамент\Treasury\Базы данных(автоматизация)\DI_DATABASE\FV\float.csv'
isin_list_path = r'M:\Финансовый департамент\Treasury\Базы данных(автоматизация)\DI_DATABASE\FV\isin_list.xlsx'
output_excel_path = r'M:\Финансовый департамент\Treasury\Базы данных(автоматизация)\DI_DATABASE\FV\filtered_bonds_data.xlsx'

# Дата за которую скачиваем (вчера)
yesterday = datetime.now() - timedelta(days=1)
year = yesterday.year
month = f'{yesterday.month:02d}'
day = f'{yesterday.day:02d}'
date_str = f'{year}-{month}-{day}'
sheet_name = f'Bonds_{year}{month}{day}'

# Формируем URL
zip_url = f'https://iss.moex.com/iss/downloads/engines/stock/markets/bonds/sessions/main/years/{year}/months/{month}/days/{day}/securities_moex_stock_bonds_main_{year}_{month}_{day}.csv.zip'

session = requests.Session()
auth_response = session.get('https://passport.moex.com/authenticate', auth=(login, password))

if auth_response.status_code == 200:
    print("Авторизация успешна")
    print(f"Запрашиваю данные за: {date_str}")

    response = session.get(zip_url)
    if response.status_code == 200:
        print("Файл успешно загружен")

        with zipfile.ZipFile(io.BytesIO(response.content)) as zip_file:
            file_name_in_zip = zip_file.namelist()[0]
            with zip_file.open(file_name_in_zip) as csv_file:
                df = pd.read_csv(csv_file, sep=';', decimal=',', encoding='cp1251', skiprows=2)

                # Преобразование числовых колонок
                for col in df.columns:
                    if df[col].dtype == 'object':
                        try:
                            df[col] = df[col].str.replace(',', '.').astype(float)
                        except (ValueError, AttributeError):
                            pass

        # Сохраняем полный файл как CSV
        df.to_csv(source_zip_path, index=False, sep=';', decimal='.', encoding='cp1251')
        print(f"Полные данные сохранены: {source_zip_path}")
        print(f"Всего записей: {len(df)}")

        # 2. Загружаем список ISIN из Excel
        try:
            isin_df = pd.read_excel(isin_list_path)
            # Ищем столбец с названием ISIN (заглавными)
            if 'ISIN' in isin_df.columns:
                isin_list = isin_df['ISIN'].dropna().astype(str).str.upper().tolist()
                print(f"Загружено {len(isin_list)} ISIN из файла: {isin_list_path}")
                print(f"Первые 5 ISIN: {isin_list[:5]}")
            else:
                print(f"Ошибка: В файле нет столбца 'ISIN'. Доступные столбцы: {isin_df.columns.tolist()}")
                isin_list = []
        except FileNotFoundError:
            print(f"Файл со списком ISIN не найден: {isin_list_path}")
            isin_list = []
        except Exception as e:
            print(f"Ошибка при чтении списка ISIN: {e}")
            isin_list = []

        # 3. Фильтруем данные по ISIN
        if isin_list and 'ISIN' in df.columns:
            df['ISIN'] = df['ISIN'].astype(str).str.upper()
            filtered_df = df[df['ISIN'].isin(isin_list)]
            
            print(f"Отфильтровано записей по ISIN: {len(filtered_df)}")
            
            # Добавляем колонку с датой данных
            filtered_df['Дата_данных'] = date_str
            
            # 4. Сохраняем в Excel с накоплением истории
            if not filtered_df.empty:
                if os.path.exists(output_excel_path):
                    # Файл существует - добавляем новый лист
                    with pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        filtered_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # Обновляем историю загрузок
                    try:
                        history_df = pd.read_excel(output_excel_path, sheet_name='История_загрузок')
                    except:
                        history_df = pd.DataFrame()
                    
                    new_history = pd.DataFrame({
                        'Дата_загрузки': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
                        'Дата_данных': [date_str],
                        'Количество_записей': [len(filtered_df)],
                        'Лист': [sheet_name]
                    })
                    
                    updated_history = pd.concat([history_df, new_history], ignore_index=True)
                    
                    with pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        updated_history.to_excel(writer, sheet_name='История_загрузок', index=False)
                    
                    print(f"Данные ДОБАВЛЕНЫ новым листом '{sheet_name}'")
                else:
                    # Создаем новый файл
                    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
                        filtered_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        history_df = pd.DataFrame({
                            'Дата_загрузки': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
                            'Дата_данных': [date_str],
                            'Количество_записей': [len(filtered_df)],
                            'Лист': [sheet_name]
                        })
                        history_df.to_excel(writer, sheet_name='История_загрузок', index=False)
                    print(f"Создан новый файл с листом '{sheet_name}'")
                
                print(f"Сохранено в: {output_excel_path}")
            else:
                print(f"Внимание: Нет данных для указанных ISIN за {date_str}")
        else:
            if 'ISIN' not in df.columns:
                print(f"Ошибка: В скачанных данных нет колонки 'ISIN'")
            else:
                print("Список ISIN пуст, фильтрация не выполнена")

    else:
        print(f"Ошибка при скачивании архива: {response.status_code}")
else:
    print(f"Ошибка авторизации: {auth_response.status_code}")
