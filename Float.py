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

        # Сохраняем полный файл как CSV (промежуточный)
        df.to_csv(source_zip_path, index=False, sep=';', decimal='.', encoding='cp1251')
        print(f"Полные данные сохранены: {source_zip_path}")
        print(f"Всего записей: {len(df)}")

        # Загружаем список ISIN из Excel
        try:
            isin_df = pd.read_excel(isin_list_path)
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

        # Фильтруем данные по ISIN
        if isin_list and 'ISIN' in df.columns:
            df['ISIN'] = df['ISIN'].astype(str).str.upper()
            filtered_df = df[df['ISIN'].isin(isin_list)].copy()  # .copy() чтобы не было warning
            
            print(f"Отфильтровано записей по ISIN: {len(filtered_df)}")
            
            # Добавляем колонку с датой данных (первой колонкой)
            filtered_df.insert(0, 'Дата_данных', date_str)
            
            # Сохраняем в Excel с добавлением строк в конец
            if not filtered_df.empty:
                if os.path.exists(output_excel_path):
                    # Файл существует - читаем старые данные и объединяем
                    existing_df = pd.read_excel(output_excel_path)
                    print(f"Существующий файл содержит {len(existing_df)} записей")
                    
                    # Объединяем старые и новые данные
                    combined_df = pd.concat([existing_df, filtered_df], ignore_index=True)
                    print(f"После добавления стало {len(combined_df)} записей")
                    
                    # Сохраняем объединенный файл
                    combined_df.to_excel(output_excel_path, index=False)
                    print(f"Данные ДОБАВЛЕНЫ в конец файла (дата {date_str})")
                else:
                    # Файл не существует - создаем новый
                    filtered_df.to_excel(output_excel_path, index=False)
                    print(f"Создан новый файл с данными за {date_str}")
                
                print(f"Сохранено в: {output_excel_path}")
            else:
                print(f"Внимание: Нет данных для указанных ISIN за {date_str}")
                
                # Даже если данных нет, фиксируем факт запроса
                no_data_row = pd.DataFrame({
                    'Дата_данных': [date_str],
                    'ISIN': ['НЕТ ДАННЫХ'],
                    'Примечание': ['По указанным ISIN данные не найдены']
                })
                
                if os.path.exists(output_excel_path):
                    existing_df = pd.read_excel(output_excel_path)
                    combined_df = pd.concat([existing_df, no_data_row], ignore_index=True)
                    combined_df.to_excel(output_excel_path, index=False)
                    print(f"Зафиксирован факт отсутствия данных за {date_str}")
        else:
            if 'ISIN' not in df.columns:
                print(f"Ошибка: В скачанных данных нет колонки 'ISIN'")
                print(f"Доступные колонки: {df.columns.tolist()}")
            else:
                print("Список ISIN пуст, фильтрация не выполнена")

    else:
        print(f"Ошибка при скачивании архива: {response.status_code}")
        print(f"URL: {zip_url}")
else:
    print(f"Ошибка авторизации: {auth_response.status_code}")
