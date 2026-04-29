import requests
import zipfile
import io
import pandas as pd
from datetime import datetime, timedelta

login = '-----'
password = '---'

# Пути к файлам
source_zip_path = r'M:\Финансовый департамент\Treasury\Базы данных(автоматизация)\DI_DATABASE\FV\processed_moex_data_daily.csv'
isin_list_path = r'M:\Финансовый департамент\Treasury\Базы данных(автоматизация)\DI_DATABASE\FV\isin_list.xlsx'  # Список ISIN
output_excel_path = r'M:\Финансовый департамент\Treasury\Базы данных(автоматизация)\DI_DATABASE\FV\filtered_bonds_data.xlsx'

# Дата за которую скачиваем (вчера)
yesterday = datetime.now() - timedelta(days=1)
year = yesterday.year
month = f'{yesterday.month:02d}'
day = f'{yesterday.day:02d}'

# Формируем URL
zip_url = f'https://iss.moex.com/iss/downloads/engines/stock/markets/bonds/sessions/main/years/{year}/months/{month}/days/{day}/securities_moex_stock_bonds_main_{year}_{month}_{day}.csv.zip'

# 1. Скачиваем и обрабатываем данные
session = requests.Session()
auth_response = session.get('https://passport.moex.com/authenticate', auth=(login, password))

if auth_response.status_code == 200:
    print("Авторизация успешна")
    print(f"Запрашиваю данные за: {year}-{month}-{day}")

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

                # Замена SUR на RUB
                df = df.applymap(lambda x: x.replace('SUR', 'RUB') if isinstance(x, str) else x)

        # Сохраняем полный файл как CSV
        df.to_csv(source_zip_path, index=False, sep=';', decimal='.', encoding='cp1251')
        print(f"Полные данные сохранены: {source_zip_path}")
        print(f"Всего записей: {len(df)}")

        # 2. Загружаем список ISIN из Excel
        try:
            isin_df = pd.read_excel(isin_list_path)
            # Предполагаем, что столбец с ISIN называется 'ISIN' или 'isin'
            # Если название другое - укажите его
            isin_column = None
            for col in isin_df.columns:
                if 'isin' in col.lower():
                    isin_column = col
                    break
            
            if isin_column is None:
                print("Ошибка: Не найден столбец с ISIN в файле")
                isin_list = []
            else:
                isin_list = isin_df[isin_column].dropna().astype(str).tolist()
                print(f"Загружено {len(isin_list)} ISIN из файла: {isin_list_path}")
        except FileNotFoundError:
            print(f"Файл со списком ISIN не найден: {isin_list_path}")
            isin_list = []
        except Exception as e:
            print(f"Ошибка при чтении списка ISIN: {e}")
            isin_list = []

        # 3. Фильтруем данные по ISIN
        if isin_list and 'isin' in df.columns:
            # Приводим ISIN в данных к строковому типу
            df['isin'] = df['isin'].astype(str)
            filtered_df = df[df['isin'].isin(isin_list)]
            
            print(f"Отфильтровано записей по ISIN: {len(filtered_df)}")
            
            # 4. Сохраняем в Excel
            if not filtered_df.empty:
                with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
                    filtered_df.to_excel(writer, sheet_name=f'Bonds_{year}{month}{day}', index=False)
                    # Добавляем лист с информацией о фильтрации
                    info_df = pd.DataFrame({
                        'Дата скачивания': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
                        'Дата данных': [f'{year}-{month}-{day}'],
                        'Всего записей в исходных данных': [len(df)],
                        'Отфильтровано записей': [len(filtered_df)],
                        'Количество ISIN в фильтре': [len(isin_list)]
                    })
                    info_df.to_excel(writer, sheet_name='Информация', index=False)
                
                print(f"Отфильтрованные данные сохранены в Excel: {output_excel_path}")
            else:
                print("Внимание: Нет данных для указанных ISIN")
                # Сохраняем пустой датафрейм с пояснением
                pd.DataFrame({'Предупреждение': ['Нет данных для указанных ISIN'], 
                             'Дата': [f'{year}-{month}-{day}']}).to_excel(output_excel_path, index=False)
        else:
            if 'isin' not in df.columns:
                print(f"Ошибка: В скачанных данных нет колонки 'isin'. Доступные колонки: {df.columns.tolist()}")
            else:
                print("Список ISIN пуст, фильтрация не выполнена")

    else:
        print(f"Ошибка при скачивании архива: {response.status_code}")
        print(f"URL: {zip_url}")
else:
    print(f"Ошибка авторизации: {auth_response.status_code}")
