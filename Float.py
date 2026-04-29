import requests
import zipfile
import io
import pandas as pd
from datetime import datetime, timedelta

login = '-----'
password = '---'
file_path = r'M:\Финансовый департамент\Treasury\Базы данных(автоматизация)\DI_DATABASE\FV\processed_moex_data_daily.csv'

# Данные за вчерашний день
yesterday = datetime.now() - timedelta(days=1)
year = yesterday.year
month = f'{yesterday.month:02d}'
day = f'{yesterday.day:02d}'

# Формируем URL для дневных данных
zip_url = f'https://iss.moex.com/iss/downloads/engines/stock/markets/bonds/sessions/main/years/{year}/months/{month}/days/{day}/securities_moex_stock_bonds_main_{year}_{month}_{day}.csv.zip'

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

                # Обработка колонок с числами
                for col in df.columns:
                    if df[col].dtype == 'object':
                        try:
                            df[col] = df[col].str.replace(',', '.').astype(float)
                        except (ValueError, AttributeError):
                            pass

                # Замена SUR на RUB
                df = df.applymap(lambda x: x.replace('SUR', 'RUB') if isinstance(x, str) else x)

        # Сохраняем файл (перезаписываем, если существует)
        df.to_csv(file_path, index=False, sep=';', decimal='.', encoding='cp1251')
        print(f"Файл сохранён: {file_path}")
        print(f"Кол-во записей: {len(df)}")

    else:
        print(f"Ошибка при скачивании архива: {response.status_code}")
        print(f"Проверьте, есть ли данные за эту дату на Мосбирже")
else:
    print(f"Ошибка авторизации: {auth_response.status_code}")
