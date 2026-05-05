import requests
import zipfile
import io
import pandas as pd
from datetime import datetime, timedelta

login = ''
password = ''
file_path = r'M:\Финансовый департамент\Treasury\Базы данных(автоматизация)\DI_DATABASE\FV\processed_moex_data_month.csv'
month = datetime.now().month - 1
month_form = f'{month:02d}'
yesterday = datetime.now().day - 1
year = datetime.now().year

zip_url = f'https://iss.moex.com/iss/downloads/engines/stock/markets/bonds/years/2026/months/{month_form}/securities_moex_stock_bonds_2026_{month_form}.csv.zip'

# Список нужных столбцов
required_columns = [
    'ISIN', 'BOARDNAME', 'SHORTNAME', 'REGNUMBER', 'MATDATE',
    'FACEVALUE', 'CURRENCYID', 'WAPRICE', 'HIGHBID', 'LOWOFFER',
    'BID', 'OFFER', 'ACCINT', 'MARKETPRICE3'
]

session = requests.Session()
auth_response = session.get('https://passport.moex.com/authenticate', auth=(login, password))

if auth_response.status_code == 200:
    print("Авторизация успешна")

    response = session.get(zip_url)
    if response.status_code == 200:
        print("Файл успешно загружен")

        with zipfile.ZipFile(io.BytesIO(response.content)) as zip_file:
            file_name_in_zip = zip_file.namelist()[0]

            with zip_file.open(file_name_in_zip) as csv_file:
                df = pd.read_csv(csv_file, sep=';', decimal=',', encoding='cp1251', skiprows=2)

                # Оставляем только нужные столбцы, которые есть в DataFrame
                existing_columns = [col for col in required_columns if col in df.columns]
                df = df[existing_columns]

                for col in df.columns:
                    if df[col].dtype == 'object':
                        try:
                            df[col] = df[col].str.replace(',', '.').astype(float)
                        except (ValueError, AttributeError):
                            pass

                df = df.applymap(lambda x: x.replace('SUR', 'RUB') if isinstance(x, str) else x)

        output_file = file_path
        df.to_csv(output_file, index=False, sep=';', decimal='.', encoding='cp1251')
        print(f"Файл сохранён как: processed_moex_data_month")
        print(f"Оставлены столбцы: {existing_columns}")

    else:
        print(f"Ошибка при скачивании архива: {response.status_code}")
else:
    print(f"Ошибка авторизации: {auth_response.status_code}")

