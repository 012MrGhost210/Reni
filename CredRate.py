Нужно сделать эксель-таблицу с рейтингами по выбранным контрагентам. рейтинги АКРА, НКР, Эксперт РА. данные нужны на дату с возможностью собирать и обновлять данные по запросу
import pandas as pd
https://dh2.efir-net.ru/swagger/index.html?urls.primaryName=DataHub%20v2.0

import json
import requests
import time
import math

from io import StringIO
from tqdm import tqdm
from requests.exceptions import RequestException
from requests.packages.urllib3.exceptions import InsecureRequestWarning
from tenacity import retry, stop_after_attempt, wait_exponential

api_url = '...'
api_login = '...'
api_pass = '...'

# Интервал между запросами для соблюдения ограничения
REQUEST_INTERVAL = 0.25  # 1/5 секунды

requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

def log_error(message, response=None):
    print(f"Error: {message}")
    if response:
        print(f"Status Code: {response.status_code}, Response: {response.text}")

@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=2, max=10))
def doPostRequest(url, body, token):
    headers = {'Content-Type': 'application/json'}
    if token:
        headers['authorization'] = f'Bearer {token}'

    try:
        response = requests.post(url, json=body, headers=headers, timeout=10, verify=False)
        response.raise_for_status()
        return response.json()
    except RequestException as e:
        log_error("Request exception occurred", response)
        raise e
    except json.JSONDecodeError:
        log_error("Failed to decode JSON response")
        raise
    finally:
        time.sleep(REQUEST_INTERVAL)

def getToken(login, password):
    url = f"{api_url}/Account/Login"
    body = {'login': login, 'password': password}
    
    try:
        token_data = doPostRequest(url, body, None)
        return token_data.get('token')
    except Exception as e:
        log_error("Failed to obtain token")
        return None       

#___________________________________________________________________

def EndOfDayOnExchanges(codes=None, codestype=None, boardIds=None, dateFrom=None, dateTo=None, official=False, pageNum=1, pageSize=100, max_size=20):
    #codes - список кодов инструментов
    #codestype - тип кодов -> codes:isins, issIds:issIds, fintoolIds:fintoolIds
    #boardIds - список кодов торговых площадок
    #dateFrom - дата начала '2024-01-01'
    #dateTo - дата оокнчания '2025-01-01'

    url = f"{api_url}/Archive/EndOfDayOnExchanges"
    
    codes_chunks = [codes[i:i + max_size] for i in range(0, len(codes), max_size)]
    
    token = getToken(api_login, api_pass)
    if not token:
        print("Failed to get token")
        return None
    
    all_chunk_pages = []
    try:
        for codes in codes_chunks:
            page_body = {
                codestype: codes,
                'dateFrom': dateFrom,
                'dateTo': dateTo,
                'boardIds': boardIds,
                'pageNum': pageNum,
                'pageSize': pageSize,
                'official': official,
                'fields': ['counter'],
            }
        
            data = doPostRequest(url, page_body, token)
            
            if data:
                page_number = math.ceil(data[0]['counter'] / pageSize)
            else:
                continue
        
            main_body = {
                codestype: codes,
                'dateFrom': dateFrom,
                'dateTo': dateTo,
                'boardIds': boardIds,
                'pageNum': pageNum,
                'pageSize': pageSize,
                'official': official,
                'fields': ['id', 'fintoolId', 'isin',
                           'time',
                           'boardid', 'boardname', 'currency',
                           'shortname_rus', 'seccode', 'name', 'secname',
                           'lclose', 'last', 'val_acc']
            }
        
            all_pages = []
            for page in tqdm(range(1, page_number + 1)):
                main_body['pageNum'] = page
                data = doPostRequest(url, main_body, token)
                if data:
                    all_pages.extend(data)
                else:
                    print("No data retrieved")
                    break
            all_chunk_pages.extend(all_pages)
    except Exception as e:
        log_error("Failed to retrieve data")
        print(f"Unexpected error during demo: {e}")

    if all_chunk_pages:
        all_chunk_pages = StringIO(json.dumps(all_chunk_pages))
        return pd.read_json(all_chunk_pages)
    else:
        return None

def CalendarV2(fintoolIds=None, startDate=None, endDate=None, eventType=None, pageNum=1, pageSize=1000, max_size=100):
    #codes - список кодов инструментов
    #startDate - дата начала '2024-01-01'
    #endDate - дата оокнчания '2025-01-01'
    #eventType - список событий -> CONV, CALL, CPN, MTY, DIV
    
    url = f"{api_url}/Info/CalendarV2"
    eventType = tuple(eventType) if len(eventType) > 1 else tuple(eventType) * 2
    
    fintoolIds_chunks = [fintoolIds[i:i + max_size] for i in range(0, len(fintoolIds), max_size)]
    
    token = getToken(api_login, api_pass)
    if not token:
        print("Failed to get token")
        return None

    all_chunk_pages = []
    try:
        for fintoolIds in fintoolIds_chunks:
            page_body = {
                'fintoolIds': fintoolIds,
                'startDate': startDate,
                'endDate': endDate,
                'pageNum': pageNum,
                'pageSize': pageSize,
                'fields': ['counter'],
                'filter': f'eventType IN {eventType}'
            }
    
            data = doPostRequest(url, page_body, token)
            
            if data:
                page_number = math.ceil(data[0]['counter'] / pageSize)
            else:
                continue
        
            main_body = {
                'fintoolIds': fintoolIds,
                'startDate': startDate,
                'endDate': endDate,
                'pageNum': pageNum,
                'pageSize': pageSize,
                'fields': ['fininstId', 'finToolID', 'id', 'isiNcode',
                           'faceFTName', 'nickname', 'coefficient',
                           'eventID', 'eventType', 'eventDate', 'beginConvDate', 'endConvDate'],
                'filter': f'eventType IN {eventType}'
            }
            
            all_pages = []
            for page in tqdm(range(1, page_number + 1)):
                main_body['pageNum'] = page
                data = doPostRequest(url, main_body, token)
                if data:
                    all_pages.extend(data)
                else:
                    print("No data retrieved")
                    break
            all_chunk_pages.extend(all_pages)
    except Exception as e:
        log_error("Failed to retrieve data")
        print(f"Unexpected error during demo: {e}")

    if all_chunk_pages:
        all_chunk_pages = StringIO(json.dumps(all_chunk_pages))
        return pd.read_json(all_chunk_pages)
    else:
        return None

def CurrencyRateHistory(baseCurrency=None, quotedCurrency=None, withHolidays=False, dateFrom=None, dateTo=None, pageNum=1, pageSize=1000):
    #baseCurrency - Базовая валюта, чей курс нужно узнать (трехбуквенный код)
    #quotedCurrency - Котируемая валюта, в которой нужно выразить курсы (трехбуквенный код)
    #dateFrom - дата начала '2024-01-01'
    #dateTo - дата оокнчания '2025-01-01'
    
    url = f"{api_url}/Archive/CurrencyRateHistory"
    
    token = getToken(api_login, api_pass)
    if not token:
        print("Failed to get token")
        return None

    all_chunk_pages = []
    try:
        page_body = {
            'baseCurrency': baseCurrency,
            'quotedCurrency': quotedCurrency,
            'withHolidays': withHolidays,
            'dateFrom': dateFrom,
            'dateTo': dateTo,
            'pageNum': pageNum,
            'pageSize': pageSize,
            'fields': ['counter'],
        }

        data = doPostRequest(url, page_body, token)
        
        if data:
            page_number = math.ceil(data[0]['counter'] / pageSize)
        else:
            pass
    
        main_body = {
            'baseCurrency': baseCurrency,
            'quotedCurrency': quotedCurrency,
            'withHolidays': withHolidays,
            'dateFrom': dateFrom,
            'dateTo': dateTo,
            'pageNum': pageNum,
            'pageSize': pageSize,
        }
        
        all_pages = []
        for page in tqdm(range(1, page_number + 1)):
            main_body['pageNum'] = page
            data = doPostRequest(url, main_body, token)
            if data:
                all_pages.extend(data)
            else:
                print("No data retrieved")
                break
        all_chunk_pages.extend(all_pages)
    except Exception as e:
        log_error("Failed to retrieve data")
        print(f"Unexpected error during demo: {e}")

    if all_chunk_pages:
        all_chunk_pages = StringIO(json.dumps(all_chunk_pages))
        return pd.read_json(all_chunk_pages)
    else:
        return None

def Instruments(codes=None, codestype=None, pageNum=1, pageSize=300, max_size=100):
    #codes - список кодов инструментов
    #codestype - тип кодов -> id:id, fintoolId:fintoolId, isin:isin, seccode:seccode
    
    url = f"{api_url}/Info/Instruments"
    codes = tuple(codes) if len(codes) > 1 else tuple(codes) * 2
    
    codes_chunks = [codes[i:i + max_size] for i in range(0, len(codes), max_size)]
    
    token = getToken(api_login, api_pass)
    if not token:
        print("Failed to get token")
        return None

    all_chunk_pages = []
    try:
        for codes in codes_chunks:
            page_body = {
                'pageNum': pageNum,
                'pageSize': pageSize,
                #"filter": f"name like '%{codes}%'"
                "filter": f"{codestype} IN {codes}"
            }
            
            data = doPostRequest(url, page_body, token)
            
            if data:
                page_number = math.ceil(data[0]['counter'] / pageSize)
            else:
                continue
        
            main_body = {
                'pageNum': pageNum,
                'pageSize': pageSize,
                #"filter": f"name like '%{codes}%'"
                "filter": f"{codestype} IN {codes}"
            }
                
            all_pages = []
            for page in tqdm(range(1, page_number + 1)):
                main_body['pageNum'] = page
                data = doPostRequest(url, main_body, token)
                if data:
                    all_pages.extend(data)
                else:
                    print("No data retrieved")
                    break
            all_chunk_pages.extend(all_pages)
    except Exception as e:
        log_error("Failed to retrieve data")
        print(f"Unexpected error during demo: {e}")

    if all_chunk_pages:
        all_chunk_pages = StringIO(json.dumps(all_chunk_pages))
        return pd.read_json(all_chunk_pages)
    else:
        return None

def Currencies(pageNum=1, pageSize=500):   
    url = f"{api_url}/Info/Currencies"
        
    token = getToken(api_login, api_pass)
    if not token:
        print("Failed to get token")
        return None

    all_chunk_pages = []
    try:
        page_body = {
            'pageNum': pageNum,
            'pageSize': pageSize
        }

        data = doPostRequest(url, page_body, token)
        
        if data:
            page_number = math.ceil(data[0]['counter'] / pageSize)
        else:
            pass
    
        main_body = {
            'pageNum': pageNum,
            'pageSize': pageSize
        }
            
        all_pages = []
        for page in tqdm(range(1, page_number + 1)):
            main_body['pageNum'] = page
            data = doPostRequest(url, main_body, token)
            if data:
                all_pages.extend(data)
            else:
                print("No data retrieved")
                break
        all_chunk_pages.extend(all_pages)
    except Exception as e:
        log_error("Failed to retrieve data")
        print(f"Unexpected error during demo: {e}")

    if all_chunk_pages:
        all_chunk_pages = StringIO(json.dumps(all_chunk_pages))
        return pd.read_json(all_chunk_pages)
    else:
        return None
