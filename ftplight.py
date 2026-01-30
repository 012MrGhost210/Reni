#!/usr/bin/env python3
"""
Скрипт для копирования файлов с FTP сервера
"""

from ftplib import FTP
import os
from datetime import datetime
import sys

def download_files_from_ftp(host, username, password, remote_dir, local_dir, file_pattern=None):
    """
    Скачивает файлы с FTP сервера
    
    Параметры:
    - host: адрес FTP сервера
    - username: имя пользователя
    - password: пароль
    - remote_dir: удаленная директория на сервере
    - local_dir: локальная директория для сохранения
    - file_pattern: шаблон имени файлов для фильтрации (например, '*.txt')
    """
    
    # Создаем локальную директорию если ее нет
    os.makedirs(local_dir, exist_ok=True)
    
    try:
        # Подключаемся к FTP серверу
        print(f"Подключение к {host}...")
        ftp = FTP(host)
        ftp.login(username, password)
        print("Успешное подключение!")
        
        # Переходим в удаленную директорию
        ftp.cwd(remote_dir)
        print(f"Переход в директорию: {remote_dir}")
        
        # Получаем список файлов
        files = ftp.nlst()
        print(f"Найдено файлов: {len(files)}")
        
        # Фильтруем файлы если указан шаблон
        if file_pattern:
            import fnmatch
            files = [f for f in files if fnmatch.fnmatch(f, file_pattern)]
            print(f"Файлов по шаблону '{file_pattern}': {len(files)}")
        
        # Скачиваем файлы
        downloaded = 0
        for filename in files:
            local_path = os.path.join(local_dir, filename)
            
            # Проверяем, является ли это файлом (а не директорией)
            try:
                file_size = ftp.size(filename)
                if file_size is not None:  # FTP возвращает None для директорий
                    print(f"Скачивание: {filename} ({file_size} байт)")
                    
                    with open(local_path, 'wb') as f:
                        ftp.retrbinary(f'RETR {filename}', f.write)
                    
                    downloaded += 1
                    print(f"✓ Скачан: {filename}")
            except:
                print(f"Пропуск (возможно директория): {filename}")
        
        # Закрываем соединение
        ftp.quit()
        print(f"\nЗавершено! Скачано файлов: {downloaded}")
        
    except Exception as e:
        print(f"Ошибка: {e}")
        sys.exit(1)

if __name__ == "__main__":
    # Конфигурация подключения (замените на свои данные)
    FTP_CONFIG = {
        'host': 'ftp.example.com',
        'username': 'your_username',
        'password': 'your_password',
        'remote_dir': '/path/to/remote/files',
        'local_dir': './downloaded_files',
        'file_pattern': None  # или '*.txt', '*.csv', 'data_*.zip'
    }
    
    # Использование:
    download_files_from_ftp(**FTP_CONFIG)
