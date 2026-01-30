#!/usr/bin/env python3
"""
Скрипт для рекурсивного копирования файлов с FTP сервера с сохранением структуры папок
"""

from ftplib import FTP
import os
import sys
from datetime import datetime
import json

class FTPRecursiveDownloader:
    def __init__(self, config_file='ftp_config.json'):
        self.config = self.load_config(config_file)
        self.ftp = None
        self.downloaded_files = 0
        self.downloaded_dirs = 0
        
    def load_config(self, config_file):
        """Загружает конфигурацию из JSON файла"""
        if not os.path.exists(config_file):
            example_config = {
                "host": "ftp.example.com",
                "port": 21,
                "username": "your_username",
                "password": "your_password",
                "remote_base_dir": "/",  # Базовая директория на сервере
                "local_base_dir": "./ftp_backup",  # Локальная базовая директория
                "exclude_dirs": [".", ".."],  # Директории для исключения
                "file_pattern": "*",  # Шаблон файлов
                "preserve_permissions": False,  # Сохранять разрешения (только Unix)
                "skip_existing": True,  # Пропускать уже существующие файлы
                "max_depth": None,  # Максимальная глубина рекурсии
                "use_tls": False,
                "passive_mode": True
            }
            with open(config_file, 'w') as f:
                json.dump(example_config, f, indent=2, ensure_ascii=False)
            print(f"Создан пример конфигурационного файла: {config_file}")
            print("Пожалуйста, заполните его своими данными.")
            sys.exit(1)
        
        with open(config_file, 'r') as f:
            return json.load(f)
    
    def connect(self):
        """Подключается к FTP серверу"""
        try:
            print(f"Подключение к {self.config['host']}:{self.config.get('port', 21)}...")
            self.ftp = FTP()
            self.ftp.connect(self.config['host'], self.config.get('port', 21))
            self.ftp.login(self.config['username'], self.config['password'])
            
            if self.config.get('passive_mode', True):
                self.ftp.set_pasv(True)
            
            print("✓ Подключение успешно!")
            return True
            
        except Exception as e:
            print(f"✗ Ошибка подключения: {e}")
            return False
    
    def is_directory(self, item):
        """Проверяет, является ли элемент директорией"""
        try:
            # Сохраняем текущую директорию
            original_dir = self.ftp.pwd()
            
            # Пробуем перейти в элемент
            self.ftp.cwd(item)
            # Если получилось, возвращаемся назад
            self.ftp.cwd(original_dir)
            return True
        except:
            return False
    
    def get_recursive_listing(self, remote_dir=".", depth=0):
        """
        Рекурсивно получает список всех файлов и папок
        Возвращает список словарей с информацией о каждом элементе
        """
        items = []
        
        try:
            # Переходим в директорию
            self.ftp.cwd(remote_dir)
            
            # Получаем список элементов в текущей директории
            for item in self.ftp.nlst():
                # Пропускаем специальные директории
                if item in self.config.get('exclude_dirs', [".", ".."]):
                    continue
                
                full_path = os.path.join(remote_dir, item).replace("\\", "/")
                
                # Проверяем, является ли директорией
                if self.is_directory(item):
                    items.append({
                        'type': 'directory',
                        'name': item,
                        'path': full_path,
                        'depth': depth
                    })
                    
                    # Проверяем максимальную глубину рекурсии
                    max_depth = self.config.get('max_depth')
                    if max_depth is None or depth < max_depth:
                        # Рекурсивно получаем содержимое поддиректории
                        sub_items = self.get_recursive_listing(full_path, depth + 1)
                        items.extend(sub_items)
                else:
                    # Это файл
                    try:
                        size = self.ftp.size(item)
                        items.append({
                            'type': 'file',
                            'name': item,
                            'path': full_path,
                            'size': size,
                            'depth': depth
                        })
                    except:
                        # Если не удалось получить размер
                        items.append({
                            'type': 'file',
                            'name': item,
                            'path': full_path,
                            'size': 0,
                            'depth': depth
                        })
            
            # Возвращаемся на уровень выше
            if remote_dir != ".":
                self.ftp.cwd("..")
                
        except Exception as e:
            print(f"Ошибка при сканировании {remote_dir}: {e}")
        
        return items
    
    def create_local_dir(self, remote_path):
        """Создает локальную директорию, соответствующую удаленной"""
        # Преобразуем удаленный путь в локальный
        remote_base = self.config['remote_base_dir'].rstrip('/')
        local_base = self.config['local_base_dir']
        
        # Убираем базовую директорию из пути
        relative_path = remote_path[len(remote_base):] if remote_path.startswith(remote_base) else remote_path
        if relative_path.startswith('/'):
            relative_path = relative_path[1:]
        
        # Собираем полный локальный путь
        local_path = os.path.join(local_base, relative_path)
        
        # Создаем директорию
        os.makedirs(local_path, exist_ok=True)
        
        return local_path
    
    def download_file(self, remote_file_path, local_dir):
        """Скачивает один файл"""
        try:
            # Получаем имя файла из пути
            filename = os.path.basename(remote_file_path)
            local_path = os.path.join(local_dir, filename)
            
            # Проверяем, нужно ли пропускать существующие файлы
            if self.config.get('skip_existing', True) and os.path.exists(local_path):
                print(f"  [ПРОПУСК] Файл уже существует: {filename}")
                return False
            
            # Получаем размер файла
            file_size = self.ftp.size(remote_file_path)
            
            print(f"  ↓ Скачиваю: {filename} ({self.format_size(file_size)})")
            
            # Скачиваем файл
            with open(local_path, 'wb') as f:
                self.ftp.retrbinary(f'RETR {remote_file_path}', f.write)
            
            self.downloaded_files += 1
            return True
            
        except Exception as e:
            print(f"  ✗ Ошибка при скачивании {remote_file_path}: {e}")
            return False
    
    def format_size(self, size_bytes):
        """Форматирует размер файла в читаемом виде"""
        if size_bytes is None:
            return "неизвестно"
        
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.1f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.1f} TB"
    
    def sync_structure(self):
        """Синхронизирует полную структуру папок и файлов"""
        if not self.ftp:
            print("Нет подключения к FTP!")
            return
        
        print("\n" + "="*60)
        print("НАЧАЛО СИНХРОНИЗАЦИИ СТРУКТУРЫ ПАПОК")
        print("="*60)
        
        try:
            # Переходим в базовую директорию на сервере
            remote_base = self.config['remote_base_dir']
            if remote_base != "/":
                self.ftp.cwd(remote_base)
                print(f"Базовая директория на сервере: {remote_base}")
            
            # Создаем локальную базовую директорию
            os.makedirs(self.config['local_base_dir'], exist_ok=True)
            
            # Получаем полную структуру директорий и файлов
            print("\nСканирую структуру сервера...")
            structure = self.get_recursive_listing(".", depth=0)
            
            # Сначала создаем все директории
            print("\nСоздаю структуру папок...")
            for item in structure:
                if item['type'] == 'directory':
                    local_dir = self.create_local_dir(item['path'])
                    indent = "  " * item['depth']
                    print(f"{indent}📁 Создана папка: {item['name']}")
                    self.downloaded_dirs += 1
            
            # Затем скачиваем все файлы
            print("\nСкачиваю файлы...")
            for item in structure:
                if item['type'] == 'file':
                    # Получаем директорию файла
                    remote_dir = os.path.dirname(item['path'])
                    local_dir = self.create_local_dir(remote_dir)
                    
                    # Скачиваем файл
                    indent = "  " * item['depth']
                    print(f"{indent}", end="")
                    self.download_file(item['path'], local_dir)
            
            # Отчет
            print("\n" + "="*60)
            print("СИНХРОНИЗАЦИЯ ЗАВЕРШЕНА")
            print("="*60)
            print(f"Обработано директорий: {self.downloaded_dirs}")
            print(f"Скачано файлов: {self.downloaded_files}")
            print(f"Локальная копия: {os.path.abspath(self.config['local_base_dir'])}")
            
        except Exception as e:
            print(f"\n✗ Ошибка при синхронизации: {e}")
    
    def disconnect(self):
        """Закрывает соединение"""
        if self.ftp:
            self.ftp.quit()
            print("\nСоединение с FTP сервером закрыто.")

def main():
    """Основная функция"""
    print("FTP Recursive Downloader v1.0")
    print("="*60)
    
    # Создаем загрузчик
    downloader = FTPRecursiveDownloader('ftp_config.json')
    
    # Подключаемся и синхронизируем
    if downloader.connect():
        try:
            downloader.sync_structure()
        finally:
            downloader.disconnect()

if __name__ == "__main__":
    main()
