#!/usr/bin/env python3
"""
СКРИПТ ДЛЯ ПОЛНОГО КОПИРОВАНИЯ С FTP СЕРВЕРА НА ДИСК C
Копирует ВСЕ файлы и папки с сохранением структуры
"""

from ftplib import FTP
import os
import sys
import time
from pathlib import Path

def is_directory(ftp, name):
    """Проверяет, является ли элемент директорией"""
    try:
        original_dir = ftp.pwd()
        ftp.cwd(name)
        ftp.cwd(original_dir)
        return True
    except:
        return False

def copy_all_from_ftp_to_c():
    """
    Копирует ВСЕ с FTP сервера в C:\ftp_backup\
    """
    
    # Параметры подключения (ЗАМЕНИТЕ НА СВОИ!)
    FTP_HOST = "ftp.renlife.com"      # Адрес FTP сервера
    FTP_USER = "Ilya.Matveev2@mos.renlife.com"               # Логин
    FTP_PASS = "@$CiaG3008"              # Пароль
    REMOTE_DIR = "/diadoc_connector"                     # Копируем с корня FTP
    LOCAL_DIR = r"M:\Инвестиционный департамент\7.0 Treasury\Diadoc"         # Куда копируем на диске C
    
    print("="*70)
    print("ПОЛНОЕ КОПИРОВАНИЕ С FTP СЕРВЕРА НА ДИСК C")
    print("="*70)
    print(f"📡 Сервер: {FTP_HOST}")
    print(f"👤 Пользователь: {FTP_USER}")
    print(f"📂 Источник: {REMOTE_DIR}")
    print(f"💾 Назначение: {LOCAL_DIR}")
    print("-"*70)
    
    # Проверяем, существует ли диск C
    if not os.path.exists('C:'):
        print("❌ ОШИБКА: Диск C: не найден!")
        input("Нажмите Enter для выхода...")
        sys.exit(1)
    
    # Создаем папку на диске C
    try:
        os.makedirs(LOCAL_DIR, exist_ok=True)
        print(f"✅ Создана папка на диске C: {LOCAL_DIR}")
    except Exception as e:
        print(f"❌ Не могу создать папку на диске C: {e}")
        input("Нажмите Enter для выхода...")
        sys.exit(1)
    
    # Подключаемся к FTP
    try:
        print("\n🔌 Подключаюсь к FTP серверу...")
        ftp = FTP(FTP_HOST, timeout=60)
        ftp.login(FTP_USER, FTP_PASS)
        ftp.set_pasv(True)  # Пассивный режим (лучше работает с фаерволами)
        print("✅ Подключение успешно!")
    except Exception as e:
        print(f"❌ Ошибка подключения к FTP: {e}")
        input("Нажмите Enter для выхода...")
        sys.exit(1)
    
    # Счетчики
    total_files = 0
    total_dirs = 0
    start_time = time.time()
    
    def recursive_copy(remote_path, local_path, depth=0):
        """Рекурсивно копирует файлы и папки"""
        nonlocal total_files, total_dirs
        
        try:
            # Переходим в удаленную папку
            ftp.cwd(remote_path)
            
            # Создаем локальную папку
            os.makedirs(local_path, exist_ok=True)
            
            # Получаем список элементов
            items = ftp.nlst()
            
            for item in items:
                if item in [".", ".."]:
                    continue
                
                remote_item = f"{remote_path}/{item}" if remote_path != "/" else f"/{item}"
                local_item = os.path.join(local_path, item)
                
                # Отступ для красивого вывода
                indent = "  " * depth
                
                if is_directory(ftp, item):
                    # Это папка
                    print(f"{indent}📁 ПАПКА: {item}")
                    total_dirs += 1
                    
                    # Рекурсивно копируем содержимое папки
                    recursive_copy(remote_item, local_item, depth + 1)
                    
                else:
                    # Это файл
                    try:
                        # Получаем размер файла
                        file_size = ftp.size(item)
                        size_str = f"({file_size} байт)" if file_size else ""
                        
                        print(f"{indent}📄 ФАЙЛ: {item} {size_str}")
                        
                        # Скачиваем файл
                        with open(local_item, 'wb') as f:
                            ftp.retrbinary(f'RETR {item}', f.write)
                        
                        total_files += 1
                        
                    except Exception as e:
                        print(f"{indent}❌ Ошибка файла {item}: {e}")
            
            # Возвращаемся на уровень выше
            if remote_path != "/":
                ftp.cwd("..")
                
        except Exception as e:
            print(f"❌ Ошибка в папке {remote_path}: {e}")
    
    # Начинаем копирование
    print("\n🚀 НАЧИНАЮ КОПИРОВАНИЕ...")
    print("-"*70)
    
    try:
        recursive_copy(REMOTE_DIR, LOCAL_DIR)
        
        # Выводим итоги
        elapsed_time = time.time() - start_time
        print("\n" + "="*70)
        print("✅ КОПИРОВАНИЕ ЗАВЕРШЕНО!")
        print("="*70)
        print(f"📁 Создано папок: {total_dirs}")
        print(f"📄 Скачано файлов: {total_files}")
        print(f"⏱️  Затрачено времени: {elapsed_time:.1f} секунд")
        print(f"💾 Сохранено в: {LOCAL_DIR}")
        print(f"📊 Занято места: {get_folder_size(LOCAL_DIR):.2f} МБ")
        print("="*70)
        
    except Exception as e:
        print(f"\n❌ КРИТИЧЕСКАЯ ОШИБКА: {e}")
    
    finally:
        # Закрываем соединение
        ftp.quit()
        print("\n🔌 Соединение с FTP сервером закрыто.")
    
    input("\nНажмите Enter для выхода...")

def get_folder_size(folder_path):
    """Вычисляет размер папки в мегабайтах"""
    total_size = 0
    for dirpath, dirnames, filenames in os.walk(folder_path):
        for f in filenames:
            fp = os.path.join(dirpath, f)
            if os.path.exists(fp):
                total_size += os.path.getsize(fp)
    return total_size / (1024 * 1024)  # В МБ

if __name__ == "__main__":
    # Автоматически запускаем копирование
    copy_all_from_ftp_to_c()
