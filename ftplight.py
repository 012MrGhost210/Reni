#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
РАБОЧИЙ СКРИПТ ДЛЯ КОПИРОВАНИЯ ФАЙЛОВ С FTP
"""

import os
import sys
import ftplib
from ftplib import FTP

def main():
    # ====== ВАШИ ДАННЫЕ ======
    FTP_HOST = "ftp.renlife.com"
    FTP_USER = "Ilya.Matveev2@mos.renlife.com"
    FTP_PASS = "кенгуруру"
    FTP_FOLDER = "/diadoc_connector"
    LOCAL_FOLDER = r"M:\Инвестиционный департамент\7.0 Treasury\Diadoc"
    # =========================
    
    print("=" * 70)
    print("КОПИРОВАНИЕ ФАЙЛОВ С FTP")
    print("=" * 70)
    print(f"Сервер: {FTP_HOST}")
    print(f"Папка на FTP: {FTP_FOLDER}")
    print(f"Сохранить в: {LOCAL_FOLDER}")
    print("-" * 70)
    
    # Настраиваем кодировку для Windows
    if sys.platform == 'win32':
        os.system('chcp 65001 > nul')
    
    # Создаем папку для сохранения
    try:
        os.makedirs(LOCAL_FOLDER, exist_ok=True)
        print(f"✅ Создана папка: {LOCAL_FOLDER}")
    except Exception as e:
        print(f"❌ Ошибка создания папки: {e}")
        input("Нажмите Enter...")
        return
    
    # Подключаемся к FTP с разными кодировками
    ftp = None
    try:
        print(f"🔌 Подключаюсь к {FTP_HOST}...")
        
        # Пробуем разные кодировки
        for encoding in ['utf-8', 'cp1251', 'cp866', None]:
            try:
                ftp = FTP(FTP_HOST, timeout=30)
                if encoding:
                    ftp.encoding = encoding
                
                # Пробуем войти
                ftp.login(FTP_USER, FTP_PASS)
                ftp.set_pasv(True)
                
                print(f"✅ Подключение успешно! Кодировка: {encoding if encoding else 'default'}")
                break
                
            except Exception as e:
                if ftp:
                    try:
                        ftp.quit()
                    except:
                        pass
                ftp = None
                continue
        
        if ftp is None:
            print("❌ Не удалось подключиться с любой кодировкой")
            input("Нажмите Enter...")
            return
            
    except Exception as e:
        print(f"❌ Ошибка подключения: {e}")
        input("Нажмите Enter...")
        return
    
    # Переходим в нужную папку
    try:
        print(f"📂 Перехожу в папку: {FTP_FOLDER}")
        ftp.cwd(FTP_FOLDER)
    except Exception as e:
        print(f"❌ Не могу перейти в папку: {e}")
        ftp.quit()
        input("Нажмите Enter...")
        return
    
    # Функция для копирования с обработкой ошибок
    def safe_retrbinary(ftp, filename, fileobj, blocksize=8192):
        """Безопасное скачивание файла"""
        try:
            ftp.retrbinary(f'RETR {filename}', fileobj.write, blocksize)
            return True
        except ftplib.error_perm as e:
            print(f"    ⚠️  Ошибка доступа: {e}")
            return False
        except Exception as e:
            print(f"    ⚠️  Другая ошибка: {e}")
            return False
    
    # Получаем список файлов
    try:
        print("📄 Получаю список файлов...")
        items = ftp.nlst()
        print(f"✅ Найдено элементов: {len(items)}")
    except Exception as e:
        print(f"❌ Ошибка получения списка: {e}")
        ftp.quit()
        input("Нажмите Enter...")
        return
    
    # Основной цикл копирования
    success = 0
    failed = 0
    
    print("\n" + "=" * 70)
    print("НАЧИНАЮ КОПИРОВАНИЕ...")
    print("=" * 70)
    
    for item in items:
        if item in [".", ".."]:
            continue
        
        # Пробуем разные декодирования для имени файла
        filename_display = str(item)
        
        # Пробуем декодировать если это bytes
        if isinstance(item, bytes):
            for encoding in ['utf-8', 'cp1251', 'cp866', 'iso-8859-1']:
                try:
                    filename_display = item.decode(encoding)
                    break
                except:
                    continue
        
        print(f"\n📝 Обрабатываю: {filename_display}")
        
        # Пробуем определить, это файл или папка
        try:
            # Пробуем получить размер файла
            try:
                size = ftp.size(item)
            except:
                size = None
            
            if size is not None:  # Это файл
                print(f"   Размер: {size} байт")
                
                # Создаем безопасное имя файла
                safe_name = filename_display
                
                # Заменяем недопустимые символы
                invalid_chars = '<>:"/\\|?*'
                for char in invalid_chars:
                    safe_name = safe_name.replace(char, '_')
                
                # Полный путь для сохранения
                local_path = os.path.join(LOCAL_FOLDER, safe_name)
                
                try:
                    # Скачиваем файл
                    with open(local_path, 'wb') as f:
                        if safe_retrbinary(ftp, item, f):
                            print(f"   ✅ Успешно сохранен как: {safe_name}")
                            success += 1
                        else:
                            print(f"   ❌ Не удалось скачать")
                            failed += 1
                            
                except Exception as e:
                    print(f"   ❌ Ошибка файловой системы: {e}")
                    failed += 1
                    
            else:  # Возможно папка
                print(f"   ⚠️  Пропускаю (вероятно папка)")
                
        except Exception as e:
            print(f"   ❌ Ошибка обработки: {e}")
            failed += 1
    
    # Закрываем соединение
    try:
        ftp.quit()
        print("\n🔌 Соединение закрыто")
    except:
        pass
    
    # Итоги
    print("\n" + "=" * 70)
    print("КОПИРОВАНИЕ ЗАВЕРШЕНО!")
    print("=" * 70)
    print(f"📊 ИТОГО:")
    print(f"   ✅ Успешно: {success} файлов")
    print(f"   ❌ Ошибок: {failed} файлов")
    print(f"   📂 Папка: {LOCAL_FOLDER}")
    print("=" * 70)
    
    input("\nНажмите Enter для выхода...")

if __name__ == "__main__":
    main()
