import os
import sys
from ftplib import FTP
import locale

def main():
    # ====== НАСТРОЙКИ ======
    FTP_HOST = "ftp.renlife.com"      # например: 192.168.1.100
    FTP_USER = "Ilya.Matveev2@mos.renlife.com"           # ваш логин
    FTP_PASS = "@$CiaG3008"          # ваш пароль
    FTP_FOLDER = "/diadoc_connector"                 # папка на FTP (начинается с /)
    LOCAL_FOLDER = r"M:\Инвестиционный департамент\7.0 Treasury\Diadoc"      # куда копировать на M: диске
    # =======================
    
    print("=" * 70)
    print("КОПИРОВАНИЕ ФАЙЛОВ С FTP СЕРВЕРА")
    print("=" * 70)
    
    # Настраиваем кодировку для Windows
    if sys.platform == 'win32':
        import ctypes
        # Устанавливаем кодировку консоли в UTF-8
        if sys.version_info >= (3, 7):
            os.system('chcp 65001 > nul')
    
    # Создаем папку для сохранения
    try:
        os.makedirs(LOCAL_FOLDER, exist_ok=True)
        print(f"📁 Папка для сохранения: {LOCAL_FOLDER}")
    except Exception as e:
        print(f"❌ Ошибка создания папки: {e}")
        input("Нажмите Enter...")
        return
    
    # Подключаемся к FTP
    try:
        print(f"🔌 Подключаюсь к {FTP_HOST}...")
        ftp = FTP(FTP_HOST, timeout=60)
        
        # Пробуем разные кодировки для логина
        try:
            ftp.login(FTP_USER, FTP_PASS)
        except:
            # Если не работает с кодировкой по умолчанию, пробуем UTF-8
            ftp.encoding = 'utf-8'
            ftp.login(FTP_USER.encode('utf-8'), FTP_PASS.encode('utf-8'))
        
        ftp.set_pasv(True)  # Важно для Windows!
        print("✅ Подключение успешно!")
        
    except Exception as e:
        print(f"❌ Ошибка подключения: {e}")
        input("Нажмите Enter...")
        return
    
    # Переходим в нужную папку
    try:
        print(f"📂 Перехожу в папку FTP: {FTP_FOLDER}")
        ftp.cwd(FTP_FOLDER)
    except Exception as e:
        print(f"❌ Не могу перейти в папку: {e}")
        print("Пробую корневую папку...")
        try:
            ftp.cwd("/")
            FTP_FOLDER = "/"
        except:
            print("Не могу получить доступ к FTP")
            ftp.quit()
            input("Нажмите Enter...")
            return
    
    # Функция для правильного преобразования имен файлов
    def safe_filename(name):
        """Преобразует имя файла в безопасное для Windows"""
        # Заменяем недопустимые символы
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            name = name.replace(char, '_')
        # Убираем пробелы в начале и конце
        name = name.strip()
        # Ограничиваем длину
        if len(name) > 200:
            name = name[:200]
        return name
    
    # Получаем список файлов
    try:
        print("📄 Получаю список файлов...")
        
        # Пробуем разные методы получения списка файлов
        try:
            files = ftp.nlst()
        except:
            # Альтернативный метод с другой кодировкой
            ftp.encoding = 'cp1251'
            files = ftp.nlst()
        
        print(f"✅ Найдено элементов: {len(files)}")
        
    except Exception as e:
        print(f"❌ Ошибка получения списка файлов: {e}")
        ftp.quit()
        input("Нажмите Enter...")
        return
    
    # Копируем файлы
    success = 0
    failed = 0
    
    print("\n" + "=" * 70)
    print("НАЧИНАЮ КОПИРОВАНИЕ...")
    print("=" * 70)
    
    for item in files:
        if item in [".", ".."]:
            continue
        
        # Пробуем разные кодировки для имени файла
        filename = item
        try:
            # Пробуем UTF-8
            if isinstance(filename, bytes):
                filename = filename.decode('utf-8')
        except:
            try:
                # Пробуем cp1251 (Windows)
                if isinstance(filename, bytes):
                    filename = filename.decode('cp1251')
            except:
                # Оставляем как есть
                pass
        
        # Создаем безопасное имя файла
        safe_name = safe_filename(filename)
        
        # Полный путь для сохранения
        local_path = os.path.join(LOCAL_FOLDER, safe_name)
        
        print(f"\n📝 Обрабатываю: {filename}")
        print(f"   Сохраняю как: {safe_name}")
        
        # Проверяем, файл это или папка
        try:
            # Пробуем получить размер файла
            size = ftp.size(filename)
            if size is not None:  # Это файл
                print(f"   Размер: {size} байт")
                
                try:
                    with open(local_path, 'wb') as f:
                        # Скачиваем файл
                        ftp.retrbinary(f'RETR {filename}', f.write)
                    print(f"   ✅ УСПЕШНО скопирован")
                    success += 1
                except Exception as e:
                    print(f"   ❌ Ошибка скачивания: {e}")
                    failed += 1
                    
                    # Пробуем альтернативный метод
                    try:
                        print("   🔄 Пробую альтернативный метод...")
                        ftp.voidcmd('TYPE I')  # Binary mode
                        with open(local_path, 'wb') as f:
                            def callback(data):
                                f.write(data)
                            ftp.retrbinary(f'RETR {filename}', callback)
                        print(f"   ✅ УСПЕШНО (альтернативный метод)")
                        success += 1
                    except:
                        print(f"   ❌ Не удалось скачать файл")
                        failed += 1
            else:
                print(f"   ⚠️  Пропускаю (папка)")
        except:
            print(f"   ⚙️  Пропускаю элемент")
    
    # Закрываем соединение
    ftp.quit()
    
    # Итоги
    print("\n" + "=" * 70)
    print("КОПИРОВАНИЕ ЗАВЕРШЕНО!")
    print("=" * 70)
    print(f"📊 РЕЗУЛЬТАТЫ:")
    print(f"   ✅ Успешно скопировано: {success} файлов")
    print(f"   ❌ Не удалось скопировать: {failed} файлов")
    print(f"   📂 Сохранено в: {LOCAL_FOLDER}")
    print("=" * 70)
    
    input("\nНажмите Enter для выхода...")

if __name__ == "__main__":
    main()
