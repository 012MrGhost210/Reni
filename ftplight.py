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

        
#!/usr/bin/env python3
"""
ПРОСТОЙ МЕТОД - копирует всю папку FTP в локалку
"""

import os
from ftplib import FTP

def copy_ftp_directory(ftp_host, ftp_user, ftp_pass, remote_dir="/", local_dir="C:/FTP_COPY"):
    """
    Копирует всю папку с FTP сервера в локальную папку
    
    Аргументы:
    ftp_host - адрес FTP сервера
    ftp_user - логин
    ftp_pass - пароль
    remote_dir - папка на FTP (по умолчанию корень /)
    local_dir - куда копировать локально (по умолчанию C:/FTP_COPY)
    """
    
    print(f"Копирую {remote_dir} с FTP -> {local_dir}")
    
    # Создаем локальную папку
    os.makedirs(local_dir, exist_ok=True)
    
    # Подключаемся к FTP
    ftp = FTP(ftp_host)
    ftp.login(ftp_user, ftp_pass)
    ftp.set_pasv(True)  # Важно для Windows
    
    # Внутренняя функция для рекурсивного копирования
    def copy_current_dir(ftp_path, local_path):
        """Копирует текущую директорию"""
        # Переходим в папку на FTP
        ftp.cwd(ftp_path)
        
        # Получаем список всего в текущей папке
        items = ftp.nlst()
        
        for item in items:
            if item in [".", ".."]:
                continue
            
            # Пробуем определить, это папка или файл
            try:
                # Сохраняем текущую позицию
                original_dir = ftp.pwd()
                
                # Пробуем перейти в элемент
                try:
                    ftp.cwd(item)
                    # Если получилось - это папка
                    
                    # Создаем локальную папку
                    local_item_path = os.path.join(local_path, item)
                    os.makedirs(local_item_path, exist_ok=True)
                    
                    # Рекурсивно копируем папку
                    print(f"Папка: {item}")
                    copy_current_dir(f"{ftp_path}/{item}", local_item_path)
                    
                    # Возвращаемся назад
                    ftp.cwd("..")
                    
                except:
                    # Не получилось перейти - это файл
                    local_item_path = os.path.join(local_path, item)
                    
                    print(f"Файл: {item}")
                    
                    # Скачиваем файл
                    with open(local_item_path, 'wb') as f:
                        ftp.retrbinary(f'RETR {item}', f.write)
                        
            except Exception as e:
                print(f"Ошибка при обработке {item}: {e}")
                continue
    
    # Запускаем копирование
    try:
        copy_current_dir(remote_dir, local_dir)
        print(f"\n✅ Успешно скопировано в: {local_dir}")
    except Exception as e:
        print(f"\n❌ Ошибка: {e}")
    finally:
        # Всегда закрываем соединение
        ftp.quit()

# ========== ПРИМЕР ИСПОЛЬЗОВАНИЯ ==========
if __name__ == "__main__":
    # ВАШИ ДАННЫЕ FTP
    FTP_HOST = "ftp.renlife.com"    # Например: 192.168.1.100 или ftp.site.com
    FTP_USER = "Ilya.Matveev2@mos.renlife.com"         # Ваш логин для FTP
    FTP_PASS = "@$CiaG3008"        # Ваш пароль
    
    # Откуда копировать на FTP (обычно / для всей папки)
    REMOTE_PATH = "/diadoc_connector"
    
    # Куда копировать локально
    LOCAL_PATH = r"M:\Инвестиционный департамент\7.0 Treasury\Diadoc"  # Или "C:/FTP_COPY"
    
    # Запускаем копирование
    copy_ftp_directory(
        ftp_host=FTP_HOST,
        ftp_user=FTP_USER,
        ftp_pass=FTP_PASS,
        remote_dir=REMOTE_PATH,
        local_dir=LOCAL_PATH
    )
    
    input("\nНажмите Enter для выхода...")


брабатываю: ???????????
   Сохраняю как: ___________
   ⚙️  Пропускаю элемент

📝 Обрабатываю: ??????????????? ????????
   Сохраняю как: _______________ ________
   ⚙️  Пропускаю элемент

📝 Обрабатываю: ????????
   Сохраняю как: ________
   ⚙️  Пропускаю элемент

📝 Обрабатываю: ????????? ?????????????
   Сохраняю как: _________ _____________
   ⚙️  Пропускаю элемент

📝 Обрабатываю: ????????? ???????
   Сохраняю как: _________ _______
   ⚙️  Пропускаю элемент
Traceback (most recent call last):
  File "c:/Users/Ilya.Matveev2/Скрипты/Files redactor/ftp light.py", line 193, in <module>
    main()
  File "c:/Users/Ilya.Matveev2/Скрипты/Files redactor/ftp light.py", line 178, in main
    ftp.quit()
  File "C:\Program Files\Python38\lib\ftplib.py", line 655, in quit
    resp = self.voidcmd('QUIT')
  File "C:\Program Files\Python38\lib\ftplib.py", line 282, in voidcmd
    return self.voidresp()
  File "C:\Program Files\Python38\lib\ftplib.py", line 255, in voidresp
    resp = self.getresp()
  File "C:\Program Files\Python38\lib\ftplib.py", line 240, in getresp
    resp = self.getmultiline()
  File "C:\Program Files\Python38\lib\ftplib.py", line 226, in getmultiline
    line = self.getline()
  File "C:\Program Files\Python38\lib\ftplib.py", line 214, in getline
    raise EOFError
EOFError
