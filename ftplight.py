import os
from ftplib import FTP

# ==== ВАШИ ДАННЫЕ ====
FTP_SERVER = "ftp.renlife.com"  # или IP адрес
FTP_USER = "Ilya.Matveev2@mos.renlife.com"
FTP_PASSWORD = "@$CiaG3008"
FTP_FOLDER = "/diadoc_connector"   # папка на FTP
LOCAL_FOLDER = r"M:\Инвестиционный департамент\7.0 Treasury\Diadoc" # куда сохранить
# =====================

def simple_copy():
    print(f"Копирую файлы из {FTP_FOLDER} на FTP...")
    
    # Создаем папку
    os.makedirs(LOCAL_FOLDER, exist_ok=True)
    
    # Подключаемся
    ftp = FTP(FTP_SERVER)
    ftp.login(FTP_USER, FTP_PASSWORD)
    
    # Переходим в нужную папку
    ftp.cwd(FTP_FOLDER)
    
    # Получаем список файлов
    files = ftp.nlst()
    
    # Копируем каждый файл
    for file in files:
        if file not in [".", ".."]:
            print(f"Копирую: {file}")
            
            local_path = os.path.join(LOCAL_FOLDER, file)
            
            try:
                with open(local_path, 'wb') as f:
                    ftp.retrbinary(f'RETR {file}', f.write)
                print(f"  OK")
            except Exception as e:
                print(f"  Ошибка: {e}")
    
    ftp.quit()
    print(f"\nГотово! Файлы в папке: {LOCAL_FOLDER}")
    input("Нажмите Enter...")

# Запускаем
if __name__ == "__main__":
    simple_copy()



Ошибка: [Errno 22] Invalid argument: 'M:\\Инвестиционный департамент\\7.0 Treasury\\Diadoc\\????????? ?????????????'
Копирую: ????????? ???????
  Ошибка: [Errno 22] Invalid argument: 'M:\\Инвестиционный департамент\\7.0 Treasury\\Diadoc\\????????? ???????'
