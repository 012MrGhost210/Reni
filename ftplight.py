import os
from ftplib import FTP

# ДАННЫЕ
ftp = FTP("ftp.renlife.com")
ftp.encoding = 'utf-8'  # ЭТО ВАЖНО!
ftp.login("Ilya.Matveev2@mos.renlife.com", "@$CiaG3008")
ftp.set_pasv(True)

# ПЕРЕХОДИМ В НУЖНУЮ ПАПКУ
ftp.cwd("/diadoc_connector")

# КУДА СОХРАНЯЕМ
local = r"M:\Инвестиционный департамент\7.0 Treasury\Diadoc"
os.makedirs(local, exist_ok=True)

print("Копирую папки из /diadoc_connector...")

# ФУНКЦИЯ КОПИРОВАНИЯ
def копировать_папку(ftp_путь, локальный_путь):
    # ПЕРЕХОДИМ В ПАПКУ НА FTP
    ftp.cwd(ftp_путь)
    
    # СОЗДАЕМ ЛОКАЛЬНУЮ ПАПКУ
    os.makedirs(локальный_путь, exist_ok=True)
    
    # ВСЕ ЧТО ЕСТЬ В ЭТОЙ ПАПКЕ
    for элемент in ftp.nlst():
        if элемент in [".", ".."]:
            continue
        
        локальный_файл = os.path.join(локальный_путь, элемент)
        
        # ПРОБУЕМ - ЕСЛИ МОЖНО ВОЙТИ, ТО ЭТО ПАПКА
        try:
            ftp.cwd(элемент)
            ftp.cwd("..")
            # ЭТО ПАПКА - КОПИРУЕМ РЕКУРСИВНО
            копировать_папку(f"{ftp_путь}/{элемент}", локальный_файл)
        except:
            # ЭТО ФАЙЛ - СКАЧИВАЕМ
            try:
                with open(локальный_файл, 'wb') as f:
                    ftp.retrbinary(f'RETR {элемент}', f.write)
                print(f"  Файл: {элемент}")
            except:
                pass
    
    # ВОЗВРАЩАЕМСЯ НАЗАД
    ftp.cwd("..")

# КОПИРУЕМ ВСЕ ПАПКИ В diadoc_connector
for папка in ftp.nlst():
    if папка not in [".", ".."]:
        print(f"\nПапка: {папка}")
        копировать_папку(папка, os.path.join(local, папка))

ftp.quit()
print(f"\n✅ Готово! Все папки в: {local}")
input("Нажмите Enter...")


Проверяю: ??????????????? ????????
  ✗ Это не папка или нет доступа: ??????????????? ????????
Проверяю: ????????
  ✗ Это не папка или нет доступа: ????????
Проверяю: ????????? ?????????????
  ✗ Это не папка или нет доступа: ????????? ?????????????
Проверяю: ????????? ???????
  ✗ Это не папка или нет доступа: ????????? ???????
