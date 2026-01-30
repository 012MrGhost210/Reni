#!/usr/bin/env python3
"""
Простой скрипт для копирования ВСЕГО с FTP
"""

from ftplib import FTP
import os
import sys

def copy_all_ftp(server, user, password, remote_path="/", local_path="./ftp_copy"):
    """Копирует абсолютно все с FTP сервера"""
    
    ftp = FTP(server)
    ftp.login(user, password)
    
    def copy_dir(remote, local):
        ftp.cwd(remote)
        os.makedirs(local, exist_ok=True)
        
        for item in ftp.nlst():
            if item in [".", ".."]:
                continue
                
            new_remote = f"{remote}/{item}" if remote != "/" else f"/{item}"
            new_local = os.path.join(local, item)
            
            try:
                ftp.cwd(item)
                ftp.cwd("..")
                # Это папка
                print(f"Папка: {item}")
                copy_dir(new_remote, new_local)
            except:
                # Это файл
                print(f"Файл: {item}")
                with open(new_local, 'wb') as f:
                    ftp.retrbinary(f'RETR {item}', f.write)
    
    print("Начинаю копирование...")
    copy_dir(remote_path, local_path)
    ftp.quit()
    print("Готово!")

# ПРОСТО ВСТАВЬТЕ СВОИ ДАННЫЕ И ЗАПУСТИТЕ!
copy_all_ftp(
    server="ftp.renlife.com",
    user="Ilya.Matveev2@mos.renlife.com",
    password="@$CiaG3008",
    remote_path="/diadoc_connector",  # откуда копировать (обычно корень)
    local_path="M:\Инвестиционный департамент\7.0 Treasury\Diadoc"  # куда сохранить
)
