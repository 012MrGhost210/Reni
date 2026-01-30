#!/usr/bin/env python3
"""
СКРИПТ КОТОРЫЙ РАБОТАЕТ КАК КОМАНДА mget *.*
"""

import os
import sys

def simple_ftp_copy():
    print("Копирую все файлы с FTP...")
    
    # Ваши данные
    ftp_server = "ftp.renlife.com"
    ftp_user = "Ilya.Matveev2@mos.renlife.com"
    ftp_pass = "@$CiaG3008"
    ftp_folder = "/diadoc_connector"
    save_folder = r"M:\Инвестиционный департамент\7.0 Treasury\Diadoc"
    
    # Создаем папку
    os.makedirs(save_folder, exist_ok=True)
    
    # Создаем команды для FTP
    ftp_commands = f"""open {ftp_server}
{ftp_user}
{ftp_pass}
binary
prompt
cd {ftp_folder}
lcd "{save_folder}"
mget *.*
quit
"""
    
    # Сохраняем команды во временный файл
    with open("ftp_temp.txt", "w") as f:
        f.write(ftp_commands)
    
    # Запускаем FTP с командным файлом
    print("Запускаю FTP...")
    os.system('ftp -s:ftp_temp.txt')
    
    # Удаляем временный файл
    os.remove("ftp_temp.txt")
    
    print(f"\n✅ Все файлы скопированы в: {save_folder}")
    
    # Показываем что скопировалось
    print("\nСодержимое папки:")
    print("-" * 40)
    try:
        files = os.listdir(save_folder)
        for f in files:
            print(f"  {f}")
    except:
        pass
    
    input("\nНажмите Enter...")

# Запускаем
simple_ftp_copy()
