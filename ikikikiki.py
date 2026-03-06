import ftplib
import os

def simple_ftp_download_all():
    """Простое скачивание всех файлов (без вложенных папок)"""
    
    ftp_server = "ftp.renlife.com"
    ftp_user = "Ilya.Matveev2@mos.renlife.com" 
    ftp_pass = "ыыыыыыы"
    ftp_folder = "/diadoc_connector"
    save_folder = r"M:\Инвестиционный департамент\7.0 Treasury\Diadoc"
    
    # Подключаемся
    ftp = ftplib.FTP(ftp_server)
    ftp.login(ftp_user, ftp_pass)
    ftp.encoding = 'utf-8'
    ftp.cwd(ftp_folder)
    
    # Создаем папку
    os.makedirs(save_folder, exist_ok=True)
    
    # Получаем список файлов
    files = ftp.nlst()  # Получаем простой список имен
    
    print(f"Найдено файлов: {len(files)}")
    
    # Скачиваем каждый файл
    for i, filename in enumerate(files, 1):
        if filename not in ['.', '..']:
            print(f"[{i}/{len(files)}] Скачиваю: {filename}")
            local_path = os.path.join(save_folder, filename)
            
            try:
                with open(local_path, 'wb') as f:
                    ftp.retrbinary(f'RETR {filename}', f.write)
                print(f"  ✓ OK")
            except Exception as e:
                print(f"  ✗ Ошибка: {e}")
    
    ftp.quit()
    print(f"\nГотово! Файлы в: {save_folder}")

simple_ftp_download_all()
