import ftplib
import os

def download_with_structure():
    """Сохраняет файлы с сохранением структуры папок"""
    
    ftp_server = "ftp.renlife.com"
    ftp_user = "Ilya.Matveev2@mos.renlife.com"
    ftp_pass = "@$CiaG3008"
    ftp_folder = "/diadoc_connector"
    save_folder = r"M:\Инвестиционный департамент\7.0 Treasury\Diadoc"
    
    target_folders = [
        "Аннулирован",
        "Документооборот завершён", 
        "Подписан",
        "Требуется аннулирование",
        "Требуется подпись"
    ]
    
    ftp = ftplib.FTP(ftp_server)
    ftp.login(ftp_user, ftp_pass)
    ftp.encoding = 'cp1251'
    ftp.cwd(ftp_folder)
    
    for folder in target_folders:
        try:
            print(f"\nОбрабатываю: {folder}")
            ftp.cwd(folder)
            
            # Создаем локальную папку с таким же именем
            local_folder = os.path.join(save_folder, folder)
            os.makedirs(local_folder, exist_ok=True)
            
            # Получаем файлы
            files = ftp.nlst()
            
            for file in files:
                if file not in ['.', '..']:
                    print(f"  Скачиваю: {file}")
                    local_path = os.path.join(local_folder, file)
                    
                    with open(local_path, 'wb') as f:
                        ftp.retrbinary(f'RETR {file}', f.write)
            
            ftp.cwd('..')
            
        except Exception as e:
            print(f"  Ошибка с папкой {folder}: {e}")
    
    ftp.quit()
    print(f"\nГотово! Файлы сохранены в {save_folder}")

# Запуск
download_with_structure()
input("Нажмите Enter...")

Обрабатываю: Аннулирован
  Ошибка с папкой Аннулирован: 550 Folder Àííóëèðîâàí not found

Обрабатываю: Документооборот завершён
  Ошибка с папкой Документооборот завершён: 550 Folder Äîêóìåíòîîáîðîò çàâåðø¸í not found

Обрабатываю: Подписан
  Ошибка с папкой Подписан: 550 Folder Ïîäïèñàí not found

Обрабатываю: Требуется аннулирование
  Ошибка с папкой Требуется аннулирование: 550 Folder Òðåáóåòñÿ àííóëèðîâàíèå not found

Обрабатываю: Требуется подпись
  Ошибка с папкой Требуется подпись: 550 Folder Òðåáóåòñÿ ïîäïèñü not found
