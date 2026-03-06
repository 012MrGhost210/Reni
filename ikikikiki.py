from ftplib import FTP
import os

# Параметры соединения
host = "ftp.renlife.com"
username = "Ilya.Matveev2@mos.renlife.com"
password = "ыыыыыы"

try:
    # Подключение к FTP серверу
    ftp = FTP(host)
    ftp.login(user=username, passwd=password)
    ftp.set_pasv(True)  # Пассивный режим
    
    print("Подключение установлено")
    
    # Переход в директорию
    ftp.cwd('/diadoc_connector')
    
    # Получение списка файлов
    files = ftp.nlst()
    
    # Локальная директория для сохранения
    local_dir = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury"
    
    # Загрузка каждого файла
    for file in files:
        try:
            local_path = os.path.join(local_dir, file)
            with open(local_path, 'wb') as local_file:
                ftp.retrbinary(f'RETR {file}', local_file.write)
            print(f"Файл {file} успешно загружен")
        except Exception as e:
            print(f"Ошибка при загрузке файла {file}: {e}")
    
    # Закрытие соединения
    ftp.quit()

except Exception as e:
    print(f"Ошибка: {e}")
