import ftplib
import os

def simple_copy():
    """Максимально простое копирование всех файлов"""
    
    # Настройки
    server = "ftp.renlife.com"
    user = "Ilya.Matveev2@mos.renlife.com"
    password = "ыыыыыыы"
    remote_dir = "/diadoc_connector"
    local_dir = r"M:\Инвестиционный департамент\7.0 Treasury\Diadoc"
    
    print(f"Копирую файлы из {remote_dir}...")
    
    # Подключение
    ftp = ftplib.FTP(server)
    ftp.login(user, password)
    ftp.encoding = 'cp1251'  # Для русских букв
    ftp.cwd(remote_dir)
    
    # Создаем локальную папку
    os.makedirs(local_dir, exist_ok=True)
    
    # Получаем все файлы (простой список)
    files = ftp.nlst()
    
    print(f"Найдено элементов: {len(files)}")
    
    # Копируем все подряд
    count = 0
    for item in files:
        if item in ['.', '..']:
            continue
            
        print(f"Копирую: {item}")
        local_path = os.path.join(local_dir, item)
        
        try:
            with open(local_path, 'wb') as f:
                ftp.retrbinary(f'RETR {item}', f.write)
            print(f"  ✓ OK")
            count += 1
        except:
            print(f"  → Пропускаю (возможно папка)")
    
    ftp.quit()
    print(f"\nСкопировано файлов: {count}")
    print(f"В папку: {local_dir}")

if __name__ == "__main__":
    simple_copy()
    input("\nНажмите Enter...")


Аннулирован
Документооборот завершён
Подписан
Требуется аннулирование
Требуется подпись
