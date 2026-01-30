import os
from ftplib import FTP

def main():
    # ====== ВАШИ ДАННЫЕ ======
    FTP_HOST = "ftp.renlife.com"
    FTP_USER = "Ilya.Matveev2@mos.renlife.com"
    FTP_PASS = "@$CiaG3008"
    FTP_FOLDER = "/diadoc_connector"
    LOCAL_FOLDER = r"M:\Инвестиционный департамент\7.0 Treasury\Diadoc"
    # =========================
    
    print("=" * 60)
    print("КОПИРОВАНИЕ ПАПОК С FTP")
    print("=" * 60)
    print(f"Папка на FTP: {FTP_FOLDER}")
    print(f"Сохранить в: {LOCAL_FOLDER}")
    print("-" * 60)
    
    # Создаем папку
    os.makedirs(LOCAL_FOLDER, exist_ok=True)
    print("✅ Локальная папка создана")
    
    # Подключаемся с UTF-8 кодировкой
    ftp = FTP(FTP_HOST)
    ftp.encoding = 'utf-8'  # Устанавливаем UTF-8!
    
    try:
        ftp.login(FTP_USER, FTP_PASS)
        print("✅ Подключение к FTP успешно")
    except Exception as e:
        print(f"❌ Ошибка подключения: {e}")
        input("Нажмите Enter...")
        return
    
    ftp.set_pasv(True)
    
    # Переходим в папку diadoc_connector
    try:
        ftp.cwd(FTP_FOLDER)
        print(f"✅ Перешел в папку: {FTP_FOLDER}")
    except Exception as e:
        print(f"❌ Не могу перейти в папку: {e}")
        ftp.quit()
        input("Нажмите Enter...")
        return
    
    # Получаем список всех элементов в папке
    print("\n📁 Получаю список папок...")
    
    try:
        # Получаем список всех элементов
        all_items = ftp.nlst()
        
        # Отфильтровываем только папки (проверяем какие элементы являются папками)
        folders = []
        
        for item in all_items:
            if item in [".", ".."]:
                continue
            
            print(f"Проверяю: {item}")
            
            # Пробуем войти в элемент - если получилось, это папка
            try:
                # Сохраняем текущую позицию
                current_dir = ftp.pwd()
                
                # Пробуем перейти в элемент
                ftp.cwd(item)
                # Если получилось - это папка!
                folders.append(item)
                
                # Возвращаемся назад
                ftp.cwd(current_dir)
                
                print(f"  ✓ Это папка: {item}")
                
            except:
                # Не получилось перейти - это не папка (или нет доступа)
                print(f"  ✗ Это не папка или нет доступа: {item}")
        
        print(f"\n✅ Найдено папок: {len(folders)}")
        
        if len(folders) == 0:
            print("⚠️  Папок не найдено!")
            ftp.quit()
            input("Нажмите Enter...")
            return
        
        # Копируем каждую папку
        print("\n" + "=" * 60)
        print("НАЧИНАЮ КОПИРОВАНИЕ ПАПОК...")
        print("=" * 60)
        
        copied_folders = 0
        
        for folder_name in folders:
            print(f"\n📂 Копирую папку: {folder_name}")
            
            # Создаем локальную папку
            local_folder_path = os.path.join(LOCAL_FOLDER, folder_name)
            os.makedirs(local_folder_path, exist_ok=True)
            
            # Рекурсивно копируем всю папку
            def copy_folder(ftp_path, local_path):
                """Рекурсивно копирует папку"""
                try:
                    # Переходим в папку на FTP
                    ftp.cwd(ftp_path)
                    
                    # Получаем все элементы в текущей папке
                    items_in_folder = ftp.nlst()
                    
                    for item in items_in_folder:
                        if item in [".", ".."]:
                            continue
                        
                        # Полный путь к элементу
                        item_ftp_path = f"{ftp_path}/{item}"
                        item_local_path = os.path.join(local_path, item)
                        
                        # Пробуем определить, папка это или файл
                        try:
                            # Пробуем войти в элемент
                            current = ftp.pwd()
                            ftp.cwd(item)
                            ftp.cwd(current)
                            
                            # Это папка - создаем и копируем рекурсивно
                            os.makedirs(item_local_path, exist_ok=True)
                            copy_folder(item_ftp_path, item_local_path)
                            
                        except:
                            # Это файл - скачиваем
                            try:
                                print(f"  📄 Скачиваю файл: {item}")
                                with open(item_local_path, 'wb') as f:
                                    ftp.retrbinary(f'RETR {item}', f.write)
                            except Exception as e:
                                print(f"  ⚠️  Ошибка файла {item}: {e}")
                    
                    # Возвращаемся на уровень выше
                    ftp.cwd("..")
                    
                except Exception as e:
                    print(f"  ❌ Ошибка в папке {ftp_path}: {e}")
            
            # Копируем текущую папку
            copy_folder(folder_name, local_folder_path)
            copied_folders += 1
            print(f"  ✅ Папка скопирована: {folder_name}")
        
        print(f"\n✅ Всего скопировано папок: {copied_folders}")
        print(f"📂 Сохранено в: {LOCAL_FOLDER}")
        
    except Exception as e:
        print(f"\n❌ Ошибка: {e}")
    
    finally:
        # Закрываем соединение
        ftp.quit()
        print("\n🔌 Соединение закрыто")
    
    print("\n" + "=" * 60)
    input("Нажмите Enter для выхода...")

if __name__ == "__main__":
    main()
