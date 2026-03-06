import ftplib
import os

def download_all_files_bruteforce():
    """Скачивает все файлы из всех папок, используя сырые имена"""
    
    print("=" * 60)
    print("УНИВЕРСАЛЬНЫЙ СКРИПТ СКАЧИВАНИЯ")
    print("=" * 60)
    
    ftp_server = "ftp.renlife.com"
    ftp_user = "Ilya.Matveev2@mos.renlife.com"
    ftp_pass = "ыыыыыыы"
    ftp_folder = "/diadoc_connector"
    save_folder = r"M:\Инвестиционный департамент\7.0 Treasury\Diadoc"
    
    try:
        # Подключаемся
        ftp = ftplib.FTP(ftp_server)
        ftp.login(ftp_user, ftp_pass)
        ftp.encoding = 'latin-1'  # Самая безопасная кодировка
        
        # Переходим в нужную папку
        ftp.cwd(ftp_folder)
        print(f"\n✓ В папке: {ftp_folder}")
        
        # Получаем список всего
        items = ftp.nlst()
        print(f"✓ Найдено элементов: {len(items)}")
        
        # Создаем папку для сохранения
        os.makedirs(save_folder, exist_ok=True)
        
        total_files = 0
        processed_folders = 0
        
        print("\n" + "=" * 50)
        print("НАЧИНАЮ СКАЧИВАНИЕ")
        print("=" * 50)
        
        # Обрабатываем каждый элемент
        for item in items:
            if item in ['.', '..']:
                continue
            
            try:
                # Пробуем зайти как в папку
                ftp.cwd(item)
                print(f"\n📁 Найдена папка: {item}")
                
                # Получаем файлы в папке
                try:
                    folder_files = ftp.nlst()
                    print(f"  Содержит {len(folder_files)} элементов")
                    
                    # Скачиваем каждый файл из папки
                    for filename in folder_files:
                        if filename in ['.', '..']:
                            continue
                        
                        try:
                            # Проверяем размер (если это файл)
                            size = ftp.size(filename)
                            
                            # Формируем имя для сохранения
                            safe_name = f"[{item}]_{filename}"
                            # Заменяем недопустимые символы
                            safe_name = safe_name.replace('/', '_').replace('\\', '_')
                            local_path = os.path.join(save_folder, safe_name)
                            
                            print(f"    📄 Скачиваю: {filename} ({size/1024:.1f} KB)")
                            
                            # Скачиваем
                            with open(local_path, 'wb') as f:
                                ftp.retrbinary(f'RETR {filename}', f.write)
                            
                            print(f"      ✓ Сохранен как: {safe_name}")
                            total_files += 1
                            
                        except:
                            # Это папка внутри папки - игнорируем
                            pass
                    
                    processed_folders += 1
                    
                except Exception as e:
                    print(f"  Не удалось получить список файлов: {e}")
                
                # Возвращаемся назад
                ftp.cwd('..')
                
            except Exception as e:
                # Это файл, а не папка
                try:
                    size = ftp.size(item)
                    print(f"\n📄 Найден файл: {item} ({size/1024:.1f} KB)")
                    
                    # Сохраняем файл в корень
                    local_path = os.path.join(save_folder, item)
                    
                    with open(local_path, 'wb') as f:
                        ftp.retrbinary(f'RETR {item}', f.write)
                    
                    print(f"  ✓ Сохранен")
                    total_files += 1
                    
                except:
                    print(f"\n⚠️ Не удалось обработать: {item}")
        
        print("\n" + "=" * 50)
        print(f"✅ ГОТОВО!")
        print(f"   Обработано папок: {processed_folders}")
        print(f"   Скачано файлов: {total_files}")
        print(f"   Сохранено в: {save_folder}")
        
        # Показываем результат
        if os.path.exists(save_folder):
            files = os.listdir(save_folder)
            print(f"\nЛокальная папка содержит {len(files)} файлов:")
            for f in sorted(files)[:20]:
                print(f"  {f}")
        
        ftp.quit()
        
    except Exception as e:
        print(f"\n❌ Ошибка: {e}")
    
    input("\nНажмите Enter для выхода...")

# Запускаем
if __name__ == "__main__":
    download_all_files_bruteforce()
