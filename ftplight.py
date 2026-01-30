import os
import sys
from ftplib import FTP

def decode_filename(encoded_name):
    """Пробует разные кодировки для декодирования имени файла"""
    # Если это уже строка, возвращаем как есть
    if isinstance(encoded_name, str):
        return encoded_name
    
    # Если это bytes, пробуем разные кодировки
    if isinstance(encoded_name, bytes):
        # Список кодировок для проверки (в порядке вероятности)
        encodings = ['cp1251', 'cp866', 'iso-8859-5', 'koi8-r', 'utf-8', 'windows-1251']
        
        for encoding in encodings:
            try:
                decoded = encoded_name.decode(encoding)
                # Проверяем, получилась ли нормальная строка
                if any(c.isalpha() for c in decoded):
                    return decoded
            except:
                continue
    
    # Если ничего не помогло, возвращаем как строку
    return str(encoded_name)

def main():
    # ====== ВАШИ ДАННЫЕ ======
    FTP_HOST = "ftp.renlife.com"
    FTP_USER = "Ilya.Matveev2@mos.renlife.com"
    FTP_PASS = "@$CiaG3008"
    FTP_FOLDER = "/diadoc_connector"
    LOCAL_FOLDER = r"M:\Инвестиционный департамент\7.0 Treasury\Diadoc"
    # =========================
    
    print("=" * 70)
    print("СКРИПТ ДЛЯ КОПИРОВАНИЯ ПАПОК С RENLIFE FTP")
    print("=" * 70)
    print(f"Сервер: {FTP_HOST}")
    print(f"Папка на FTP: {FTP_FOLDER}")
    print(f"Сохранить в: {LOCAL_FOLDER}")
    print("-" * 70)
    
    # Создаем папку для сохранения
    os.makedirs(LOCAL_FOLDER, exist_ok=True)
    print("✅ Локальная папка создана")
    
    # Пробуем подключиться с разными кодировками
    ftp = None
    
    for encoding in [None, 'cp1251', 'utf-8', 'cp866']:
        try:
            print(f"\n🔌 Пробую подключиться с кодировкой: {encoding or 'по умолчанию'}")
            ftp = FTP(FTP_HOST, timeout=30)
            
            if encoding:
                ftp.encoding = encoding
            
            ftp.login(FTP_USER, FTP_PASS)
            ftp.set_pasv(True)
            print(f"✅ Успешно! Кодировка: {encoding or 'default'}")
            break
            
        except Exception as e:
            print(f"❌ Не удалось: {e}")
            if ftp:
                try:
                    ftp.quit()
                except:
                    pass
            ftp = None
    
    if ftp is None:
        print("\n❌ Не удалось подключиться ни с одной кодировкой")
        input("Нажмите Enter...")
        return
    
    # Переходим в нужную папку
    try:
        ftp.cwd(FTP_FOLDER)
        print(f"✅ Перешел в папку: {FTP_FOLDER}")
    except Exception as e:
        print(f"❌ Не могу перейти в папку: {e}")
        ftp.quit()
        input("Нажмите Enter...")
        return
    
    # Получаем список RAW элементов (без декодирования)
    print("\n📋 Получаю список элементов...")
    
    try:
        # Используем raw команду для получения списка
        ftp.voidcmd('TYPE A')  # Переключаемся в ASCII режим
        
        lines = []
        ftp.retrlines('LIST', lines.append)
        
        print(f"✅ Получено строк: {len(lines)}")
        
        # Разбираем список
        folders = []
        
        for line in lines:
            parts = line.split()
            if len(parts) < 9:
                continue
            
            # Тип элемента (первый символ)
            item_type = parts[0][0]
            
            # Имя элемента (все что после 8, воссоединяем)
            encoded_name = ' '.join(parts[8:])
            
            # Декодируем имя
            decoded_name = decode_filename(encoded_name)
            
            if item_type == 'd':  # 'd' означает directory (папка)
                print(f"📁 Найдена папка: {decoded_name}")
                folders.append(decoded_name)
            else:
                print(f"📄 Найден файл: {decoded_name} (пропускаем)")
        
        print(f"\n✅ Всего найдено папок: {len(folders)}")
        
        if len(folders) == 0:
            print("⚠️  Папок не найдено!")
            
            # Пробуем альтернативный способ
            print("\n🔄 Пробую альтернативный способ...")
            try:
                items = ftp.nlst()
                print(f"Найдено элементов через NLST: {len(items)}")
                
                for item in items:
                    if item not in [".", ".."]:
                        decoded_item = decode_filename(item)
                        print(f"  • {decoded_item}")
                        
                        # Пробуем проверить, папка ли это
                        try:
                            original_dir = ftp.pwd()
                            ftp.cwd(item)
                            ftp.cwd(original_dir)
                            folders.append(decoded_item)
                            print(f"    ✓ Это папка")
                        except:
                            print(f"    ✗ Это не папка или нет доступа")
                
            except Exception as e:
                print(f"Ошибка альтернативного способа: {e}")
        
        # Если все равно нет папок
        if len(folders) == 0:
            print("\n❌ Папок для копирования не найдено!")
            ftp.quit()
            input("Нажмите Enter...")
            return
        
        # Копируем папки
        print("\n" + "=" * 70)
        print("НАЧИНАЮ КОПИРОВАНИЕ ПАПОК...")
        print("=" * 70)
        
        for folder_name in folders:
            print(f"\n📂 Копирую папку: {folder_name}")
            
            # Создаем локальную папку
            local_folder_path = os.path.join(LOCAL_FOLDER, folder_name)
            os.makedirs(local_folder_path, exist_ok=True)
            
            # Рекурсивная функция для копирования
            def copy_folder_recursive(ftp_rel_path, local_full_path):
                """Рекурсивно копирует папку с FTP"""
                try:
                    # Переходим в папку на FTP
                    ftp.cwd(ftp_rel_path)
                    
                    # Получаем список элементов в папке
                    items_in_folder = []
                    ftp.retrlines('LIST', items_in_folder.append)
                    
                    for line in items_in_folder:
                        parts = line.split()
                        if len(parts) < 9:
                            continue
                        
                        item_type = parts[0][0]
                        encoded_item_name = ' '.join(parts[8:])
                        decoded_item_name = decode_filename(encoded_item_name)
                        
                        if decoded_item_name in [".", ".."]:
                            continue
                        
                        item_local_path = os.path.join(local_full_path, decoded_item_name)
                        
                        if item_type == 'd':
                            # Это подпапка
                            os.makedirs(item_local_path, exist_ok=True)
                            print(f"  📁 Подпапка: {decoded_item_name}/")
                            copy_folder_recursive(
                                f"{ftp_rel_path}/{decoded_item_name}",
                                item_local_path
                            )
                        else:
                            # Это файл
                            print(f"  📄 Файл: {decoded_item_name}")
                            try:
                                with open(item_local_path, 'wb') as f:
                                    ftp.retrbinary(f'RETR {decoded_item_name}', f.write)
                            except Exception as e:
                                print(f"    ⚠️  Ошибка скачивания: {e}")
                    
                    # Возвращаемся на уровень выше
                    ftp.cwd("..")
                    
                except Exception as e:
                    print(f"  ❌ Ошибка в папке {ftp_rel_path}: {e}")
            
            # Копируем текущую папку
            copy_folder_recursive(folder_name, local_folder_path)
            print(f"  ✅ Папка скопирована: {folder_name}")
        
        print(f"\n✅ Все папки скопированы!")
        print(f"📂 Сохранено в: {LOCAL_FOLDER}")
        
    except Exception as e:
        print(f"\n❌ Ошибка при обработке: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        # Закрываем соединение
        try:
            ftp.quit()
            print("\n🔌 Соединение закрыто")
        except:
            pass
    
    print("\n" + "=" * 70)
    input("Нажмите Enter для выхода...")

if __name__ == "__main__":
    main()
