import ftplib
import os

def deep_ftp_diagnostics():
    print("=" * 60)
    print("ГЛУБОКАЯ ДИАГНОСТИКА FTP СЕРВЕРА")
    print("=" * 60)
    
    ftp_server = "ftp.renlife.com"
    ftp_user = "Ilya.Matveev2@mos.renlife.com"
    ftp_pass = "ыыыыыыы"
    ftp_folder = "/diadoc_connector"
    
    try:
        # Подключаемся
        ftp = ftplib.FTP(ftp_server)
        ftp.login(ftp_user, ftp_pass)
        
        print(f"\n✓ Подключено к {ftp_server}")
        print(f"✓ Пользователь: {ftp_user}")
        
        # Переходим в папку
        ftp.cwd(ftp_folder)
        print(f"✓ В папке: {ftp_folder}")
        
        # Получаем детальную информацию разными методами
        print("\n" + "=" * 40)
        print("МЕТОД 1: LIST (подробный список)")
        print("=" * 40)
        
        # Используем LIST для получения детальной информации
        files_list = []
        ftp.dir(files_list.append)
        
        print(f"Найдено элементов: {len(files_list)}")
        for i, item in enumerate(files_list, 1):
            print(f"{i:2}. {item}")
            
            # Пробуем разобрать строку
            parts = item.split()
            if len(parts) >= 9:
                filename = ' '.join(parts[8:])
                print(f"   Имя файла: {filename}")
                print(f"   Права: {parts[0]}")
                print(f"   Размер: {parts[4]} байт")
        
        print("\n" + "=" * 40)
        print("МЕТОД 2: NLST (простой список)")
        print("=" * 40)
        
        # Используем NLST для получения простого списка
        try:
            nlst_items = ftp.nlst()
            print(f"Найдено элементов: {len(nlst_items)}")
            for i, item in enumerate(nlst_items, 1):
                print(f"{i:2}. {item}")
                
                # Пытаемся определить тип (файл или папка)
                try:
                    ftp.cwd(item)
                    print(f"   → ЭТО ПАПКА")
                    ftp.cwd('..')
                except:
                    try:
                        size = ftp.size(item)
                        print(f"   → ЭТО ФАЙЛ, размер: {size} байт")
                    except:
                        print(f"   → НЕИЗВЕСТНЫЙ ТИП")
        except Exception as e:
            print(f"Ошибка NLST: {e}")
        
        print("\n" + "=" * 40)
        print("МЕТОД 3: MLSD (стандартный листинг)")
        print("=" * 40)
        
        # Пробуем MLSD (более современный метод)
        try:
            mlsd_items = list(ftp.mlsd())
            print(f"Найдено элементов: {len(mlsd_items)}")
            for name, facts in mlsd_items:
                print(f"  Имя: {name}")
                print(f"  Тип: {facts.get('type', 'unknown')}")
                print(f"  Размер: {facts.get('size', 'unknown')}")
                print(f"  Модификация: {facts.get('modify', 'unknown')}")
                print("  ---")
        except Exception as e:
            print(f"MLSD не поддерживается: {e}")
        
        print("\n" + "=" * 40)
        print("МЕТОД 4: ПРОВЕРКА РАЗНЫХ КОДИРОВОК ДЛЯ LIST")
        print("=" * 40)
        
        encodings = ['cp1251', 'utf-8', 'koi8-r', 'cp866', 'latin-1', 'cp1252', 'mac-cyrillic']
        
        for encoding in encodings:
            try:
                ftp.encoding = encoding
                print(f"\nКодировка: {encoding}")
                
                # Получаем LIST с этой кодировкой
                list_data = []
                ftp.dir(list_data.append)
                
                for item in list_data:
                    # Показываем байтовое представление для анализа
                    print(f"  {item}")
                    
                    # Показываем hex значения для первых символов
                    if item:
                        bytes_item = item.encode('latin-1', errors='ignore')
                        hex_str = ' '.join([f'{b:02x}' for b in bytes_item[:20]])
                        print(f"  HEX: {hex_str}")
                        
            except Exception as e:
                print(f"  Ошибка: {e}")
        
        # Проверяем содержимое папок
        print("\n" + "=" * 40)
        print("ПРОВЕРКА СОДЕРЖИМОГО ПАПОК")
        print("=" * 40)
        
        # Получаем список элементов
        items = ftp.nlst()
        
        for item in items:
            if item not in ['.', '..']:
                try:
                    # Пробуем зайти в папку
                    ftp.cwd(item)
                    print(f"\n📁 Папка: {item}")
                    
                    # Смотрим что внутри
                    subitems = ftp.nlst()
                    print(f"  Содержит {len(subitems)} элементов")
                    
                    # Показываем первые 5
                    for subitem in subitems[:5]:
                        print(f"    - {subitem}")
                    
                    ftp.cwd('..')
                except:
                    print(f"\n📄 Файл: {item}")
        
        ftp.quit()
        
    except Exception as e:
        print(f"\n❌ Ошибка: {e}")
    
    input("\nНажмите Enter для выхода...")

# Запускаем глубокую диагностику
deep_ftp_diagnostics()
