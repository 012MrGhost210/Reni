import ftplib
import os
import chardet

def check_ftp_encoding_and_folders():
    print("=" * 60)
    print("ДИАГНОСТИКА FTP СЕРВЕРА")
    print("=" * 60)
    
    ftp_server = "ftp.renlife.com"
    ftp_user = "Ilya.Matveev2@mos.renlife.com"
    ftp_pass = "ыыыыыыы"
    ftp_folder = "/diadoc_connector"
    
    try:
        # Подключаемся
        ftp = ftplib.FTP(ftp_server)
        ftp.login(ftp_user, ftp_pass)
        
        print(f"\nПодключено к {ftp_server}")
        
        # Переходим в папку
        ftp.cwd(ftp_folder)
        print(f"В папке: {ftp_folder}")
        
        # Пробуем разные кодировки
        encodings = ['cp1251', 'utf-8', 'koi8-r', 'cp866', 'latin-1']
        
        print("\nПробуем разные кодировки для получения списка папок:")
        print("-" * 50)
        
        for encoding in encodings:
            try:
                ftp.encoding = encoding
                print(f"\nКодировка: {encoding}")
                
                # Получаем список
                items = ftp.nlst()
                print(f"Получено элементов: {len(items)}")
                
                # Показываем русские названия (если есть)
                for item in items[:10]:  # Первые 10
                    # Проверяем, похоже ли на русский
                    if any(ord(c) > 127 for c in item):
                        print(f"  {item}")
                
                # Если нашли русские названия
                if any(any(ord(c) > 127 for c in item) for item in items):
                    print(f"✓ В кодировке {encoding} видны русские буквы!")
                    
                    # Показываем все папки
                    print("\nПапки на сервере:")
                    for item in items:
                        if item not in ['.', '..']:
                            try:
                                ftp.cwd(item)
                                print(f"  📁 {item} (папка)")
                                ftp.cwd('..')
                            except:
                                print(f"  📄 {item} (файл)")
                    
                    # Запоминаем рабочую кодировку
                    working_encoding = encoding
                    break
                    
            except Exception as e:
                print(f"Ошибка с {encoding}: {e}")
        
        if 'working_encoding' in locals():
            print(f"\n✅ РАБОЧАЯ КОДИРОВКА: {working_encoding}")
            
            # Проверяем наличие наших папок
            print("\nПроверяем наличие целевых папок:")
            target_folders = [
                "Аннулирован",
                "Документооборот завершён",
                "Подписан",
                "Требуется аннулирование", 
                "Требуется подпись"
            ]
            
            items = ftp.nlst()
            items_lower = [item.lower() for item in items]
            
            for target in target_folders:
                found = False
                for item in items:
                    # Сравниваем без учета регистра
                    if target.lower() in item.lower() or item.lower() in target.lower():
                        print(f"  {target} -> найдено как: {item}")
                        found = True
                        break
                if not found:
                    print(f"  {target} -> НЕ НАЙДЕНО!")
        else:
            print("\n❌ Не удалось найти рабочую кодировку")
        
        ftp.quit()
        
    except Exception as e:
        print(f"Ошибка: {e}")
    
    input("\nНажмите Enter для выхода...")
    Подключено к ftp.renlife.com
В папке: /diadoc_connector

Пробуем разные кодировки для получения списка папок:
--------------------------------------------------

Кодировка: cp1251
Получено элементов: 5

Кодировка: utf-8
Получено элементов: 5

Кодировка: koi8-r
Получено элементов: 5

Кодировка: cp866
Получено элементов: 5

Кодировка: latin-1
Получено элементов: 5

❌ Не удалось найти рабочую кодировку

# Запускаем диагностику
check_ftp_encoding_and_folders()
