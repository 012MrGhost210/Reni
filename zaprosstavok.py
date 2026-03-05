import win32com.client
import pythoncom
import os

# Конфигурация
STOP_FILE = r"D:\мои_письма\stop.txt"  # Путь к файлу-флагу
TARGET_FOLDER = "Мои готовые письма"    # Имя папки в Outlook

def send_from_specific_folder():
    """Отправляет письма из конкретной папки Outlook"""
    
    # Проверяем stop файл
    if os.path.exists(STOP_FILE):
        print(f"Найден {STOP_FILE}, отправка пропущена")
        return
    
    print(f"Поиск папки '{TARGET_FOLDER}'...")
    
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Ищем нужную папку
        target_folder = None
        
        # Поиск во всех почтовых ящиках
        for store in namespace.Folders:
            for folder in store.Folders:
                if folder.Name == TARGET_FOLDER:
                    target_folder = folder
                    break
            if target_folder:
                break
        
        if not target_folder:
            print(f"Папка '{TARGET_FOLDER}' не найдена")
            return
        
        print(f"Папка найдена, начинаем отправку...")
        
        sent_count = 0
        
        # Отправляем все письма из папки
        for item in target_folder.Items:
            if item.Class == 43:  # Письмо
                try:
                    item.Send()
                    sent_count += 1
                    print(f"✓ {sent_count}. {item.Subject}")
                except Exception as e:
                    print(f"✗ Ошибка: {e}")
        
        print(f"\nОтправлено писем: {sent_count}")
        
        # Создаем stop.txt
        with open(STOP_FILE, 'w') as f:
            f.write(f"OK\n{sent_count}")
        
        print(f"Создан {STOP_FILE}")
        
    except Exception as e:
        print(f"Ошибка: {e}")
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    send_from_specific_folder()
