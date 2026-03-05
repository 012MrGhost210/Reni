import win32com.client
import pythoncom
import os

# Конфигурация
DRAFTS_FOLDER = r"D:\мои_письма"  # Папка с .msg файлами
STOP_FILE = r"D:\мои_письма\stop.txt"  # Файл-флаг

def send_msg_files():
    """Отправляет .msg файлы из указанной папки"""
    
    # Проверяем stop.txt
    if os.path.exists(STOP_FILE):
        print(f"Найден {STOP_FILE}, отправка отменена")
        return
    
    print(f"Поиск .msg файлов в {DRAFTS_FOLDER}...")
    
    try:
        # Подключаемся к Outlook
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        
        # Находим все .msg файлы
        msg_files = [f for f in os.listdir(DRAFTS_FOLDER) if f.lower().endswith('.msg')]
        
        if not msg_files:
            print("Нет .msg файлов для отправки")
            # Создаем stop.txt даже если файлов нет
            with open(STOP_FILE, 'w') as f:
                f.write("No files found")
            return
        
        print(f"Найдено файлов: {len(msg_files)}")
        
        sent_count = 0
        
        # Отправляем каждый файл
        for filename in msg_files:
            filepath = os.path.join(DRAFTS_FOLDER, filename)
            try:
                # Открываем .msg файл в Outlook
                msg = outlook.CreateItemFromTemplate(filepath)
                
                # Отправляем
                msg.Send()
                sent_count += 1
                print(f"✓ Отправлено: {filename}")
                
            except Exception as e:
                print(f"✗ Ошибка с {filename}: {e}")
        
        print(f"\nВсего отправлено: {sent_count}")
        
        # Создаем stop.txt
        with open(STOP_FILE, 'w') as f:
            f.write(f"OK\n{sent_count}")
        
        print(f"Создан {STOP_FILE}")
        
    except Exception as e:
        print(f"Ошибка: {e}")
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    send_msg_files()
