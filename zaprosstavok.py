import win32com.client
import pythoncom
import os
from datetime import datetime

# Конфигурация
FLAG_FILE = r"C:\temp\outlook_sent.flag"

def send_drafts():
    """Отправляет черновики если нет флага"""
    
    # Если есть флаг - выходим
    if os.path.exists(FLAG_FILE):
        print(f"[{datetime.now()}] Флаг найден, пропускаем отправку")
        return
    
    print(f"[{datetime.now()}] Начинаем отправку...")
    
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        drafts = namespace.GetDefaultFolder(16)
        
        sent = 0
        errors = 0
        
        for item in drafts.Items:
            if item.Class == 43:  # Письмо
                try:
                    item.Send()
                    sent += 1
                    print(f"  Отправлено: {item.Subject}")
                except Exception as e:
                    errors += 1
                    print(f"  Ошибка: {e}")
        
        print(f"Итого: отправлено {sent}, ошибок {errors}")
        
        # Создаем флаг
        with open(FLAG_FILE, 'w') as f:
            f.write(f"Sent: {sent}, Errors: {errors}\n")
            f.write(f"Date: {datetime.now()}")
        
        print(f"Флаг создан: {FLAG_FILE}")
        
    except Exception as e:
        print(f"Критическая ошибка: {e}")
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    send_drafts()
