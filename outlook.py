import win32com.client
import pythoncom
import os

def send_email():
    # ВВЕДИ СВОИ ДАННЫЕ ЗДЕСЬ:
    to_emails = "ivan@mail.ru; petr@mail.ru"  # Кому (через ;)
    subject = "Тестовое письмо"                # Тема
    body = """Привет!
    
Это текст письма.
Можно писать в несколько строк."""             # Текст письма
    
    # УКАЖИ ПУТЬ К ФАЙЛУ ВЛОЖЕНИЯ (или оставь пустым "", если без вложения)
    attachment_path = r"C:\Users\Username\Desktop\file.pdf"
    
    # ========== КОД ОТПРАВКИ (НИЧЕГО НЕ МЕНЯЙ) ==========
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        
        mail.To = to_emails
        mail.Subject = subject
        mail.Body = body
        
        # Добавляем вложение, если указан путь
        if attachment_path:
            if os.path.exists(attachment_path):
                mail.Attachments.Add(attachment_path)
                print(f"✓ Вложение добавлено: {os.path.basename(attachment_path)}")
            else:
                print(f"✗ Файл не найден: {attachment_path}")
                print("Письмо будет отправлено без вложения!")
        
        mail.Send()
        print("✓ Письмо успешно отправлено!")
        print(f"  Кому: {to_emails}")
        print(f"  Тема: {subject}")
        
    except Exception as e:
        print(f"✗ Ошибка: {e}")
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    send_email()
