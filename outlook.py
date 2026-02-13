import win32com.client
import pythoncom
import os

def send_draft():
    # ========== ВВЕДИ СВОИ ДАННЫЕ ЗДЕСЬ ==========
    
    # Путь к файлу драфта (msg файл)
    draft_path = r"C:\Users\Username\Documents\draft.msg"
    
    # Текст, который нужно вставить (вместо старого)
    new_text = """Добрый день! Это новый текст письма.
    
Все остальное (получатели, тема, подпись) остается из драфта."""
    
    # Путь к файлу вложения (или оставь пустым "", если без вложения)
    attachment_path = r"C:\Users\Username\Documents\file.pdf"
    
    # ========== КОД ОТПРАВКИ ==========
    
    pythoncom.CoInitialize()
    try:
        # Открываем драфт
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItemFromTemplate(draft_path)
        
        # Заменяем текст
        mail.Body = new_text
        
        # Добавляем вложение если указано
        if attachment_path:
            if os.path.exists(attachment_path):
                mail.Attachments.Add(attachment_path)
                print(f"✓ Вложение добавлено: {os.path.basename(attachment_path)}")
            else:
                print(f"✗ Файл не найден: {attachment_path}")
        
        # Отправляем
        mail.Send()
        print("✓ Письмо отправлено!")
        
    except Exception as e:
        print(f"✗ Ошибка: {e}")
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    send_draft()
