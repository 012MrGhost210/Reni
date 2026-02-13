import win32com.client
import pythoncom
import os
from pathlib import Path

def send_from_template():
    # ========== ВВЕДИ СВОИ ДАННЫЕ ЗДЕСЬ ==========
    
    # Путь к файлу шаблона (там где твоя подпись)
    template_path = r"C:\Users\Username\Documents\template.msg"
    
    # Кому отправляем (через ; если несколько)
    to_emails = "ivan@mail.ru; petr@mail.ru"
    
    # Тема письма
    subject = "Письмо с подписью"
    
    # Текст, который нужно добавить в шаблон (в начало или конец)
    additional_text = """
Добрый день!
    
Это основной текст письма, который я хочу добавить в шаблон.
Можно писать что угодно, подпись останется из шаблона.
    
Спасибо!
"""
    
    # Куда добавить текст: "start" - в начало, "end" - в конец
    text_position = "end"  # или "start"
    
    # Папка с вложениями (все файлы оттуда прикрепятся)
    attachments_folder = r"C:\Users\Username\Documents\attachments"
    
    # ========== КОД ОТПРАВКИ (НИЧЕГО НЕ МЕНЯЙ) ==========
    
    pythoncom.CoInitialize()
    try:
        # Проверяем существование шаблона
        if not os.path.exists(template_path):
            print(f"✗ Шаблон не найден: {template_path}")
            return
        
        # Открываем Outlook и шаблон
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItemFromTemplate(template_path)
        
        # Устанавливаем получателей и тему
        mail.To = to_emails
        mail.Subject = subject
        
        # Добавляем текст в шаблон
        if additional_text:
            if text_position == "start":
                mail.Body = additional_text + "\n\n" + mail.Body
                print("✓ Текст добавлен в начало")
            else:
                mail.Body = mail.Body + "\n\n" + additional_text
                print("✓ Текст добавлен в конец")
        
        # Добавляем все файлы из папки как вложения
        if os.path.exists(attachments_folder):
            folder = Path(attachments_folder)
            files_added = 0
            
            for file_path in folder.glob("*"):
                if file_path.is_file():
                    try:
                        mail.Attachments.Add(str(file_path))
                        print(f"  ✓ Вложение: {file_path.name}")
                        files_added += 1
                    except Exception as e:
                        print(f"  ✗ Не удалось добавить {file_path.name}: {e}")
            
            if files_added == 0:
                print("  ! В папке нет файлов")
            else:
                print(f"✓ Всего добавлено вложений: {files_added}")
        else:
            print(f"! Папка с вложениями не найдена: {attachments_folder}")
            create = input("Создать папку? (да/нет): ").lower()
            if create in ['да', 'д', 'yes', 'y']:
                os.makedirs(attachments_folder)
                print(f"✓ Папка создана: {attachments_folder}")
        
        # Показываем что получилось перед отправкой
        print("\n" + "="*50)
        print("ПРОВЕРЬ ПИСЬМО:")
        print(f"Кому: {mail.To}")
        print(f"Тема: {mail.Subject}")
        print(f"Текст содержит подпись из шаблона + твой текст")
        print("="*50)
        
        # Спрашиваем подтверждение
        confirm = input("\nОтправить? (да/нет): ").lower()
        
        if confirm in ['да', 'д', 'yes', 'y']:
            mail.Send()
            print("✓ Письмо отправлено!")
        else:
            print("✗ Отправка отменена")
            # Сохраняем как черновик на всякий случай
            draft_name = f"draft_{Path(template_path).stem}.msg"
            draft_path = os.path.join(os.path.dirname(template_path), draft_name)
            mail.SaveAs(draft_path)
            print(f"✓ Черновик сохранен: {draft_path}")
        
    except Exception as e:
        print(f"✗ Ошибка: {e}")
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    send_from_template()
