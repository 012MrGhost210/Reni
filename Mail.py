import win32com.client
import os
import re
from datetime import datetime
import pythoncom
import sys

class EmailDraftSender:
    def __init__(self):
        """Инициализация объекта Outlook"""
        try:
            pythoncom.CoInitialize()
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
        except Exception as e:
            print(f"Ошибка подключения к Outlook: {e}")
            sys.exit(1)
    
    def load_draft_template(self, draft_path):
        """
        Загрузка шаблона письма из файла .msg или .oft
        
        Args:
            draft_path: путь к файлу шаблона (.msg или .oft)
        
        Returns:
            объект MailItem
        """
        try:
            if not os.path.exists(draft_path):
                raise FileNotFoundError(f"Файл шаблона не найден: {draft_path}")
            
            # Создаем объект для работы с файловой системой
            file_system = self.namespace.GetFolderFromID(
                self.namespace.Folders.Item(1).EntryID
            )
            
            # Загружаем шаблон
            mail_item = self.outlook.CreateItemFromTemplate(draft_path)
            print(f"Шаблон загружен: {draft_path}")
            return mail_item
            
        except Exception as e:
            print(f"Ошибка загрузки шаблона: {e}")
            return None
    
    def create_email_from_draft(self, draft_path, recipients, subject=None, body_template=None):
        """
        Создание письма на основе шаблона с заменой получателей и текста
        
        Args:
            draft_path: путь к файлу шаблона
            recipients: список получателей (email или имена)
            subject: новая тема письма (опционально)
            body_template: новый текст письма (опционально)
        
        Returns:
            объект MailItem
        """
        mail = self.load_draft_template(draft_path)
        
        if mail is None:
            return None
        
        # Очищаем существующих получателей
        mail.To = ""
        mail.CC = ""
        mail.BCC = ""
        
        # Добавляем новых получателей
        if isinstance(recipients, list):
            mail.To = "; ".join(recipients)
        else:
            mail.To = recipients
        
        # Меняем тему если указана
        if subject:
            mail.Subject = subject
        
        # Меняем тело письма если указано
        if body_template:
            mail.Body = body_template
        
        return mail
    
    def create_custom_email(self, draft_path, recipients, custom_data):
        """
        Создание письма с заменой плейсхолдеров в теле письма
        
        Args:
            draft_path: путь к файлу шаблона
            recipients: список получателей
            custom_data: словарь с данными для замены в тексте
                         например: {"{NAME}": "Иван", "{DATE}": "15.07.2026"}
        
        Returns:
            объект MailItem
        """
        mail = self.load_draft_template(draft_path)
        
        if mail is None:
            return None
        
        # Обновляем получателей
        if isinstance(recipients, list):
            mail.To = "; ".join(recipients)
        else:
            mail.To = recipients
        
        # Заменяем плейсхолдеры в теле письма
        if custom_data and mail.Body:
            body = mail.Body
            for placeholder, value in custom_data.items():
                body = body.replace(placeholder, str(value))
            mail.Body = body
        
        # Заменяем плейсхолдеры в теме
        if custom_data and mail.Subject:
            subject = mail.Subject
            for placeholder, value in custom_data.items():
                subject = subject.replace(placeholder, str(value))
            mail.Subject = subject
        
        return mail
    
    def show_email_window(self, mail_item):
        """
        Отобразить окно письма для редактирования
        
        Args:
            mail_item: объект MailItem
        
        Returns:
            bool: True если письмо было отправлено, False если закрыто без отправки
        """
        if mail_item is None:
            print("Ошибка: письмо не создано")
            return False
        
        try:
            # Отображаем окно письма
            mail_item.Display(False)
            print("Окно письма открыто. Отредактируйте и отправьте или закройте.")
            
            # Ждем пока пользователь закроет окно
            # Примечание: это блокирующий вызов, но мы не можем точно отследить отправку
            # Поэтому просто возвращаем True, предполагая что пользователь отправит письмо
            return True
            
        except Exception as e:
            print(f"Ошибка при отображении письма: {e}")
            return False
    
    def send_email_direct(self, mail_item):
        """
        Отправить письмо без отображения окна
        
        Args:
            mail_item: объект MailItem
        
        Returns:
            bool: True если отправлено успешно
        """
        if mail_item is None:
            print("Ошибка: письмо не создано")
            return False
        
        try:
            mail_item.Send()
            print("Письмо отправлено")
            return True
        except Exception as e:
            print(f"Ошибка при отправке письма: {e}")
            return False


def main():
    """Основная функция с примером использования"""
    
    # Создаем экземпляр класса
    sender = EmailDraftSender()
    
    # Путь к шаблону письма
    draft_path = r"C:\Templates\my_template.msg"  # Измените на ваш путь
    
    # Проверяем существование файла
    if not os.path.exists(draft_path):
        print(f"Файл шаблона не найден: {draft_path}")
        print("Создайте шаблон в Outlook и сохраните его как .msg файл")
        return
    
    # Список получателей
    recipients = [
        "user1@example.com",
        "user2@example.com",
        "user3@example.com"
    ]
    
    # Тема письма (если нужно изменить)
    new_subject = "Важное уведомление от 15.07.2026"
    
    # Новый текст письма (если нужно изменить)
    new_body = """Уважаемые коллеги!
    
    Это автоматическое письмо с обновленной информацией.
    
    С уважением,
    Отдел автоматизации"""
    
    # Вариант 1: Простая замена получателей и текста
    print("\n=== Вариант 1: Простая замена ===")
    mail = sender.create_email_from_draft(
        draft_path=draft_path,
        recipients=recipients,
        subject=new_subject,
        body_template=new_body
    )
    
    if mail:
        # Открываем окно для редактирования и отправки
        sender.show_email_window(mail)
    
    # Вариант 2: Замена плейсхолдеров в шаблоне
    print("\n=== Вариант 2: Замена плейсхолдеров ===")
    
    # Данные для замены в шаблоне
    # В шаблоне должны быть плейсхолдеры: {NAME}, {DATE}, {REPORT}
    custom_data = {
        "{NAME}": "Иван Петров",
        "{DATE}": datetime.now().strftime("%d.%m.%Y"),
        "{REPORT}": "Отчет о продажах за июнь 2026"
    }
    
    mail2 = sender.create_custom_email(
        draft_path=draft_path,
        recipients=["custom@example.com"],
        custom_data=custom_data
    )
    
    if mail2:
        # Открываем окно для редактирования и отправки
        sender.show_email_window(mail2)
    
    # Вариант 3: Отправка без отображения окна
    print("\n=== Вариант 3: Автоматическая отправка ===")
    mail3 = sender.create_email_from_draft(
        draft_path=draft_path,
        recipients=["auto@example.com"],
        subject="Автоматическая отправка",
        body_template="Это письмо отправлено автоматически."
    )
    
    if mail3:
        sender.send_email_direct(mail3)
    
    print("\nСкрипт завершен.")


def interactive_mode():
    """Интерактивный режим с вводом данных пользователем"""
    
    sender = EmailDraftSender()
    
    print("=== Интерактивный режим отправки писем ===")
    print("Для выхода введите 'exit'")
    
    while True:
        print("\n--- Новое письмо ---")
        
        # Ввод пути к шаблону
        draft_path = input("Путь к шаблону (.msg или .oft): ").strip()
        if draft_path.lower() == 'exit':
            break
        
        if not os.path.exists(draft_path):
            print("Файл не найден!")
            continue
        
        # Ввод получателей
        recipients_input = input("Получатели (через запятую): ").strip()
        if recipients_input.lower() == 'exit':
            break
        
        recipients = [r.strip() for r in recipients_input.split(',')]
        
        # Ввод темы
        subject = input("Тема письма (Enter - оставить без изменений): ").strip()
        if subject == '':
            subject = None
        
        # Ввод текста
        print("Текст письма (введите 'END' на новой строке для завершения):")
        lines = []
        while True:
            line = input()
            if line == 'END':
                break
            lines.append(line)
        
        body = '\n'.join(lines) if lines else None
        
        # Создаем письмо
        mail = sender.create_email_from_draft(
            draft_path=draft_path,
            recipients=recipients,
            subject=subject,
            body_template=body
        )
        
        if mail:
            # Отображаем окно
            sender.show_email_window(mail)
            print("Письмо открыто в редакторе. Отредактируйте и отправьте.")
        
        again = input("\nОтправить еще одно письмо? (y/n): ").strip().lower()
        if again != 'y':
            break


if __name__ == "__main__":
    # Выбор режима работы
    print("Выберите режим:")
    print("1 - Автоматический (с примерами)")
    print("2 - Интерактивный (ручной ввод)")
    print("3 - Выход")
    
    choice = input("Ваш выбор (1/2/3): ").strip()
    
    if choice == '1':
        main()
    elif choice == '2':
        interactive_mode()
    else:
        print("Выход")
