import win32com.client as win32

def send_test_email(recipients):
    """
    Отправляет тестовое сообщение указанным адресатам через Outlook
    """
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    
    mail.To = recipients
    mail.Subject = "Тестовое сообщение"
    mail.Body = "Это тестовое сообщение, отправленное через Python."
    
    mail.Send()
    print(f"Письмо отправлено адресатам: {recipients}")

# Пример использования с разными адресатами
send_test_email("user1@example.com")  # одному адресату
# или нескольким сразу
# send_test_email("user1@example.com; user2@example.com; user3@example.com")
