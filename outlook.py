import os
import win32com.client
import pythoncom
from pathlib import Path
from datetime import datetime
import shutil

class OutlookDraftEditor:
    def __init__(self):
        """Инициализация подключения к Outlook"""
        self.outlook = None
        
    def connect_to_outlook(self):
        """Подключение к Outlook"""
        try:
            pythoncom.CoInitialize()
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            print("✓ Подключено к Outlook")
            return True
        except Exception as e:
            print(f"✗ Ошибка подключения к Outlook: {e}")
            return False
    
    def modify_and_send_draft(self, file_path, modifications=None):
        """
        Изменяет черновик письма и отправляет его
        
        Args:
            file_path: Путь к файлу письма
            modifications: Словарь с изменениями {
                'subject': 'Новая тема',
                'body': 'Новый текст',
                'to': 'новый@email.com',  # получатели (через ;)
                'cc': 'копия@email.com',   # копия
                'bcc': 'скрытая@email.com', # скрытая копия
                'attachments_add': ['путь/к/файлу1.pdf', 'путь/к/файлу2.docx'],  # добавить вложения
                'attachments_remove': ['старый_файл.pdf'],  # удалить вложения (по имени)
                'attachments_replace': [  # заменить вложения
                    {'old': 'старый.pdf', 'new': 'новый.pdf'}
                ],
                'body_append': 'Текст в конец',  # добавить текст в конец
                'body_prepend': 'Текст в начало',  # добавить текст в начало
                'body_replace': {'старое': 'новое'}  # замена текста
            }
        """
        try:
            if not os.path.exists(file_path):
                print(f"✗ Файл не найден: {file_path}")
                return False
            
            print(f"\n--- Обрабатываю файл: {os.path.basename(file_path)} ---")
            
            # Открываем письмо из файла
            mail = self.outlook.CreateItemFromTemplate(file_path)
            
            if modifications:
                self._apply_modifications(mail, modifications)
            
            # Отправляем письмо
            mail.Send()
            print(f"✓ Письмо успешно отправлено!")
            return True
            
        except Exception as e:
            print(f"✗ Ошибка при обработке письма: {e}")
            return False
    
    def _apply_modifications(self, mail, mods):
        """Применяет изменения к письму"""
        
        # 1. Изменение темы
        if 'subject' in mods and mods['subject']:
            old_subject = mail.Subject
            mail.Subject = mods['subject']
            print(f"  • Тема изменена: '{old_subject}' -> '{mail.Subject}'")
        
        # 2. Изменение получателей
        if 'to' in mods and mods['to']:
            mail.To = mods['to']
            print(f"  • Получатели: {mail.To}")
        
        if 'cc' in mods and mods['cc']:
            mail.CC = mods['cc']
            print(f"  • Копия: {mail.CC}")
            
        if 'bcc' in mods and mods['bcc']:
            mail.BCC = mods['bcc']
            print(f"  • Скрытая копия: {mail.BCC}")
        
        # 3. Изменение тела письма
        current_body = mail.Body
        
        if 'body' in mods and mods['body']:
            # Полная замена текста
            mail.Body = mods['body']
            print(f"  • Текст полностью заменен")
            
        else:
            # Частичные изменения
            new_body = current_body
            
            # Добавление в начало
            if 'body_prepend' in mods and mods['body_prepend']:
                new_body = mods['body_prepend'] + "\n\n" + new_body
                print(f"  • Текст добавлен в начало")
            
            # Добавление в конец
            if 'body_append' in mods and mods['body_append']:
                new_body = new_body + "\n\n" + mods['body_append']
                print(f"  • Текст добавлен в конец")
            
            # Замена подстрок
            if 'body_replace' in mods and mods['body_replace']:
                for old, new in mods['body_replace'].items():
                    if old in new_body:
                        new_body = new_body.replace(old, new)
                        print(f"  • Замена: '{old}' -> '{new}'")
            
            mail.Body = new_body
        
        # 4. Работа с вложениями
        # Удаление вложений
        if 'attachments_remove' in mods and mods['attachments_remove']:
            attachments_to_remove = []
            for attachment in mail.Attachments:
                for name_to_remove in mods['attachments_remove']:
                    if name_to_remove.lower() in attachment.FileName.lower():
                        attachments_to_remove.append(attachment.FileName)
                        mail.Attachments.Remove(attachment.Index)
            
            if attachments_to_remove:
                print(f"  • Удалены вложения: {', '.join(attachments_to_remove)}")
        
        # Добавление вложений
        if 'attachments_add' in mods and mods['attachments_add']:
            added = []
            not_found = []
            for file_path in mods['attachments_add']:
                if os.path.exists(file_path):
                    mail.Attachments.Add(file_path)
                    added.append(os.path.basename(file_path))
                else:
                    not_found.append(file_path)
            
            if added:
                print(f"  • Добавлены вложения: {', '.join(added)}")
            if not_found:
                print(f"  ✗ Не найдены файлы: {', '.join(not_found)}")
        
        # Замена вложений
        if 'attachments_replace' in mods and mods['attachments_replace']:
            for replacement in mods['attachments_replace']:
                old_name = replacement.get('old', '')
                new_path = replacement.get('new', '')
                
                if os.path.exists(new_path):
                    # Ищем и удаляем старое вложение
                    for attachment in mail.Attachments:
                        if old_name.lower() in attachment.FileName.lower():
                            mail.Attachments.Remove(attachment.Index)
                            break
                    
                    # Добавляем новое
                    mail.Attachments.Add(new_path)
                    print(f"  • Вложение заменено: {old_name} -> {os.path.basename(new_path)}")
                else:
                    print(f"  ✗ Файл для замены не найден: {new_path}")
    
    def create_from_template(self, template_path, modifications=None, send_immediately=True):
        """
        Создает письмо из шаблона, применяет изменения и отправляет/сохраняет
        
        Args:
            template_path: Путь к файлу шаблона (.msg, .oft)
            modifications: Словарь с изменениями
            send_immediately: Отправить сразу (True) или сохранить как черновик (False)
        
        Returns:
            bool: Успешно ли выполнено
        """
        try:
            if not os.path.exists(template_path):
                print(f"✗ Шаблон не найден: {template_path}")
                return False
            
            print(f"\n--- Создаю письмо из шаблона: {os.path.basename(template_path)} ---")
            
            # Создаем письмо из шаблона
            mail = self.outlook.CreateItemFromTemplate(template_path)
            
            # Применяем изменения
            if modifications:
                self._apply_modifications(mail, modifications)
            
            if send_immediately:
                mail.Send()
                print(f"✓ Письмо отправлено!")
            else:
                # Сохраняем как черновик
                draft_folder = os.path.join(os.path.dirname(template_path), "drafts")
                os.makedirs(draft_folder, exist_ok=True)
                
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                draft_path = os.path.join(draft_folder, f"draft_{timestamp}.msg")
                
                mail.SaveAs(draft_path)
                print(f"✓ Черновик сохранен: {draft_path}")
            
            return True
            
        except Exception as e:
            print(f"✗ Ошибка: {e}")
            return False
    
    def batch_process(self, folder_path, modifications=None, pattern="*.msg"):
        """
        Пакетная обработка нескольких писем
        
        Args:
            folder_path: Папка с письмами
            modifications: Изменения для всех писем
            pattern: Шаблон поиска файлов
        
        Returns:
            tuple: (успешно, с ошибками)
        """
        successful = []
        failed = []
        
        folder = Path(folder_path)
        if not folder.exists():
            print(f"✗ Папка не существует: {folder_path}")
            return successful, failed
        
        print(f"\n=== Пакетная обработка писем в папке: {folder_path} ===")
        
        for file_path in folder.glob(pattern):
            if self.modify_and_send_draft(str(file_path), modifications):
                successful.append(str(file_path))
            else:
                failed.append(str(file_path))
        
        print(f"\n=== Результаты ===")
        print(f"✓ Успешно: {len(successful)}")
        print(f"✗ С ошибками: {len(failed)}")
        
        return successful, failed
    
    def close(self):
        """Закрытие соединения"""
        try:
            if self.outlook:
                self.outlook = None
            pythoncom.CoUninitialize()
            print("\nСоединение закрыто")
        except:
            pass

def main():
    """Примеры использования"""
    
    # Создаем экземпляр редактора
    editor = OutlookDraftEditor()
    
    try:
        # Подключаемся к Outlook
        if not editor.connect_to_outlook():
            print("Не удалось подключиться к Outlook")
            return
        
        # ===== ПРИМЕР 1: Простое изменение и отправка =====
        modifications1 = {
            'subject': 'Обновленное письмо с новыми данными',
            'body_append': 'С уважением,\nАвтоматическая рассылка',
            'to': 'ivanov@company.ru; petrov@company.ru',
            'attachments_add': [
                r'C:\Users\Username\Documents\report.pdf',
                r'C:\Users\Username\Documents\data.xlsx'
            ]
        }
        
        # ===== ПРИМЕР 2: Персонализированное письмо =====
        client_name = "ООО Ромашка"
        current_date = datetime.now().strftime("%d.%m.%Y")
        
        modifications2 = {
            'subject': f'Коммерческое предложение для {client_name}',
            'body_replace': {
                '{CLIENT_NAME}': client_name,
                '{CURRENT_DATE}': current_date,
                '{MANAGER_NAME}': 'Иван Петров'
            },
            'body_append': '\n\nДанное письмо сформировано автоматически.',
            'attachments_replace': [
                {'old': 'template_price.pdf', 'new': r'C:\Users\Username\Documents\price_{client_name}.pdf'}
            ]
        }
        
        # ===== ПРИМЕР 3: Массовая рассылка =====
        clients = [
            {'name': 'ООО Альфа', 'email': 'alfa@company.ru'},
            {'name': 'ООО Бета', 'email': 'beta@company.ru'},
            {'name': 'ООО Гамма', 'email': 'gamma@company.ru'}
        ]
        
        for client in clients:
            client_mods = {
                'subject': f'Специальное предложение для {client["name"]}',
                'to': client['email'],
                'body_replace': {
                    '{CLIENT_NAME}': client['name'],
                    '{DISCOUNT}': '15%'
                }
            }
            
            # Используем общий шаблон
            editor.create_from_template(
                r'C:\Users\Username\Documents\templates\offer_template.msg',
                client_mods,
                send_immediately=True
            )
        
        # ===== ПРИМЕР 4: Пакетная обработка =====
        # Отправляем все письма из папки с добавлением стандартной подписи
        batch_mods = {
            'body_append': '\n\n---\nС уважением,\nОтдел продаж\nТел: +7 (999) 123-45-67',
            'attachments_add': [r'C:\Users\Username\Documents\catalog.pdf']
        }
        
        # editor.batch_process(r'C:\Users\Username\Documents\Outlook Drafts', batch_mods)
        
        # ===== ПРИМЕР 5: Создание черновика без отправки =====
        draft_mods = {
            'subject': 'Черновик для проверки',
            'body': 'Это черновик письма, нужно проверить перед отправкой',
            'to': 'reviewer@company.ru'
        }
        
        editor.create_from_template(
            r'C:\Users\Username\Documents\templates\blank.msg',
            draft_mods,
            send_immediately=False
        )
        
        # ===== ИНТЕРАКТИВНЫЙ РЕЖИМ =====
        print("\n" + "="*50)
        print("ИНТЕРАКТИВНЫЙ РЕЖИМ")
        print("="*50)
        
        # Здесь можно добавить логику для интерактивной работы
        # Например, чтение данных из Excel файла
        
    except KeyboardInterrupt:
        print("\nПрограмма прервана пользователем")
    except Exception as e:
        print(f"\nПроизошла ошибка: {e}")
    finally:
        editor.close()

if __name__ == "__main__":
    main()
