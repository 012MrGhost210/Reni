import os
import win32com.client
import pythoncom
import time
from pathlib import Path

class OutlookMailSender:
    def __init__(self):
        """Инициализация подключения к Outlook"""
        self.outlook = None
        
    def connect_to_outlook(self):
        """Подключение к Outlook"""
        try:
            pythoncom.CoInitialize()
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            print("Подключено к Outlook")
            return True
        except Exception as e:
            print(f"Ошибка подключения к Outlook: {e}")
            return False
    
    def send_draft_from_file(self, file_path):
        """
        Отправляет черновик письма из файла .msg или .oft
        
        Args:
            file_path: Путь к файлу письма (.msg, .oft)
        
        Returns:
            bool: Успешно ли отправлено письмо
        """
        try:
            if not os.path.exists(file_path):
                print(f"Файл не найден: {file_path}")
                return False
            
            print(f"Обрабатываю файл: {file_path}")
            
            # Создаем письмо из шаблона
            mail_item = self.outlook.CreateItemFromTemplate(file_path)
            
            # Отправляем письмо
            mail_item.Send()
            
            print(f"Письмо отправлено: {os.path.basename(file_path)}")
            return True
            
        except Exception as e:
            print(f"Ошибка при отправке письма: {e}")
            return False
    
    def find_and_send_drafts(self, folder_path, extensions=None):
        """
        Ищет и отправляет все черновики в указанной папке
        
        Args:
            folder_path: Путь к папке с черновиками
            extensions: Список расширений файлов (по умолчанию ['.msg', '.oft'])
        
        Returns:
            list: Список отправленных файлов
        """
        if extensions is None:
            extensions = ['.msg', '.oft']
        
        sent_files = []
        
        try:
            folder = Path(folder_path)
            
            if not folder.exists():
                print(f"Папка не существует: {folder_path}")
                return sent_files
            
            print(f"Поиск файлов в папке: {folder_path}")
            
            # Ищем файлы с нужными расширениями
            for ext in extensions:
                for file_path in folder.glob(f"*{ext}"):
                    if self.send_draft_from_file(str(file_path)):
                        sent_files.append(str(file_path))
                        
                        # Удаляем файл после отправки (раскомментировать если нужно)
                        # try:
                        #     os.remove(str(file_path))
                        #     print(f"Файл удален: {file_path.name}")
                        # except Exception as e:
                        #     print(f"Не удалось удалить файл: {e}")
            
            return sent_files
            
        except Exception as e:
            print(f"Ошибка при поиске файлов: {e}")
            return sent_files
    
    def close(self):
        """Закрытие соединения"""
        try:
            if self.outlook:
                self.outlook = None
            pythoncom.CoUninitialize()
            print("Соединение закрыто")
        except:
            pass

def main():
    """Основная функция"""
    # Укажите путь к папке с черновиками
    drafts_folder = r"C:\Users\Username\Documents\Outlook Drafts"
    
    # Или используйте относительный путь
    # drafts_folder = os.path.join(os.path.dirname(__file__), "drafts")
    
    # Создаем экземпляр отправителя
    sender = OutlookMailSender()
    
    try:
        # Подключаемся к Outlook
        if not sender.connect_to_outlook():
            print("Не удалось подключиться к Outlook. Проверьте, что Outlook установлен и запущен.")
            return
        
        # Создаем папку если ее нет
        os.makedirs(drafts_folder, exist_ok=True)
        
        print(f"Папка для черновиков: {drafts_folder}")
        print("Поместите файлы .msg или .oft в эту папку")
        print("Нажмите Enter для поиска и отправки писем...")
        input()
        
        # Ищем и отправляем черновики
        sent_files = sender.find_and_send_drafts(drafts_folder)
        
        if sent_files:
            print(f"\nУспешно отправлено {len(sent_files)} писем:")
            for file in sent_files:
                print(f"  - {os.path.basename(file)}")
        else:
            print("\nФайлы для отправки не найдены")
            
    except KeyboardInterrupt:
        print("\nПрограмма прервана пользователем")
    except Exception as e:
        print(f"\nПроизошла ошибка: {e}")
    finally:
        # Закрываем соединение
        sender.close()

if __name__ == "__main__":
    main()
