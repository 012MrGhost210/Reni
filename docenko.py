import os
import shutil
from pathlib import Path
import win32com.client as win32
from datetime import datetime

def send_outlook_message(recipients):
    """
    Отправляет простое уведомление через Outlook
    """
    try:
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        
        mail.To = recipients
        mail.Subject = "Выгрузка данных в Диадок"
        
        # Только дата, время и сообщение о выгрузке
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        mail.Body = f"{current_time} - все выгружено"
        
        mail.Send()
        print(f"✅ Уведомление отправлено: {recipients}")
        return True
    except Exception as e:
        print(f"❌ Ошибка при отправке email: {e}")
        return False

def copy_contents_from_nested_folders(source_root, target_base, company_folders, email_recipients):
    """
    Копирует всё содержимое из указанных папок (company_folders),
    которые могут лежать на любом уровне вложенности внутри source_root,
    в соответствующие подпапки внутри target_base.
    """
    source_root = Path(source_root)
    target_base = Path(target_base)
    
    # Создаем целевую папку, если её нет
    target_base.mkdir(parents=True, exist_ok=True)
    
    print(f"Поиск в: {source_root}")
    print(f"Целевая папка: {target_base}")
    print(f"Ищем папки: {company_folders}")
    print("-" * 50)
    
    copied_count = 0
    skipped_count = 0
    
    for company_folder_name in company_folders:
        # Создаем соответствующую папку в целевой директории
        company_target_path = target_base / company_folder_name
        company_target_path.mkdir(parents=True, exist_ok=True)
        
        # Ищем все папки с таким названием в исходной директории
        found_paths = list(source_root.rglob(company_folder_name))
        
        if not found_paths:
            print(f"Предупреждение: Папка '{company_folder_name}' не найдена.")
            continue
        
        print(f"\nОбработка папок для: {company_folder_name}")
        
        for company_path in found_paths:
            # Проверяем, что это действительно папка, а не файл с таким именем
            if not company_path.is_dir():
                print(f"  Пропуск: {company_path} — это не папка.")
                continue
            
            print(f"  Найдена папка: {company_path}")
            
            # Перебираем всё содержимое найденной папки компании
            for item_path in company_path.iterdir():
                # Формируем путь назначения в соответствующей подпапке
                destination_path = company_target_path / item_path.name
                
                try:
                    # Если элемент - файл
                    if item_path.is_file():
                        # Копируем файл с заменой существующего
                        shutil.copy2(item_path, destination_path)
                        print(f"    Скопирован файл: {item_path.name}")
                        copied_count += 1
                    
                    # Если элемент - папка
                    elif item_path.is_dir():
                        # Копируем всю папку рекурсивно (с заменой существующей)
                        if destination_path.exists():
                            shutil.rmtree(destination_path)
                        shutil.copytree(item_path, destination_path)
                        print(f"    Скопирована папка: {item_path.name}/")
                        copied_count += 1
                    
                except Exception as e:
                    print(f"    Ошибка при копировании {item_path.name}: {e}")
    
    print("-" * 50)
    print(f"Готово! Скопировано элементов: {copied_count}")
    
    # Отправка простого уведомления
    if email_recipients:
        send_outlook_message(email_recipients)

if __name__ == "__main__":
    # ВАШИ ПУТИ
    source_folder = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\diadoc_connector"
    target_folder = r"\\fs-01.renlife.com\alldocs\Финансовый департамент\Treasury\Базы данных(автоматизация)\Диадок"
    
    # Список точных названий папок компаний
    companies = [
        "7702358512-ООО -УК Райффайзен-",
        "7744000951-АО -УК -СПУТНИК - УПРАВЛЕНИЕ КАПИТАЛОМ-",
        "7825489723-ТКБ Инвестмент Партнерс (АО)"
    ]
    
    # Email получателей (можно несколько через точку с запятой)
    email_recipients = "your.email@company.com; manager@company.com"
    
    # Запуск
    copy_contents_from_nested_folders(source_folder, target_folder, companies, email_recipients)
