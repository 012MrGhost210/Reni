import os
import re
import shutil
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict


def fix_encoding(filename: str) -> str:
    """
    Исправляет кодировку имени файла из Windows-1251 в UTF-8
    """
    try:
        # Пробуем перекодировать из Windows-1251 в UTF-8
        return filename.encode('cp437').decode('cp1251')
    except:
        try:
            # Если не получилось, пробуем другие варианты
            return filename.encode('latin-1').decode('cp1251')
        except:
            return filename


def find_today_folder(base_path: str) -> Optional[str]:
    """
    Ищет папку с сегодняшней датой в формате ГГГГ_ММ_ДД
    """
    today_str = datetime.now().strftime("%Y_%m_%d")
    today_folder = os.path.join(base_path, today_str)
    
    if os.path.exists(today_folder) and os.path.isdir(today_folder):
        return today_folder
    return None


def find_documents_folder(base_folder: str) -> Optional[str]:
    """
    Ищет папку 'Документы от Гаранта СД НТД' в указанной директории
    """
    target_name = "Документы от Гаранта СД НТД"
    
    try:
        for item in os.listdir(base_folder):
            item_path = os.path.join(base_folder, item)
            if os.path.isdir(item_path) and item == target_name:
                return item_path
    except PermissionError:
        print(f"Нет доступа к папке {base_folder}")
        return None
    
    return None


def find_zip_files(folder: str) -> Dict[str, str]:
    """
    Ищет zip-архивы с маркерами ПР и ПН
    Возвращает словарь {маркер: полный_путь_к_архиву}
    """
    zip_files = {}
    pattern_pr = re.compile(r"ПР", re.IGNORECASE)
    pattern_pn = re.compile(r"ПН", re.IGNORECASE)
    
    try:
        for file in os.listdir(folder):
            if file.lower().endswith('.zip'):
                file_path = os.path.join(folder, file)
                
                # Проверяем наличие маркеров
                if pattern_pr.search(file):
                    zip_files['ПР'] = file_path
                elif pattern_pn.search(file):
                    zip_files['ПН'] = file_path
    except PermissionError:
        print(f"Нет доступа к папке {folder}")
        return {}
    
    return zip_files


def find_file_in_zip(zip_ref: zipfile.ZipFile, patterns: list) -> Optional[str]:
    """
    Ищет файл в архиве по списку паттернов с учетом кодировки
    """
    for file in zip_ref.namelist():
        # Пробуем разные варианты имени файла
        original_name = file
        decoded_name = fix_encoding(file)
        
        # Проверяем оба варианта имени
        for pattern in patterns:
            if pattern in original_name or pattern in decoded_name:
                # Проверяем, что это Excel файл
                if file.lower().endswith(('.xls', '.xlsx')):
                    return file
    
    return None


def extract_file_from_zip(zip_path: str, filename_patterns: list, destination: str) -> bool:
    """
    Извлекает файл из архива по шаблону имени и копирует в destination
    """
    try:
        # Проверяем, существует ли архив
        if not os.path.exists(zip_path):
            print(f"Архив не найден: {zip_path}")
            return False
            
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            # Ищем файл по шаблонам
            file_to_extract = find_file_in_zip(zip_ref, filename_patterns)
            
            if not file_to_extract:
                print(f"Файл с шаблонами {filename_patterns} не найден в архиве {zip_path}")
                print(f"Доступные файлы в архиве:")
                for f in zip_ref.namelist():
                    if f.lower().endswith(('.xls', '.xlsx')):
                        decoded = fix_encoding(f)
                        print(f"  - {decoded}")
                return False
            
            # Создаем директорию назначения
            dest_dir = os.path.dirname(destination)
            os.makedirs(dest_dir, exist_ok=True)
            
            # Извлекаем файл
            extracted_path = zip_ref.extract(file_to_extract, dest_dir)
            
            # Переименовываем файл
            final_dest = destination
            shutil.move(extracted_path, final_dest)
            print(f"Файл успешно извлечен и сохранен: {final_dest}")
            return True
                    
    except zipfile.BadZipFile:
        print(f"Ошибка: файл {zip_path} не является zip-архивом или поврежден")
        return False
    except Exception as e:
        print(f"Ошибка при работе с архивом {zip_path}: {e}")
        return False


def get_date_from_filename(filename: str) -> str:
    """
    Извлекает дату из имени файла (формат 2026-06-26 -> 26.06.2026)
    """
    # Ищем дату в формате ГГГГ-ММ-ДД
    date_pattern = re.compile(r'(\d{4})-(\d{2})-(\d{2})')
    match = date_pattern.search(filename)
    
    if match:
        year, month, day = match.groups()
        return f"{day}.{month}.{year}"
    
    # Если не найдено, используем текущую дату
    return datetime.now().strftime("%d.%m.%Y")


def main():
    # Пути
    path_x = r"Q:\Финансовый отдел\01.Перечень имущества Фонда (СД)"
    path_y = r"\\fs-01.renlife.com\alldocs\Финансовый департамент\Treasury\18. НПФ\1. Отчеты\1.1 Ежедневные отчеты\СПУТНИК\Акутальные"
    path_z = r"\\fs-01.renlife.com\alldocs\Финансовый департамент\Treasury\18. НПФ\1. Отчеты\1.1 Ежедневные отчеты\ФОНД\Актуальные данные"
    path_i = r"\\fs-01.renlife.com\alldocs\Финансовый департамент\Treasury\18. НПФ\1. Отчеты\1.1 Ежедневные отчеты\ВИМ"
    
    print("Начинаем обработку...")
    print(f"Поиск в папке: {path_x}")
    
    # Шаг 1: Находим папку с сегодняшней датой
    today_folder = find_today_folder(path_x)
    if not today_folder:
        print(f"Папка с сегодняшней датой не найдена в {path_x}")
        return
    
    print(f"Найдена папка: {today_folder}")
    
    # Шаг 2: Находим папку "Документы от Гаранта СД НТД"
    docs_folder = find_documents_folder(today_folder)
    if not docs_folder:
        print(f"Папка 'Документы от Гаранта СД НТД' не найдена в {today_folder}")
        return
    
    print(f"Найдена папка: {docs_folder}")
    
    # Шаг 3: Находим zip-архивы
    zip_files = find_zip_files(docs_folder)
    if not zip_files:
        print("Архивы с маркерами ПР или ПН не найдены")
        try:
            print(f"Файлы в папке {docs_folder}:")
            for file in os.listdir(docs_folder):
                print(f"  - {file}")
        except:
            pass
        return
    
    print(f"Найдены архивы: {zip_files}")
    
    # Шаг 4: Обработка архива ПР
    if 'ПР' in zip_files:
        pr_zip = zip_files['ПР']
        pr_filename = os.path.basename(pr_zip)
        
        # Извлекаем дату из имени файла
        date_str = get_date_from_filename(pr_filename)
        print(f"Извлеченная дата для ПР: {date_str}")
        
        # Формируем имя целевого файла
        target_filename_y = f"{date_str}_СЧА УК СПУТНИК - УПРАВЛЕНИЕ КАПИТАЛОМ (Д.У. 301024_1).xls"
        dest_path_y = os.path.join(path_y, target_filename_y)
        
        # Извлекаем файл с разными вариантами поиска
        print(f"Извлечение из {pr_zip} в {dest_path_y}")
        patterns = ["СЧА", "УПРАВЛЕНИЕ КАПИТАЛОМ", "301024"]
        extract_file_from_zip(pr_zip, patterns, dest_path_y)
        
        # Для второго пути (z) с другой датой
        target_filename_z = f"{date_str}_СЧА УК СПУТНИК - УПРАВЛЕНИЕ КАПИТАЛОМ (Д.У. 080825_1).xls"
        dest_path_z = os.path.join(path_z, target_filename_z)
        
        print(f"Извлечение из {pr_zip} в {dest_path_z}")
        patterns_z = ["СЧА", "УПРАВЛЕНИЕ КАПИТАЛОМ", "080825"]
        extract_file_from_zip(pr_zip, patterns_z, dest_path_z)
    else:
        print("Архив ПР не найден")
    
    # Шаг 5: Обработка архива ПН
    if 'ПН' in zip_files:
        pn_zip = zip_files['ПН']
        pn_filename = os.path.basename(pn_zip)
        
        # Извлекаем дату из имени файла
        date_str = get_date_from_filename(pn_filename)
        print(f"Извлеченная дата для ПН: {date_str}")
        
        # Формируем имя целевого файла
        target_filename = f"{date_str}_Расчет стоимости активов ПН с учетом портфеля НПФ.xls"
        dest_path_pn = os.path.join(path_i, target_filename)
        
        print(f"Извлечение из {pn_zip} в {dest_path_pn}")
        # Для ПН ищем файл с "Расчет" или "стоимости активов"
        patterns_pn = ["Расчет", "стоимости активов", "НПФ"]
        extract_file_from_zip(pn_zip, patterns_pn, dest_path_pn)
    else:
        print("Архив ПН не найден")
    
    print("Обработка завершена!")


if __name__ == "__main__":
    main()
