import os
import re
import shutil
import zipfile
from datetime import datetime
from pathlib import Path


def find_today_folder(base_path: str) -> str | None:
    """
    Ищет папку с сегодняшней датой в формате ГГГГ_ММ_ДД
    """
    today_str = datetime.now().strftime("%Y_%m_%d")
    today_folder = os.path.join(base_path, today_str)
    
    if os.path.exists(today_folder) and os.path.isdir(today_folder):
        return today_folder
    return None


def find_documents_folder(base_folder: str) -> str | None:
    """
    Ищет папку 'Документы от Гаранта СД НТД' в указанной директории
    """
    target_name = "Документы от Гаранта СД НТД"
    
    for item in os.listdir(base_folder):
        item_path = os.path.join(base_folder, item)
        if os.path.isdir(item_path) and item == target_name:
            return item_path
    
    return None


def find_zip_files(folder: str) -> dict:
    """
    Ищет zip-архивы с маркерами ПР и ПН
    Возвращает словарь {маркер: полный_путь_к_архиву}
    """
    zip_files = {}
    pattern_pr = re.compile(r"ПР", re.IGNORECASE)
    pattern_pn = re.compile(r"ПН", re.IGNORECASE)
    
    for file in os.listdir(folder):
        if file.lower().endswith('.zip'):
            file_path = os.path.join(folder, file)
            
            # Проверяем наличие маркеров
            if pattern_pr.search(file):
                zip_files['ПР'] = file_path
            elif pattern_pn.search(file):
                zip_files['ПН'] = file_path
    
    return zip_files


def extract_file_from_zip(zip_path: str, filename_pattern: str, destination: str) -> bool:
    """
    Извлекает файл из архива по шаблону имени и копирует в destination
    """
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            # Ищем файл по шаблону
            for file in zip_ref.namelist():
                if filename_pattern in file and file.lower().endswith('.xls'):
                    # Извлекаем файл во временную директорию
                    temp_dir = os.path.dirname(destination)
                    os.makedirs(temp_dir, exist_ok=True)
                    
                    # Извлекаем файл
                    extracted_path = zip_ref.extract(file, temp_dir)
                    
                    # Получаем имя извлеченного файла
                    extracted_filename = os.path.basename(extracted_path)
                    final_dest = os.path.join(temp_dir, destination)
                    
                    # Перемещаем/переименовываем файл в нужное место
                    shutil.move(extracted_path, final_dest)
                    print(f"Файл успешно извлечен и сохранен: {final_dest}")
                    return True
                    
        print(f"Файл с шаблоном '{filename_pattern}' не найден в архиве {zip_path}")
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


def get_second_part_from_pr_archive(zip_path: str) -> str:
    """
    Извлекает вторую часть имени (цифры после подчеркивания) из архива ПР
    """
    filename = os.path.basename(zip_path)
    # Ищем последовательность цифр после последнего подчеркивания
    pattern = re.compile(r'_(\d+)\.zip$')
    match = pattern.search(filename)
    
    if match:
        return match.group(1)
    return ""


def main():
    # Пути (замените на свои)
    path_x = r"Q:\Финансовый отдел\01.Перечень имущества Фонда (СД)"  # Путь для поиска папки с датой
    path_y = r"\\fs-01.renlife.com\alldocs\Финансовый департамент\Treasury\18. НПФ\1. Отчеты\1.1 Ежедневные отчеты\СПУТНИК\Акутальные"  # Базовый путь для файла ПР
    path_z = r"\\fs-01.renlife.com\alldocs\Финансовый департамент\Treasury\18. НПФ\1. Отчеты\1.1 Ежедневные отчеты\ФОНД\Актуальные данные"  # Базовый путь для второго файла ПР (с другой датой)
    path_i = r"\\fs-01.renlife.com\alldocs\Финансовый департамент\Treasury\18. НПФ\1. Отчеты\1.1 Ежедневные отчеты\ВИМ"
    
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
        return
    
    print(f"Найдены архивы: {zip_files}")
    
    # Шаг 4: Обработка архива ПР
    if 'ПР' in zip_files:
        pr_zip = zip_files['ПР']
        pr_filename = os.path.basename(pr_zip)
        
        # Извлекаем дату из имени файла
        date_str = get_date_from_filename(pr_filename)
        
        # Формируем имя целевого файла
        target_filename = f"{date_str}_СЧА УК СПУТНИК - УПРАВЛЕНИЕ КАПИТАЛОМ (Д.У. 301024_1).xls"
        dest_path_y = os.path.join(path_y, target_filename)
        
        # Извлекаем файл
        print(f"Извлечение из {pr_zip} в {dest_path_y}")
        extract_file_from_zip(pr_zip, "СЧА", dest_path_y)
        
        # Для второго пути (z) с другой датой
        # Используем ту же дату, но другой ID (080825_1)
        target_filename_z = f"{date_str}_СЧА УК СПУТНИК - УПРАВЛЕНИЕ КАПИТАЛОМ (Д.У. 080825_1).xls"
        dest_path_z = os.path.join(path_z, target_filename_z)
        
        print(f"Извлечение из {pr_zip} в {dest_path_z}")
        extract_file_from_zip(pr_zip, "СЧА", dest_path_z)
    else:
        print("Архив ПР не найден")
    
    # Шаг 5: Обработка архива ПН
    if 'ПН' in zip_files:
        pn_zip = zip_files['ПН']
        pn_filename = os.path.basename(pn_zip)
        
        # Извлекаем дату из имени файла
        date_str = get_date_from_filename(pn_filename)
        
        # Формируем имя целевого файла
        target_filename = f"{date_str}_Расчет стоимости активов ПН с учетом портфеля НПФ.xls"
        
        # Путь для сохранения (можно задать другой путь)
        dest_path_pn = os.path.join(path_y, target_filename)  # или другой путь
        
        print(f"Извлечение из {pn_zip} в {dest_path_pn}")
        extract_file_from_zip(pn_zip, "Расчётстоимости активов", dest_path_pn)
    else:
        print("Архив ПН не найден")


if __name__ == "__main__":
    main()

Traceback (most recent call last):
  File "c:/Users/Ilya.Matveev2/Скрипты/SCHA_NPF.py", line 9, in <module>
    def find_today_folder(base_path: str) -> str | None:
TypeError: unsupported operand type(s) for |: 'type' and 'NoneType'
