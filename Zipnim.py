import os
import re
import shutil
import zipfile
from datetime import datetime
from typing import Optional, Dict


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
    
    try:
        for file in os.listdir(folder):
            if file.lower().endswith('.zip'):
                file_path = os.path.join(folder, file)
                
                # Проверяем наличие маркеров в имени файла
                if 'ПР' in file:
                    zip_files['ПР'] = file_path
                elif 'ПН' in file:
                    zip_files['ПН'] = file_path
    except PermissionError:
        print(f"Нет доступа к папке {folder}")
        return {}
    
    return zip_files


def get_available_files(zip_path: str) -> list:
    """
    Получает список всех файлов в архиве с их байтовыми представлениями
    """
    files = []
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            for name in zip_ref.namelist():
                if name.lower().endswith(('.xls', '.xlsx')):
                    files.append(name)
    except Exception as e:
        print(f"Ошибка при чтении архива: {e}")
    return files


def find_file_by_patterns(zip_path: str, patterns: list) -> Optional[str]:
    """
    Ищет файл в архиве по списку паттернов
    Использует прямое сравнение байтов
    """
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            for name in zip_ref.namelist():
                if not name.lower().endswith(('.xls', '.xlsx')):
                    continue
                
                # Проверяем каждый паттерн
                for pattern in patterns:
                    # Пробуем разные варианты сравнения
                    if pattern in name:
                        print(f"Найден файл по паттерну '{pattern}': {name}")
                        return name
                    
                    # Пробуем сравнить без учета регистра
                    if pattern.lower() in name.lower():
                        print(f"Найден файл по паттерну '{pattern}' (без учета регистра): {name}")
                        return name
                    
                    # Пробуем искать в байтовом представлении
                    try:
                        # Конвертируем имя в байты и обратно для поиска
                        name_bytes = name.encode('latin-1')
                        pattern_bytes = pattern.encode('latin-1')
                        if pattern_bytes in name_bytes:
                            print(f"Найден файл по паттерну '{pattern}' (байтовое сравнение): {name}")
                            return name
                    except:
                        pass
    except Exception as e:
        print(f"Ошибка при поиске в архиве: {e}")
    
    return None


def extract_and_rename(zip_path: str, source_filename: str, destination: str) -> bool:
    """
    Извлекает файл из архива и переименовывает его
    """
    try:
        # Создаем директорию назначения
        dest_dir = os.path.dirname(destination)
        os.makedirs(dest_dir, exist_ok=True)
        
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            # Извлекаем файл
            extracted_path = zip_ref.extract(source_filename, dest_dir)
            
            # Переименовываем
            shutil.move(extracted_path, destination)
            print(f"Файл успешно извлечен и сохранен: {destination}")
            return True
            
    except Exception as e:
        print(f"Ошибка при извлечении файла: {e}")
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
        return
    
    print(f"Найдены архивы: {zip_files}")
    
    # Шаг 4: Обработка архива ПР
    if 'ПР' in zip_files:
        pr_zip = zip_files['ПР']
        pr_filename = os.path.basename(pr_zip)
        
        # Извлекаем дату из имени файла
        date_str = get_date_from_filename(pr_filename)
        print(f"Извлеченная дата для ПР: {date_str}")
        
        # Получаем список файлов в архиве
        print("Сканируем архив ПР...")
        files_in_zip = get_available_files(pr_zip)
        print(f"Найдено {len(files_in_zip)} файлов в архиве")
        
        # Ищем файл с 301024
        target_file_301 = None
        for f in files_in_zip:
            if '301024' in f:
                target_file_301 = f
                print(f"Найден файл для 301024: {f}")
                break
        
        if target_file_301:
            # Формируем имя целевого файла для path_y
            target_filename_y = f"{date_str}_СЧА УК СПУТНИК - УПРАВЛЕНИЕ КАПИТАЛОМ (Д.У. 301024_1).xls"
            dest_path_y = os.path.join(path_y, target_filename_y)
            
            # Извлекаем файл
            print(f"Извлечение файла для 301024 в {dest_path_y}")
            extract_and_rename(pr_zip, target_file_301, dest_path_y)
        else:
            print("Файл с 301024 не найден в архиве ПР")
        
        # Ищем файл с 080825
        target_file_080 = None
        for f in files_in_zip:
            if '080825' in f:
                target_file_080 = f
                print(f"Найден файл для 080825: {f}")
                break
        
        if target_file_080:
            # Формируем имя целевого файла для path_z
            target_filename_z = f"{date_str}_СЧА УК СПУТНИК - УПРАВЛЕНИЕ КАПИТАЛОМ (Д.У. 080825_1).xls"
            dest_path_z = os.path.join(path_z, target_filename_z)
            
            # Извлекаем файл
            print(f"Извлечение файла для 080825 в {dest_path_z}")
            extract_and_rename(pr_zip, target_file_080, dest_path_z)
        else:
            print("Файл с 080825 не найден в архиве ПР")
    else:
        print("Архив ПР не найден")
    
    # Шаг 5: Обработка архива ПН
    if 'ПН' in zip_files:
        pn_zip = zip_files['ПН']
        pn_filename = os.path.basename(pn_zip)
        
        # Извлекаем дату из имени файла
        date_str = get_date_from_filename(pn_filename)
        print(f"Извлеченная дата для ПН: {date_str}")
        
        # Получаем список файлов в архиве
        print("Сканируем архив ПН...")
        files_in_zip = get_available_files(pn_zip)
        print(f"Найдено {len(files_in_zip)} файлов в архиве")
        
        # Ищем файл для ПН
        target_file_pn = None
        # Пробуем найти по разным паттернам
        patterns = ['Расчет', 'стоимости', 'активов', 'НПФ']
        for f in files_in_zip:
            for pattern in patterns:
                if pattern in f:
                    target_file_pn = f
                    print(f"Найден файл для ПН по паттерну '{pattern}': {f}")
                    break
            if target_file_pn:
                break
        
        if target_file_pn:
            # Формируем имя целевого файла
            target_filename_pn = f"{date_str}_Расчет стоимости активов ПН с учетом портфеля НПФ.xls"
            dest_path_pn = os.path.join(path_i, target_filename_pn)
            
            # Извлекаем файл
            print(f"Извлечение файла для ПН в {dest_path_pn}")
            extract_and_rename(pn_zip, target_file_pn, dest_path_pn)
        else:
            print("Файл для ПН не найден в архиве")
            print("Доступные файлы:")
            for f in files_in_zip[:5]:  # Показываем первые 5 файлов
                print(f"  - {f}")
    else:
        print("Архив ПН не найден")
    
    print("Обработка завершена!")


if __name__ == "__main__":
    main()
