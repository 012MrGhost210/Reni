import os
import shutil
import zipfile
from datetime import datetime
import re


def extract_zip_with_proper_names(zip_path: str, extract_to: str) -> bool:
    """
    Разархивирует файлы в указанную папку с правильными именами
    """
    try:
        os.makedirs(extract_to, exist_ok=True)
        
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            # Извлекаем все файлы
            zip_ref.extractall(extract_to)
            print(f"✓ Архив распакован: {os.path.basename(zip_path)} -> {extract_to}")
            return True
    except Exception as e:
        print(f"✗ Ошибка при распаковке {zip_path}: {e}")
        return False


def find_and_copy_files(source_folder: str, patterns: list, destination: str, new_name: str) -> bool:
    """
    Ищет файлы по паттернам и копирует их с новым именем
    """
    try:
        files = os.listdir(source_folder)
        
        # Ищем файлы, которые подходят под паттерны
        found_files = []
        for file in files:
            if not file.lower().endswith(('.xls', '.xlsx')):
                continue
            
            # Проверяем каждый паттерн
            for pattern in patterns:
                if pattern in file:
                    found_files.append(file)
                    print(f"  Найден файл: {file} (паттерн: {pattern})")
                    break
        
        if not found_files:
            print(f"  Файлы с паттернами {patterns} не найдены в {source_folder}")
            return False
        
        # Берем первый найденный файл
        source_file = os.path.join(source_folder, found_files[0])
        
        # Создаем папку назначения
        dest_dir = os.path.dirname(destination)
        os.makedirs(dest_dir, exist_ok=True)
        
        # Копируем файл с новым именем
        shutil.copy2(source_file, destination)
        print(f"  ✓ Файл скопирован: {destination}")
        return True
        
    except Exception as e:
        print(f"  ✗ Ошибка при копировании: {e}")
        return False


def cleanup_temp_folder(folder_path: str):
    """
    Удаляет временную папку со всем содержимым
    """
    try:
        if os.path.exists(folder_path):
            shutil.rmtree(folder_path)
            print(f"✓ Временная папка удалена: {folder_path}")
    except Exception as e:
        print(f"✗ Ошибка при удалении {folder_path}: {e}")


def main():
    # Пути
    path_x = r"Q:\Финансовый отдел\01.Перечень имущества Фонда (СД)"
    path_y = r"\\fs-01.renlife.com\alldocs\Финансовый департамент\Treasury\18. НПФ\1. Отчеты\1.1 Ежедневные отчеты\СПУТНИК\Акутальные"
    path_z = r"\\fs-01.renlife.com\alldocs\Финансовый департамент\Treasury\18. НПФ\1. Отчеты\1.1 Ежедневные отчеты\ФОНД\Актуальные данные"
    path_i = r"\\fs-01.renlife.com\alldocs\Финансовый департамент\Treasury\18. НПФ\1. Отчеты\1.1 Ежедневные отчеты\ВИМ"
    
    # Временная папка для распаковки
    temp_base = r"C:\Temp\GarantExtract"
    
    print("="*80)
    print("НАЧАЛО ОБРАБОТКИ")
    print("="*80)
    
    # Шаг 1: Находим папку с сегодняшней датой
    today_str = datetime.now().strftime("%Y_%m_%d")
    today_folder = os.path.join(path_x, today_str)
    
    if not os.path.exists(today_folder):
        print(f"✗ Папка не найдена: {today_folder}")
        return
    
    print(f"✓ Найдена папка: {today_folder}")
    
    # Шаг 2: Находим папку с документами
    docs_folder = os.path.join(today_folder, "Документы от Гаранта СД НТД")
    
    if not os.path.exists(docs_folder):
        print(f"✗ Папка не найдена: {docs_folder}")
        return
    
    print(f"✓ Найдена папка: {docs_folder}")
    
    # Шаг 3: Находим архивы
    zip_files = {}
    for file in os.listdir(docs_folder):
        if file.lower().endswith('.zip'):
            if 'ПР' in file:
                zip_files['ПР'] = os.path.join(docs_folder, file)
            elif 'ПН' in file:
                zip_files['ПН'] = os.path.join(docs_folder, file)
    
    if not zip_files:
        print("✗ Архивы не найдены")
        return
    
    print(f"✓ Найдены архивы: {', '.join(zip_files.keys())}")
    print("="*80)
    
    # Создаем временную папку для этого запуска
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    temp_folder = os.path.join(temp_base, f"extract_{timestamp}")
    
    try:
        # Шаг 4: Обрабатываем ПР архив
        if 'ПР' in zip_files:
            print("\nОБРАБОТКА АРХИВА ПР")
            print("-"*80)
            
            pr_zip = zip_files['ПР']
            pr_temp = os.path.join(temp_folder, "ПР")
            
            # Извлекаем дату из имени архива
            date_match = re.search(r'(\d{4})-(\d{2})-(\d{2})', pr_zip)
            if date_match:
                date_str = f"{date_match.group(3)}.{date_match.group(2)}.{date_match.group(1)}"
            else:
                date_str = datetime.now().strftime("%d.%m.%Y")
            
            print(f"Дата: {date_str}")
            
            # Разархивируем
            if extract_zip_with_proper_names(pr_zip, pr_temp):
                # Копируем файл для 301024
                dest_file = f"{date_str}_СЧА УК СПУТНИК - УПРАВЛЕНИЕ КАПИТАЛОМ (Д.У. 301024_1).xls"
                dest_path = os.path.join(path_y, dest_file)
                print(f"\nПоиск файла для 301024...")
                find_and_copy_files(pr_temp, ['301024', 'СЧА'], dest_path, dest_file)
                
                # Копируем файл для 080825
                dest_file = f"{date_str}_СЧА УК СПУТНИК - УПРАВЛЕНИЕ КАПИТАЛОМ (Д.У. 080825_1).xls"
                dest_path = os.path.join(path_z, dest_file)
                print(f"\nПоиск файла для 080825...")
                find_and_copy_files(pr_temp, ['080825', 'СЧА'], dest_path, dest_file)
        
        # Шаг 5: Обрабатываем ПН архив
        if 'ПН' in zip_files:
            print("\nОБРАБОТКА АРХИВА ПН")
            print("-"*80)
            
            pn_zip = zip_files['ПН']
            pn_temp = os.path.join(temp_folder, "ПН")
            
            # Извлекаем дату
            date_match = re.search(r'(\d{4})-(\d{2})-(\d{2})', pn_zip)
            if date_match:
                date_str = f"{date_match.group(3)}.{date_match.group(2)}.{date_match.group(1)}"
            else:
                date_str = datetime.now().strftime("%d.%m.%Y")
            
            print(f"Дата: {date_str}")
            
            # Разархивируем
            if extract_zip_with_proper_names(pn_zip, pn_temp):
                # Копируем файл ПН
                dest_file = f"{date_str}_Расчет стоимости активов ПН с учетом портфеля НПФ.xls"
                dest_path = os.path.join(path_i, dest_file)
                print(f"\nПоиск файла для ПН...")
                find_and_copy_files(pn_temp, ['Расчет', 'стоимости', 'активов', 'НПФ'], dest_path, dest_file)
        
        print("\n" + "="*80)
        print("ОБРАБОТКА ЗАВЕРШЕНА")
        print("="*80)
        
    finally:
        # Чистим временную папку
        print(f"\nОчистка временных файлов...")
        cleanup_temp_folder(temp_folder)


if __name__ == "__main__":
    main()
