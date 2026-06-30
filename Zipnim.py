import os
import re
import shutil
import zipfile
from datetime import datetime
from typing import Optional, Dict


def decode_filename(filename: str) -> str:
    """
    Пытается расшифровать имя файла в разных кодировках
    """
    # Список кодировок для проб
    encodings = ['cp866', 'cp1251', 'koi8-r', 'latin-1', 'windows-1251', 'utf-8']
    
    # Сначала пробуем просто вернуть оригинал
    print(f"  Оригинал: {filename}")
    
    # Пробуем каждую кодировку
    for encoding in encodings:
        try:
            decoded = filename.encode('latin-1').decode(encoding)
            # Проверяем, есть ли русские буквы в результате
            if re.search(r'[А-Яа-я]', decoded):
                print(f"  Расшифровано ({encoding}): {decoded}")
                return decoded
        except:
            continue
    
    # Если ничего не помогло, пробуем принудительно cp1251
    try:
        decoded = filename.encode('latin-1').decode('cp1251')
        return decoded
    except:
        pass
    
    print(f"  Не удалось расшифровать, возвращаем оригинал")
    return filename


def find_file_in_zip(zip_ref: zipfile.ZipFile, patterns: list) -> Optional[str]:
    """
    Ищет файл по паттернам с расшифровкой имен
    """
    for filename in zip_ref.namelist():
        # Пропускаем не Excel файлы
        if not filename.lower().endswith(('.xls', '.xlsx')):
            continue
        
        # Расшифровываем имя
        decoded_name = decode_filename(filename)
        
        # Проверяем каждый паттерн
        for pattern in patterns:
            # В оригинале
            if pattern in filename:
                print(f"  Найден по паттерну '{pattern}' в оригинале!")
                return filename
            
            # В расшифрованном
            if pattern in decoded_name:
                print(f"  Найден по паттерну '{pattern}' в расшифрованном имени!")
                return filename
            
            # Без учета регистра
            if pattern.lower() in decoded_name.lower():
                print(f"  Найден по паттерну '{pattern}' без учета регистра!")
                return filename
    
    print(f"  Файл с паттернами {patterns} не найден")
    return None


def extract_file_from_zip(zip_path: str, patterns: list, destination: str) -> bool:
    """
    Извлекает файл из архива
    """
    try:
        if not os.path.exists(zip_path):
            print(f"Архив не найден: {zip_path}")
            return False
        
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            print(f"\nПоиск по паттернам: {patterns}")
            print("Файлы в архиве:")
            
            # Показываем все файлы для диагностики
            for name in zip_ref.namelist():
                if name.lower().endswith(('.xls', '.xlsx')):
                    decoded = decode_filename(name)
                    print(f"  {name} -> {decoded}")
            
            # Ищем файл
            target_file = find_file_in_zip(zip_ref, patterns)
            
            if not target_file:
                print(f"Файл не найден")
                return False
            
            # Создаем директорию и извлекаем
            dest_dir = os.path.dirname(destination)
            os.makedirs(dest_dir, exist_ok=True)
            
            extracted_path = zip_ref.extract(target_file, dest_dir)
            shutil.move(extracted_path, destination)
            print(f"Файл сохранен: {destination}")
            return True
            
    except Exception as e:
        print(f"Ошибка: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    # Пути
    path_x = r"Q:\Финансовый отдел\01.Перечень имущества Фонда (СД)"
    path_y = r"\\fs-01.renlife.com\alldocs\Финансовый департамент\Treasury\18. НПФ\1. Отчеты\1.1 Ежедневные отчеты\СПУТНИК\Акутальные"
    path_z = r"\\fs-01.renlife.com\alldocs\Финансовый департамент\Treasury\18. НПФ\1. Отчеты\1.1 Ежедневные отчеты\ФОНД\Актуальные данные"
    path_i = r"\\fs-01.renlife.com\alldocs\Финансовый департамент\Treasury\18. НПФ\1. Отчеты\1.1 Ежедневные отчеты\ВИМ"
    
    print("Начинаем обработку...")
    print("=" * 60)
    
    # Шаг 1: Папка с датой
    today_str = datetime.now().strftime("%Y_%m_%d")
    today_folder = os.path.join(path_x, today_str)
    
    if not os.path.exists(today_folder):
        print(f"Папка не найдена: {today_folder}")
        return
    
    print(f"Найдена папка: {today_folder}")
    
    # Шаг 2: Папка с документами
    docs_folder = os.path.join(today_folder, "Документы от Гаранта СД НТД")
    
    if not os.path.exists(docs_folder):
        print(f"Папка не найдена: {docs_folder}")
        return
    
    print(f"Найдена папка: {docs_folder}")
    
    # Шаг 3: Архивы
    zip_files = {}
    for file in os.listdir(docs_folder):
        if file.lower().endswith('.zip'):
            if 'ПР' in file:
                zip_files['ПР'] = os.path.join(docs_folder, file)
            elif 'ПН' in file:
                zip_files['ПН'] = os.path.join(docs_folder, file)
    
    if not zip_files:
        print("Архивы не найдены")
        return
    
    print(f"Найдены архивы: {list(zip_files.keys())}")
    print("=" * 60)
    
    # Шаг 4: ПР архив
    if 'ПР' in zip_files:
        pr_zip = zip_files['ПР']
        print(f"\nОбработка ПР: {os.path.basename(pr_zip)}")
        
        # Извлекаем дату
        date_match = re.search(r'(\d{4})-(\d{2})-(\d{2})', pr_zip)
        if date_match:
            date_str = f"{date_match.group(3)}.{date_match.group(2)}.{date_match.group(1)}"
        else:
            date_str = datetime.now().strftime("%d.%m.%Y")
        
        print(f"Дата: {date_str}")
        
        # Файл 301024
        dest_path = os.path.join(path_y, f"{date_str}_СЧА УК СПУТНИК - УПРАВЛЕНИЕ КАПИТАЛОМ (Д.У. 301024_1).xls")
        extract_file_from_zip(pr_zip, ['301024', 'СЧА', 'УПРАВЛЕНИЕ'], dest_path)
        
        # Файл 080825
        dest_path = os.path.join(path_z, f"{date_str}_СЧА УК СПУТНИК - УПРАВЛЕНИЕ КАПИТАЛОМ (Д.У. 080825_1).xls")
        extract_file_from_zip(pr_zip, ['080825', 'СЧА', 'УПРАВЛЕНИЕ'], dest_path)
    
    # Шаг 5: ПН архив
    if 'ПН' in zip_files:
        pn_zip = zip_files['ПН']
        print(f"\nОбработка ПН: {os.path.basename(pn_zip)}")
        
        # Извлекаем дату
        date_match = re.search(r'(\d{4})-(\d{2})-(\d{2})', pn_zip)
        if date_match:
            date_str = f"{date_match.group(3)}.{date_match.group(2)}.{date_match.group(1)}"
        else:
            date_str = datetime.now().strftime("%d.%m.%Y")
        
        print(f"Дата: {date_str}")
        
        # Файл ПН
        dest_path = os.path.join(path_i, f"{date_str}_Расчет стоимости активов ПН с учетом портфеля НПФ.xls")
        extract_file_from_zip(pn_zip, ['Расчет', 'стоимости', 'активов', 'НПФ'], dest_path)
    
    print("\n" + "=" * 60)
    print("Обработка завершена!")


if __name__ == "__main__":
    main()
