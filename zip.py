import os
import zipfile
import re
from pathlib import Path

def find_and_extract_file(root_folder, target_filename, date_pattern=r'\d{4}-\d{2}-\d{2}'):
    """
    Ищет ZIP архивы в root_folder, находит в них файл,
    где в начале дата, а потом target_filename
    
    Args:
        root_folder: корневая папка для поиска
        target_filename: искомое название файла (без даты)
        date_pattern: регулярное выражение для даты
    """
    
    # Проходим по всем папкам и файлам
    for foldername, subfolders, filenames in os.walk(root_folder):
        for filename in filenames:
            # Ищем ZIP архивы
            if filename.endswith('.zip'):
                zip_path = os.path.join(foldername, filename)
                
                try:
                    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                        # Получаем список файлов в архиве
                        for file_in_zip in zip_ref.namelist():
                            # Проверяем, не папка ли это
                            if file_in_zip.endswith('/'):
                                continue
                                
                            # Извлекаем имя файла без пути
                            base_name = os.path.basename(file_in_zip)
                            
                            # Проверяем паттерн: дата + искомое название
                            pattern = f"^{date_pattern}_{target_filename}$"
                            if re.match(pattern, base_name, re.IGNORECASE):
                                print(f"Найден файл: {base_name} в архиве {zip_path}")
                                
                                # Создаем папку для распаковки
                                extract_folder = os.path.join(foldername, 'extracted')
                                os.makedirs(extract_folder, exist_ok=True)
                                
                                # Распаковываем конкретный файл
                                zip_ref.extract(file_in_zip, extract_folder)
                                print(f"Файл распакован в: {extract_folder}")
                                
                except Exception as e:
                    print(f"Ошибка при обработке {zip_path}: {e}")

# Пример использования
if __name__ == "__main__":
    root_folder = "/path/to/your/folder"  # Замените на свой путь
    target = "отчет.xlsx"  # Искомое название файла
    
    find_and_extract_file(root_folder, target)
