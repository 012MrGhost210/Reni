import os
import shutil
import re
from datetime import datetime
import logging

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('file_organizer.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

def extract_month_from_date(filename):
    """
    Извлекает номер месяца из имени файла с датой в различных форматах
    """
    try:
        # Ищем паттерны даты в имени файла
        # Формат с точками: 25.02.2026
        dot_pattern = r'(\d{2})\.(\d{2})\.(\d{4})'
        dot_match = re.search(dot_pattern, filename)
        if dot_match:
            month = int(dot_match.group(2))
            logging.info(f"Извлечен месяц {month} из даты {dot_match.group(0)}")
            return str(month)
        
        # Формат без точек: 03032026 (ДДММГГГГ)
        plain_pattern = r'(\d{2})(\d{2})(\d{4})'
        plain_match = re.search(plain_pattern, filename)
        if plain_match:
            # В формате ДДММГГГГ, месяц - это вторая группа из двух цифр
            day = plain_match.group(1)
            month = plain_match.group(2)
            year = plain_match.group(3)
            
            # Проверяем, что месяц корректен (01-12)
            month_int = int(month)
            if 1 <= month_int <= 12:
                logging.info(f"Извлечен месяц {month_int} из даты {day}{month}{year}")
                return str(month_int)
            else:
                logging.warning(f"Некорректный месяц {month_int} в дате {day}{month}{year}")
        
        # Если не нашли дату, пробуем найти любую последовательность цифр,
        # которая может содержать дату
        numbers = re.findall(r'\d+', filename)
        for num in numbers:
            if len(num) == 8:  # Формат ДДММГГГГ
                month_candidate = num[2:4]
                try:
                    month_int = int(month_candidate)
                    if 1 <= month_int <= 12:
                        logging.info(f"Извлечен месяц {month_int} из числа {num}")
                        return str(month_int)
                except ValueError:
                    continue
        
        # Если ничего не нашли, используем текущий месяц
        current_month = datetime.now().month
        logging.warning(f"Не удалось извлечь дату из '{filename}', используем текущий месяц: {current_month}")
        return str(current_month)
        
    except Exception as e:
        logging.error(f"Ошибка при извлечении месяца из {filename}: {str(e)}")
        current_month = datetime.now().month
        return str(current_month)

def clean_filename(filename):
    """
    Удаляет {Документооборот завершён} из начала имени файла
    """
    pattern = r'^{Документооборот завершён}'
    return re.sub(pattern, '', filename)

def process_files(source_root, destinations):
    """
    Основная функция обработки файлов
    """
    
    logging.info(f"Начинаем поиск в папке: {source_root}")
    
    for root, dirs, files in os.walk(source_root):
        # Проверяем каждый тип назначения
        for dest_type, dest_config in destinations.items():
            folder_pattern = dest_config['folder_pattern']
            folder_name = os.path.basename(root)
            
            # Проверяем, содержит ли текущая папка нужный паттерн
            if folder_pattern.lower() in folder_name.lower():
                logging.info(f"Найдена целевая папка: {root}")
                files_found = False
                
                # Ищем файлы в этой папке
                for file in files:
                    # Проверяем, подходит ли файл под критерии
                    file_lower = file.lower()
                    file_pattern = dest_config['file_pattern'].lower()
                    
                    if file_pattern in file_lower:
                        # Проверяем наличие даты в файле
                        has_date = (re.search(r'\d{2}\.\d{2}\.\d{4}', file) or 
                                   re.search(r'\d{8}', file))
                        
                        if has_date:
                            files_found = True
                            logging.info(f"Найден целевой файл: {file}")
                            
                            try:
                                # Извлекаем месяц из имени файла
                                month_num = extract_month_from_date(file)
                                
                                # Очищаем имя файла
                                clean_name = clean_filename(file)
                                
                                # Формируем путь назначения
                                dest_base = dest_config['dest_path']
                                month_folder = os.path.join(dest_base, month_num)
                                
                                # Создаем папку месяца, если её нет
                                os.makedirs(month_folder, exist_ok=True)
                                
                                # Полный путь к исходному и целевому файлу
                                src_file = os.path.join(root, file)
                                dest_file = os.path.join(month_folder, clean_name)
                                
                                # Проверяем, не существует ли уже файл
                                if os.path.exists(dest_file):
                                    base, ext = os.path.splitext(clean_name)
                                    counter = 1
                                    while os.path.exists(dest_file):
                                        new_name = f"{base}_{counter}{ext}"
                                        dest_file = os.path.join(month_folder, new_name)
                                        counter += 1
                                    logging.info(f"Файл уже существует, создаем копию: {new_name}")
                                
                                # Копируем файл
                                shutil.copy2(src_file, dest_file)
                                logging.info(f"Файл скопирован: {src_file} -> {dest_file}")
                                
                            except Exception as e:
                                logging.error(f"Ошибка при обработке файла {file}: {str(e)}")
                
                if not files_found:
                    logging.info(f"В папке {root} не найдено подходящих файлов")
    
    logging.info("Поиск завершен")

def main():
    # Настройка корневой папки для поиска
    # ЗАМЕНИТЕ НА ВАШ ПУТЬ
    source_root = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\diadoc_connector\Документооборот завершён"
    
    # Настройка папок назначения с параметрами
    destinations = {
        'raiffeisen': {
            'folder_pattern': 'УК Райффайзен',
            'file_pattern': 'сводный отчет',
            'dest_path': r'M:\Финансовый департамент\Treasury\3. ЗАКРЫТИЕ\Отчеты УК\Сводные отчеты УК\2026\Райф'  # Путь X
        },
        'sputnik': {
            'folder_pattern': 'СПУТНИК',
            'file_pattern': 'сводный',
            'dest_path': r'M:\Финансовый департамент\Treasury\3. ЗАКРЫТИЕ\Отчеты УК\Сводные отчеты УК\2026\Спутник'     # Путь Y
        },
        'tkb': {
            'folder_pattern': 'ТКБ',
            'file_pattern': 'сводный отчет',
            'dest_path': r'M:\Финансовый департамент\Treasury\3. ЗАКРЫТИЕ\Отчеты УК\Сводные отчеты УК\2026\ТКБ'         # Путь Z
        }
    }
    
    # Проверяем существование исходной папки
    if not os.path.exists(source_root):
        logging.error(f"Исходная папка не существует: {source_root}")
        return
    
    # Проверяем существование папок назначения
    for dest_type, dest_config in destinations.items():
        dest_path = dest_config['dest_path']
        if not os.path.exists(dest_path):
            try:
                os.makedirs(dest_path)
                logging.info(f"Создана папка назначения: {dest_path}")
            except Exception as e:
                logging.error(f"Не удалось создать папку {dest_path}: {str(e)}")
    
    # Запускаем обработку
    try:
        process_files(source_root, destinations)
        logging.info("Обработка завершена!")
    except Exception as e:
        logging.error(f"Критическая ошибка: {str(e)}")

if __name__ == "__main__":
    main()
