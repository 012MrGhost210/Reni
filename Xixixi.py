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

def extract_month_from_date(date_str):
    """
    Извлекает номер месяца из даты в различных форматах
    """
    # Пробуем разные форматы даты
    date_formats = [
        r'(\d{2})\.(\d{2})\.(\d{4})',  # 25.02.2026
        r'(\d{2})(\d{2})(\d{4})'       # 03032026 (день, месяц, год)
    ]
    
    for date_format in date_formats:
        match = re.search(date_format, date_str)
        if match:
            if '.' in date_str:
                # Формат с точками: группа 2 - месяц
                month = int(match.group(2))
            else:
                # Формат без точек: группа 2 - месяц (03 из 03032026)
                month = int(match.group(1)[2:4]) if len(match.group(1)) == 2 else int(match.group(2))
            return str(month)
    
    # Если дата не найдена, используем текущий месяц
    current_month = datetime.now().month
    logging.warning(f"Не удалось извлечь дату из '{date_str}', используем текущий месяц: {current_month}")
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
    
    Параметры:
    source_root: корневая папка для поиска
    destinations: словарь с настройками для каждого типа папок
    """
    
    logging.info(f"Начинаем поиск в папке: {source_root}")
    
    for root, dirs, files in os.walk(source_root):
        # Проверяем каждый тип назначения
        for dest_type, dest_config in destinations.items():
            folder_pattern = dest_config['folder_pattern']
            
            # Проверяем, содержит ли текущая папка нужный паттерн
            if folder_pattern.lower() in os.path.basename(root).lower():
                logging.info(f"Найдена целевая папка: {root}")
                
                # Ищем файлы в этой папке
                for file in files:
                    # Проверяем, подходит ли файл под критерии
                    file_lower = file.lower()
                    if (dest_config['file_pattern'].lower() in file_lower and 
                        re.search(r'\d{2,}[._]?\d{2,}[._]?\d{4,}', file)):
                        
                        logging.info(f"Найден целевой файл: {file}")
                        
                        try:
                            # Извлекаем дату из имени файла
                            date_match = re.search(r'(\d{2}[._]?\d{2}[._]?\d{4})', file)
                            if date_match:
                                date_str = date_match.group(1)
                                month_num = extract_month_from_date(date_str)
                                
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
                                
                                # Копируем файл
                                shutil.copy2(src_file, dest_file)
                                logging.info(f"Файл скопирован: {src_file} -> {dest_file}")
                            else:
                                logging.warning(f"Не удалось извлечь дату из файла: {file}")
                                
                        except Exception as e:
                            logging.error(f"Ошибка при обработке файла {file}: {str(e)}")

def main():
    # Настройка корневой папки для поиска
    # Замените на ваш путь
    source_root = r"C:\путь\к\исходной\папке"
    
    # Настройка папок назначения с параметрами
    destinations = {
        'raiffeisen': {
            'folder_pattern': 'УК Райффайзен',
            'file_pattern': 'сводный отчет',
            'dest_path': r'X:\путь\к\папке\Raiffeisen'  # Путь X
        },
        'sputnik': {
            'folder_pattern': 'СПУТНИК',
            'file_pattern': 'сводный',
            'dest_path': r'Y:\путь\к\папке\Sputnik'     # Путь Y
        },
        'tkb': {
            'folder_pattern': 'ТКБ',
            'file_pattern': 'сводный отчет',
            'dest_path': r'Z:\путь\к\папке\TKB'         # Путь Z
        }
    }
    
    # Проверяем существование исходной папки
    if not os.path.exists(source_root):
        logging.error(f"Исходная папка не существует: {source_root}")
        return
    
    # Запускаем обработку
    try:
        process_files(source_root, destinations)
        logging.info("Обработка завершена!")
    except Exception as e:
        logging.error(f"Критическая ошибка: {str(e)}")

if __name__ == "__main__":
    main()
