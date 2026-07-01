import os
import shutil
import re
from datetime import datetime
from pathlib import Path

# ==================== НАСТРОЙКИ (меняй здесь) ====================

# Базовый путь назначения
BASE_DEST_DIR = r"M:\Финансовый департамент\Treasury\Отчеты Брокера и Спецдепозитария"

# Что удаляем из названия при копировании
REMOVE_FROM_NAME = [
    "{Документооборот завершен}",
]

# ==================== КОНФИГУРАЦИЯ ИСТОЧНИКОВ ====================

# Каждый источник - это словарь с:
# - source_dir: путь к папке-источнику (управляющий извлекается из последней папки)
# - manager: можно указать вручную или оставить None для автоподстановки
# - configs: конфиги для этого источника

SOURCES_CONFIG = [
    {
        "source_dir": r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\Diadoc\diadoc_connector\Документооборот завершен\7710183778-АО УК -ПЕРВАЯ-",
        "manager": None,  # None = взять из названия последней папки
        "configs": {
            # Правило 1: Выгрузка операций
            "{Документооборот завершен}Выгрузка операций_": {
                "date_format": "%d%m%Y",
                "date_regex": r"(\d{8})",
                "destination_template": r"10.НАПФ - Ценные бумаги\*ГГГГ*\*Управляющий*\*Месяц*",
                "date_at_start": False,
                "skip_prefix": False,
            },
            # Правило 2: Журнал учета операций
            "{Документооборот завершен}I02_514833_k_d_": {
                "date_format": "%d%m%Y",
                "date_regex": r"(\d{6})",
                "destination_template": r"6.Журнал учета операций\*ГГГГ*\*Месяц*\*Управляющий*",
                "date_at_start": False,
                "skip_prefix": False,
            },
            # Правило 3: Отчеты брокера
            "{Документооборот завершен}": {
                "date_format": "%Y.%m.%d",
                "date_regex": r"(\d{4}\.\d{2}\.\d{2})",
                "destination_template": r"2 Отчеты брокера\*ГГГГ*\*Месяц*\*Управляющий*",
                "date_at_start": True,
                "skip_prefix": True,
            }
        }
    },
    # ===== ДОБАВЛЯЙ НОВЫХ УПРАВЛЯЮЩИХ СЮДА =====
    {
        "source_dir": r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\Diadoc\diadoc_connector\Документооборот завершен\7710183778-АО УК -ВТОРАЯ-",
        "manager": "УК Вторая",  # можно указать вручную
        "configs": {
            # Те же правила, что и для первого управляющего
            "{Документооборот завершен}Выгрузка операций_": {
                "date_format": "%d%m%Y",
                "date_regex": r"(\d{8})",
                "destination_template": r"10.НАПФ - Ценные бумаги\*ГГГГ*\*Управляющий*\*Месяц*",
                "date_at_start": False,
                "skip_prefix": False,
            },
            "{Документооборот завершен}I02_514833_k_d_": {
                "date_format": "%d%m%Y",
                "date_regex": r"(\d{6})",
                "destination_template": r"6.Журнал учета операций\*ГГГГ*\*Месяц*\*Управляющий*",
                "date_at_start": False,
                "skip_prefix": False,
            },
            "{Документооборот завершен}": {
                "date_format": "%Y.%m.%d",
                "date_regex": r"(\d{4}\.\d{2}\.\d{2})",
                "destination_template": r"2 Отчеты брокера\*ГГГГ*\*Месяц*\*Управляющий*",
                "date_at_start": True,
                "skip_prefix": True,
            }
        }
    }
]

# ==================== ОСНОВНАЯ ЛОГИКА ====================

# Расширения файлов, которые обрабатываем
ALLOWED_EXTENSIONS = {'.pdf', '.xlsx', '.xls', '.doc', '.docx', '.txt'}

def extract_manager_from_path(source_dir):
    """
    Извлекает название управляющего из последней папки пути
    """
    # Нормализуем путь
    normalized_path = source_dir.replace('\\', '/').rstrip('/')
    # Берем последнюю часть
    manager = os.path.basename(normalized_path)
    return manager

def clean_filename(filename):
    """Удаляет из имени файла все части из REMOVE_FROM_NAME"""
    new_name = filename
    for remove_part in REMOVE_FROM_NAME:
        new_name = new_name.replace(remove_part, "")
    return new_name

def extract_date_from_filename(filename, configs):
    """
    Извлекает дату из имени файла на основе конфигурации корней
    """
    for root, config in configs.items():
        if filename.startswith(root):
            # Получаем часть после корня
            if config.get("skip_prefix", False):
                name_part = filename
            else:
                name_part = filename[len(root):]
            
            # Ищем дату по регулярке
            match = re.search(config["date_regex"], name_part)
            if match:
                date_str = match.group(1)
                try:
                    date_obj = datetime.strptime(date_str, config["date_format"])
                    return date_obj, root, config
                except ValueError:
                    continue
    
    return None, None, None

def build_destination_path(date_obj, config, original_filename, manager_name):
    """
    Строит путь назначения на основе шаблона
    """
    year = str(date_obj.year)
    month = f"{date_obj.month:02d}"
    day = f"{date_obj.day:02d}"
    
    # Получаем шаблон пути
    template = config["destination_template"]
    
    # Заменяем переменные
    dest_path = template.replace("*ГГГГ*", year)
    dest_path = dest_path.replace("*Месяц*", month)
    dest_path = dest_path.replace("*День*", day)
    dest_path = dest_path.replace("*Управляющий*", manager_name)
    
    # Полный путь
    full_path = os.path.join(BASE_DEST_DIR, dest_path)
    
    # Создаем папки
    Path(full_path).mkdir(parents=True, exist_ok=True)
    
    # Очищаем имя файла
    clean_name = clean_filename(original_filename)
    if not clean_name or clean_name == os.path.splitext(original_filename)[1]:
        clean_name = original_filename
    
    return os.path.join(full_path, clean_name)

def process_source(source_config):
    """
    Обрабатывает один источник
    """
    source_dir = source_config["source_dir"]
    configs = source_config["configs"]
    
    # Определяем имя управляющего
    if source_config.get("manager"):
        manager_name = source_config["manager"]
    else:
        manager_name = extract_manager_from_path(source_dir)
    
    print(f"\n{'='*70}")
    print(f"📁 Управляющий: {manager_name}")
    print(f"📂 Источник: {source_dir}")
    print(f"{'='*70}")
    
    # Проверяем существование папки-источника
    if not os.path.exists(source_dir):
        print(f"❌ Папка-источник не найдена: {source_dir}")
        return 0, 0, 0, 0
    
    # Получаем список файлов
    files = [f for f in os.listdir(source_dir) if os.path.isfile(os.path.join(source_dir, f))]
    
    if not files:
        print("ℹ️ В папке-источнике нет файлов")
        return 0, 0, 0, 0
    
    processed = 0
    skipped = 0
    skipped_exists = 0
    skipped_no_date = 0
    
    for filename in files:
        file_path = os.path.join(source_dir, filename)
        
        # Проверяем расширение
        ext = os.path.splitext(filename)[1].lower()
        if ALLOWED_EXTENSIONS and ext not in ALLOWED_EXTENSIONS:
            print(f"⏭️ Пропускаем (неподдерживаемое расширение): {filename}")
            skipped += 1
            continue
        
        # Извлекаем дату и конфиг
        date_obj, root, config = extract_date_from_filename(filename, configs)
        
        if date_obj is None:
            print(f"⚠️ Не удалось извлечь дату из: {filename}")
            skipped_no_date += 1
            continue
        
        # Формируем путь назначения
        dest_path = build_destination_path(date_obj, config, filename, manager_name)
        
        # Проверяем, существует ли уже файл
        if os.path.exists(dest_path):
            print(f"⏭️ Файл уже существует, пропускаем: {os.path.basename(dest_path)}")
            skipped_exists += 1
            continue
        
        try:
            # Копируем файл
            shutil.copy2(file_path, dest_path)
            print(f"✅ {filename}")
            print(f"   -> {dest_path}")
            processed += 1
        except Exception as e:
            print(f"❌ Ошибка при копировании {filename}: {e}")
            skipped += 1
    
    print(f"\n📊 Итог по {manager_name}: обработано {processed}, пропущено {skipped_exists} (существуют), {skipped_no_date} (нет даты), {skipped} (ошибки)")
    
    return processed, skipped, skipped_exists, skipped_no_date

def process_all_sources():
    """Обрабатывает все источники"""
    total_processed = 0
    total_skipped = 0
    total_exists = 0
    total_no_date = 0
    
    for source_config in SOURCES_CONFIG:
        processed, skipped, skipped_exists, skipped_no_date = process_source(source_config)
        total_processed += processed
        total_skipped += skipped
        total_exists += skipped_exists
        total_no_date += skipped_no_date
    
    print(f"\n{'='*70}")
    print(f"📊 ОБЩИЙ ИТОГ ПО ВСЕМ УПРАВЛЯЮЩИМ:")
    print(f"   ✅ Обработано: {total_processed}")
    print(f"   ⏭️ Пропущено (уже существуют): {total_exists}")
    print(f"   ⚠️ Пропущено (нет даты): {total_no_date}")
    print(f"   ❌ Пропущено (ошибки/неподдерживаемые): {total_skipped}")
    print(f"{'='*70}")

# ==================== ЗАПУСК ====================

if __name__ == "__main__":
    process_all_sources()
