      ⏭️ Пропускаем (неподдерживаемое расширение): íα«¬Ñα_¡á»Σ_»α«τ¿Ñ_09072026_256/æó«ñ¡δ⌐ «ΓτÑΓ «í «ßΓáΓ¬áσ_1881-æä_ÆôD0052680_09072026.pdf
      ⚠️ Не удалось распознать файл: íα«¬Ñα_¡á»Σ_»α«τ¿Ñ_09072026_256/æó«ñ¡δ⌐ «ΓτÑΓ «í «ßΓáΓ¬áσ_1881-æä_ÆôD0052680_09072026.xml

📦 Обработка архива: {Документооборот завершен}Брокер_напф_прочие_09072026_257.zip
   📅 Дата архива: 09.07.2026
   📁 Портфель: 257
   📄 Найдено файлов в архиве: 3
      ⏭️ Пропускаем (неподдерживаемое расширение): íα«¬Ñα_¡á»Σ_»α«τ¿Ñ_09072026_257/
      ⚠️ Не удалось распознать файл: íα«¬Ñα_¡á»Σ_»α«τ¿Ñ_09072026_257/257  éδúαπº¬á ìÇÅö 20260709.xml
import os
import shutil
import re
import zipfile
import sys
from datetime import datetime
from pathlib import Path

# ==================== НАСТРОЙКА КОДИРОВКИ ====================

# Устанавливаем кодировку для консоли
if sys.platform == 'win32':
    import locale
    # Устанавливаем кодировку для Windows
    sys.stdout.reconfigure(encoding='utf-8') if hasattr(sys.stdout, 'reconfigure') else None
    
# Кодировка для работы с файловой системой
FS_ENCODING = 'utf-8' if sys.platform != 'win32' else 'cp1251'

# ==================== НАСТРОЙКИ ====================

# Базовый путь назначения
BASE_DEST_DIR = r"M:\Финансовый департамент\Treasury\Отчеты Брокера и Спецдепозитария"

# Папка-источник с архивами
SOURCE_DIR = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\Diadoc\diadoc_connector\Документооборот завершен\7702358512-ООО -УК Райффайзен-"

# Маппинг названий управляющих (исходное -> красивое)
MANAGER_MAPPING = {
    "7702358512-ООО -УК Райффайзен-": "Райффайзен Капитал",
    # Добавляй других управляющих сюда
}

# Маппинг брокеров
BROKER_MAPPING = {
    "82748": "Брокер Райффазенбанк",
    # Добавляй других брокеров сюда
}

# Разрешенные расширения файлов в архиве
ALLOWED_EXTENSIONS = {'.xml', '.txt', '.csv', '.xlsx', '.xls'}

# ==================== КОНФИГУРАЦИЯ ====================

# Правила для архивов
ARCHIVE_RULES = [
    {
        # Правило 1: Архив с Брокер_напф_прочие
        "archive_pattern": "{Документооборот завершен}Брокер_напф_прочие_",
        "archive_regex": r"Брокер_напф_прочие_(\d{8})_(\d+)",  # дата + код портфеля
        "archive_date_format": "%d%m%Y",
    }
]

# Правила для файлов внутри архива
FILE_RULES = [
    {
        # Правило 1: XML файл "Выгрузка НАПФ"
        "file_pattern": r"(\d+)\s+Выгрузка НАПФ\s+(\d{8})\.xml",
        "file_date_format": "%Y%m%d",
        "destination_template": r"10.НАПФ - Ценные бумаги\*ГГГГ*\*Управляющий*\*Портфель*\*Месяц*",
        "identifier": "Выгрузка НАПФ",
    },
    {
        # Правило 2: trades файл
        "file_pattern": r"(\d+)_trades_(\d{8})_\d+",
        "file_date_format": "%Y%m%d",
        "destination_template": r"2 Отчеты брокера\*ГГГГ*\*Месяц*\*Управляющий*\*Портфель*",
        "identifier": "trades",
    },
    {
        # Правило 3: brok_rpt файл
        "file_pattern": r"brok_rpt_(\d+)_(\d{8})_\d+_final",
        "file_date_format": "%Y%m%d",
        "destination_template": r"2 Отчеты брокера\*ГГГГ*\*Месяц*\*Управляющий*\*Портфель*",
        "identifier": "brok_rpt",
    },
]

# ==================== ОСНОВНАЯ ЛОГИКА ====================

def get_manager_name(source_dir):
    """Извлекает имя управляющего из пути"""
    # Работаем с путем в кодировке UTF-8
    normalized_path = source_dir.replace('\\', '/').rstrip('/')
    raw_name = os.path.basename(normalized_path)
    
    # Пытаемся декодировать, если это байты
    if isinstance(raw_name, bytes):
        try:
            raw_name = raw_name.decode('utf-8')
        except UnicodeDecodeError:
            try:
                raw_name = raw_name.decode('cp1251')
            except UnicodeDecodeError:
                raw_name = str(raw_name)
    
    return MANAGER_MAPPING.get(raw_name, raw_name)

def get_month_name(month_number):
    """Возвращает название месяца на русском"""
    months = {
        1: "Январь", 2: "Февраль", 3: "Март",
        4: "Апрель", 5: "Май", 6: "Июнь",
        7: "Июль", 8: "Август", 9: "Сентябрь",
        10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"
    }
    return months.get(month_number, "")

def format_month_folder(date_obj):
    """Форматирует папку месяца: 06.Июнь 2026"""
    month_num = f"{date_obj.month:02d}"
    month_name = get_month_name(date_obj.month)
    year = date_obj.year
    return f"{month_num}.{month_name} {year}"

def extract_archive_info(filename):
    """Извлекает дату и код портфеля из имени архива"""
    for rule in ARCHIVE_RULES:
        if not filename.startswith(rule["archive_pattern"]):
            continue
        
        # Ищем дату и код портфеля
        match = re.search(rule["archive_regex"], filename)
        if match:
            date_str = match.group(1)
            portfolio_code = match.group(2)
            try:
                date_obj = datetime.strptime(date_str, rule["archive_date_format"])
                return date_obj, portfolio_code
            except ValueError:
                continue
    
    return None, None

def extract_file_info(filename):
    """Извлекает информацию из имени файла внутри архива"""
    for rule in FILE_RULES:
        match = re.search(rule["file_pattern"], filename)
        if match:
            if rule["identifier"] in ["trades", "brok_rpt"]:
                broker_code = match.group(1)
                date_str = match.group(2)
                portfolio_code = None
            else:
                portfolio_code = match.group(1)
                date_str = match.group(2)
                broker_code = None
            
            try:
                date_obj = datetime.strptime(date_str, rule["file_date_format"])
                return date_obj, portfolio_code, broker_code, rule
            except ValueError:
                continue
    
    return None, None, None, None

def build_destination_path(date_obj, rule, manager_name, portfolio_code, broker_code, original_filename):
    """Строит путь назначения для файла"""
    year = str(date_obj.year)
    month_folder = format_month_folder(date_obj)
    
    # Получаем шаблон пути
    dest_path = rule["destination_template"]
    
    # Определяем портфель для пути
    portfolio_folder = portfolio_code if portfolio_code else broker_code
    
    # Заменяем переменные
    dest_path = dest_path.replace("*ГГГГ*", year)
    dest_path = dest_path.replace("*Месяц*", month_folder)
    dest_path = dest_path.replace("*Управляющий*", manager_name)
    dest_path = dest_path.replace("*Портфель*", portfolio_folder)
    
    # Полный путь
    full_path = os.path.join(BASE_DEST_DIR, dest_path)
    
    # Создаем папки
    Path(full_path).mkdir(parents=True, exist_ok=True)
    
    return os.path.join(full_path, original_filename)

def normalize_filename(filename):
    """Нормализует имя файла для сравнения"""
    name_without_ext, ext = os.path.splitext(filename)
    for char in ['+', ' ', '_', '-']:
        name_without_ext = name_without_ext.replace(char, '')
    return name_without_ext.lower() + ext

def check_file_exists(dest_path):
    """Проверяет, существует ли файл"""
    dest_dir = os.path.dirname(dest_path)
    target_filename = os.path.basename(dest_path)
    
    if not os.path.exists(dest_dir):
        return False
    
    target_normalized = normalize_filename(target_filename)
    
    try:
        existing_files = [f for f in os.listdir(dest_dir) if os.path.isfile(os.path.join(dest_dir, f))]
    except PermissionError:
        return False
    
    for existing_file in existing_files:
        if normalize_filename(existing_file) == target_normalized:
            return True
    
    return False

def process_archive(zip_path, manager_name):
    """Обрабатывает один архив"""
    archive_name = os.path.basename(zip_path)
    print(f"\n📦 Обработка архива: {archive_name}")
    
    # Извлекаем информацию из имени архива
    archive_date, portfolio_code = extract_archive_info(archive_name)
    if not archive_date:
        print(f"   ⚠️ Не удалось определить дату/портфель из имени архива")
        return 0, 0
    
    print(f"   📅 Дата архива: {archive_date.strftime('%d.%m.%Y')}")
    print(f"   📁 Портфель: {portfolio_code}")
    
    processed = 0
    skipped = 0
    
    try:
        # Открываем архив с правильной кодировкой
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            # Получаем список файлов в архиве
            files_in_zip = zip_ref.namelist()
            print(f"   📄 Найдено файлов в архиве: {len(files_in_zip)}")
            
            for filename in files_in_zip:
                # Проверяем расширение
                ext = os.path.splitext(filename)[1].lower()
                if ALLOWED_EXTENSIONS and ext not in ALLOWED_EXTENSIONS:
                    print(f"      ⏭️ Пропускаем (неподдерживаемое расширение): {filename}")
                    skipped += 1
                    continue
                
                # Извлекаем информацию из имени файла
                file_date, file_portfolio, broker_code, rule = extract_file_info(filename)
                
                if not file_date:
                    print(f"      ⚠️ Не удалось распознать файл: {filename}")
                    skipped += 1
                    continue
                
                # Определяем код портфеля (из архива или из файла)
                final_portfolio = portfolio_code if portfolio_code else file_portfolio
                
                if not final_portfolio:
                    print(f"      ⚠️ Не удалось определить портфель для: {filename}")
                    skipped += 1
                    continue
                
                # Для brok_rpt и trades используем broker_code как портфель
                if rule["identifier"] in ["trades", "brok_rpt"]:
                    final_portfolio = broker_code
                
                # Строим путь назначения
                dest_path = build_destination_path(
                    file_date, rule, manager_name, 
                    final_portfolio, broker_code, filename
                )
                
                # Проверяем, существует ли файл
                if check_file_exists(dest_path):
                    print(f"      ⏭️ Файл уже существует: {filename}")
                    skipped += 1
                    continue
                
                # Извлекаем файл из архива
                try:
                    # Читаем файл из архива
                    file_data = zip_ref.read(filename)
                    
                    # Создаем папки если их нет
                    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                    
                    # Записываем файл с правильной кодировкой
                    with open(dest_path, 'wb') as f:
                        f.write(file_data)
                    
                    print(f"      ✅ {filename}")
                    print(f"         -> {dest_path}")
                    processed += 1
                except Exception as e:
                    print(f"      ❌ Ошибка при извлечении {filename}: {e}")
                    skipped += 1
                    
    except zipfile.BadZipFile:
        print(f"   ❌ Файл не является zip-архивом: {zip_path}")
        return 0, 0
    except UnicodeDecodeError as e:
        print(f"   ❌ Ошибка кодировки при открытии архива: {e}")
        print(f"   🔄 Пробуем открыть с другой кодировкой...")
        # Пробуем открыть с другой кодировкой
        try:
            with zipfile.ZipFile(zip_path, 'r', metadata_encoding='cp866') as zip_ref:
                # ... аналогичная логика
                pass
        except Exception as e2:
            print(f"   ❌ Не удалось открыть архив: {e2}")
            return 0, 0
    except Exception as e:
        print(f"   ❌ Ошибка при открытии архива: {e}")
        return 0, 0
    
    return processed, skipped

def process_files():
    """Основная функция обработки"""
    print("=" * 70)
    print("📁 Начинаем обработку архивов")
    print(f"📂 Источник: {SOURCE_DIR}")
    print(f"📂 Назначение: {BASE_DEST_DIR}")
    print("=" * 70)
    
    # Проверяем существование папки-источника
    if not os.path.exists(SOURCE_DIR):
        print(f"❌ Папка-источник не найдена: {SOURCE_DIR}")
        return
    
    # Получаем имя управляющего
    manager_name = get_manager_name(SOURCE_DIR)
    print(f"👤 Управляющий: {manager_name}")
    print("=" * 70)
    
    # Получаем список архивов
    files = [f for f in os.listdir(SOURCE_DIR) if os.path.isfile(os.path.join(SOURCE_DIR, f))]
    
    # Фильтруем только zip архивы
    archives = [f for f in files if f.lower().endswith(('.zip', '.rar'))]
    
    if not archives:
        print("ℹ️ В папке-источнике нет архивов")
        return
    
    print(f"📄 Найдено архивов: {len(archives)}")
    print("=" * 70)
    
    total_processed = 0
    total_skipped = 0
    
    for archive_name in archives:
        archive_path = os.path.join(SOURCE_DIR, archive_name)
        processed, skipped = process_archive(archive_path, manager_name)
        total_processed += processed
        total_skipped += skipped
    
    # Итоги
    print("\n" + "=" * 70)
    print("📊 ИТОГ:")
    print(f"   ✅ Извлечено файлов: {total_processed}")
    print(f"   ⏭️ Пропущено: {total_skipped}")
    print("=" * 70)

# ==================== ЗАПУСК ====================

if __name__ == "__main__":
    # Устанавливаем правильную кодировку для вывода
    import sys
    import io
    if sys.platform == 'win32':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    
    process_files()

import os
import shutil
import re
from datetime import datetime
from pathlib import Path

# ==================== НАСТРОЙКИ ====================

# Базовый путь назначения
BASE_DEST_DIR = r"M:\Финансовый департамент\Treasury\Отчеты Брокера и Спецдепозитария\test"

# Папка-источник
SOURCE_DIR = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\Diadoc\diadoc_connector\Документооборот завершен\7710183778-АО УК -ПЕРВАЯ-"

# Маппинг названий управляющих (исходное -> красивое)
MANAGER_MAPPING = {
    "7710183778-АО УК -ПЕРВАЯ-": "УК Первая",
    # Добавляй других управляющих сюда:
    # "7710183778-АО УК -ВТОРАЯ-": "УК Вторая",
    # "7710183778-АО УК -ТРЕТЬЯ-": "УК Третья",
}

# Разрешенные расширения файлов
ALLOWED_EXTENSIONS = {'.pdf', '.xlsx', '.xls', '.doc', '.docx', '.txt', '.xml'}

# Что удаляем из названия при копировании (только фигурные скобки с содержимым)
REMOVE_FROM_NAME = [
    "{Документооборот завершен}",
]

# ==================== КОНФИГУРАЦИЯ ФАЙЛОВ ====================

# ВАЖНО: правила проверяются ПО ПОРЯДКУ!
# Сначала самые специфичные, потом общие
FILE_RULES = [
    {
        # Правило 1: Выгрузка операций
        "name_pattern": "{Документооборот завершен}Выгрузка операций_",
        "date_regex": r"(\d{8})",  # 08072026 -> 8 цифр
        "date_format": "%d%m%Y",    # ДДММГГГГ
        "destination": r"10.НАПФ - Ценные бумаги\*ГГГГ*\*Управляющий*\*Месяц*",
        "identifier": "Выгрузка операций",  # для отладки
    },
    {
        # Правило 2: I02 (Журнал учета операций)
        "name_pattern": "{Документооборот завершен}I02_514833_k_d_",
        "date_regex": r"(\d{6})",   # 260706 -> 6 цифр
        "date_format": "%y%m%d",    # ГГММДД
        "destination": r"2 Отчеты брокера\*ГГГГ*\*Месяц*\*Управляющий*",
        "identifier": "I02_514833_k_d_",
    },
    {
        # Правило 3: Отчеты брокера (журнал учета ДС)
        "name_pattern": "{Документооборот завершен}",
        "date_regex": r"(\d{4}\.\d{2}\.\d{2})_27011_журнал учета ДС",  # ищем дату + _27011_журнал учета ДС
        "date_format": "%Y.%m.%d",   # ГГГГ.ММ.ДД
        "destination": r"6.Журнал учета операций\*ГГГГ*\*Месяц*\*Управляющий*",
        "identifier": "журнал учета ДС",
    },
]

# ==================== ОСНОВНАЯ ЛОГИКА ====================

def get_manager_name(source_dir):
    """Извлекает имя управляющего из пути и преобразует по маппингу"""
    normalized_path = source_dir.replace('\\', '/').rstrip('/')
    raw_name = os.path.basename(normalized_path)
    return MANAGER_MAPPING.get(raw_name, raw_name)

def get_month_name(month_number):
    """Возвращает название месяца на русском"""
    months = {
        1: "Январь", 2: "Февраль", 3: "Март",
        4: "Апрель", 5: "Май", 6: "Июнь",
        7: "Июль", 8: "Август", 9: "Сентябрь",
        10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"
    }
    return months.get(month_number, "")

def format_month_folder(date_obj):
    """Форматирует папку месяца: 06.Июнь или 06.Июнь 2026"""
    month_num = f"{date_obj.month:02d}"
    month_name = get_month_name(date_obj.month)
    year = date_obj.year
    return f"{month_num}.{month_name} {year}"

def clean_filename(filename):
    """Удаляет из имени файла все части из REMOVE_FROM_NAME"""
    new_name = filename
    for remove_part in REMOVE_FROM_NAME:
        new_name = new_name.replace(remove_part, "")
    return new_name

def extract_date_from_filename(filename, rule):
    """Извлекает дату из имени файла по правилу"""
    # Проверяем, начинается ли файл с нужного паттерна
    if not filename.startswith(rule["name_pattern"]):
        return None
    
    # Убираем префикс
    name_part = filename[len(rule["name_pattern"]):]
    
    # Ищем дату по регулярке
    match = re.search(rule["date_regex"], name_part)
    if not match:
        return None
    
    date_str = match.group(1)
    
    try:
        date_obj = datetime.strptime(date_str, rule["date_format"])
        return date_obj
    except ValueError:
        return None

def build_destination_path(date_obj, rule, manager_name, original_filename):
    """Строит путь назначения"""
    year = str(date_obj.year)
    month_folder = format_month_folder(date_obj)
    
    # Получаем шаблон пути
    dest_path = rule["destination"]
    
    # Заменяем переменные
    dest_path = dest_path.replace("*ГГГГ*", year)
    dest_path = dest_path.replace("*Месяц*", month_folder)
    dest_path = dest_path.replace("*Управляющий*", manager_name)
    
    # Полный путь
    full_path = os.path.join(BASE_DEST_DIR, dest_path)
    
    # Создаем папки
    Path(full_path).mkdir(parents=True, exist_ok=True)
    
    # Очищаем имя файла (удаляем только {Документооборот завершен})
    clean_name = clean_filename(original_filename)
    
    return os.path.join(full_path, clean_name)

def normalize_filename(filename):
    """Нормализует имя файла для сравнения (удаляет +, пробелы и т.д.)"""
    name_without_ext, ext = os.path.splitext(filename)
    for char in ['+', ' ', '_', '-']:
        name_without_ext = name_without_ext.replace(char, '')
    return name_without_ext.lower() + ext

def check_file_exists(dest_path):
    """Проверяет, существует ли файл (с учетом нормализации)"""
    dest_dir = os.path.dirname(dest_path)
    target_filename = os.path.basename(dest_path)
    
    if not os.path.exists(dest_dir):
        return False
    
    target_normalized = normalize_filename(target_filename)
    
    try:
        existing_files = [f for f in os.listdir(dest_dir) if os.path.isfile(os.path.join(dest_dir, f))]
    except PermissionError:
        return False
    
    for existing_file in existing_files:
        if normalize_filename(existing_file) == target_normalized:
            return True
    
    return False

def process_files():
    """Основная функция обработки"""
    print("=" * 70)
    print("📁 Начинаем обработку файлов")
    print(f"📂 Источник: {SOURCE_DIR}")
    print(f"📂 Назначение: {BASE_DEST_DIR}")
    print("=" * 70)
    
    # Проверяем существование папки-источника
    if not os.path.exists(SOURCE_DIR):
        print(f"❌ Папка-источник не найдена: {SOURCE_DIR}")
        return
    
    # Получаем имя управляющего
    manager_name = get_manager_name(SOURCE_DIR)
    print(f"👤 Управляющий: {manager_name}")
    print("=" * 70)
    
    # Получаем список файлов
    files = [f for f in os.listdir(SOURCE_DIR) if os.path.isfile(os.path.join(SOURCE_DIR, f))]
    
    if not files:
        print("ℹ️ В папке-источнике нет файлов")
        return
    
    print(f"📄 Найдено файлов: {len(files)}")
    print("=" * 70)
    
    processed = 0
    skipped = 0
    skipped_exists = 0
    skipped_no_rule = 0
    skipped_no_date = 0
    
    for filename in files:
        file_path = os.path.join(SOURCE_DIR, filename)
        
        # Проверяем расширение
        ext = os.path.splitext(filename)[1].lower()
        if ALLOWED_EXTENSIONS and ext not in ALLOWED_EXTENSIONS:
            print(f"⏭️ Пропускаем (неподдерживаемое расширение): {filename}")
            skipped += 1
            continue
        
        # Ищем подходящее правило (ПО ПОРЯДКУ!)
        found_rule = None
        date_obj = None
        
        for rule in FILE_RULES:
            date_obj = extract_date_from_filename(filename, rule)
            if date_obj:
                found_rule = rule
                break
        
        if not found_rule:
            print(f"⚠️ Не найдено правило для: {filename}")
            skipped_no_rule += 1
            continue
        
        if not date_obj:
            print(f"⚠️ Не удалось извлечь дату из: {filename}")
            skipped_no_date += 1
            continue
        
        # Для отладки - показываем какое правило сработало
        print(f"🔍 Правило: {found_rule.get('identifier', 'unknown')} -> {filename}")
        
        # Строим путь назначения
        dest_path = build_destination_path(date_obj, found_rule, manager_name, filename)
        
        # Проверяем, существует ли файл
        if check_file_exists(dest_path):
            print(f"⏭️ Файл уже существует: {os.path.basename(dest_path)}")
            skipped_exists += 1
            continue
        
        # Копируем файл
        try:
            shutil.copy2(file_path, dest_path)
            print(f"✅ {filename}")
            print(f"   -> {dest_path}")
            processed += 1
        except Exception as e:
            print(f"❌ Ошибка при копировании {filename}: {e}")
            skipped += 1
    
    # Итоги
    print("\n" + "=" * 70)
    print("📊 ИТОГ:")
    print(f"   ✅ Обработано: {processed}")
    print(f"   ⏭️ Пропущено (уже существуют): {skipped_exists}")
    print(f"   ⚠️ Пропущено (нет правила): {skipped_no_rule}")
    print(f"   ⚠️ Пропущено (нет даты): {skipped_no_date}")
    print(f"   ❌ Пропущено (ошибки/неподдерживаемые): {skipped}")
    print("=" * 70)

# ==================== ЗАПУСК ====================

if __name__ == "__main__":
    process_files()
