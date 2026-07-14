import os
import shutil
import re
import zipfile
import sys
import io
from datetime import datetime
from pathlib import Path

# ==================== НАСТРОЙКА КОДИРОВКИ ====================

# Устанавливаем кодировку для консоли
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# ==================== НАСТРОЙКИ ====================

# Базовый путь назначения
BASE_DEST_DIR = r"M:\Финансовый департамент\Treasury\Отчеты Брокера и Спецдепозитария\test2"

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
ALLOWED_EXTENSIONS = {'.xml', '.txt', '.csv', '.xlsx', '.xls', '.pdf'}

# Кодировка имен файлов внутри zip-архивов (DOS-архиваторы под Windows пишут cp866)
ZIP_METADATA_ENCODING = 'cp866'

# ==================== КОНФИГУРАЦИЯ ====================

# Правила для архивов
ARCHIVE_RULES = [
    {
        "archive_pattern": "{Документооборот завершен}Брокер_напф_прочие_",
        "archive_regex": r"Брокер_напф_прочие_(\d{8})_(\d+)",
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

def decode_zip_filename(name, flag_bits=0):
    """
    Восстанавливает имя файла из архива.

    Модуль zipfile по спецификации ZIP декодирует имена как cp437,
    если не установлен UTF-8 флаг. Русские архиваторы под Windows
    обычно пишут имена в cp866 — поэтому перекодируем cp437 -> cp866.
    """
    # Если установлен UTF-8 флаг (бит 11) — имя уже корректное
    if flag_bits & 0x800:
        return name

    try:
        # Возвращаем исходные байты (cp437) и декодируем правильно (cp866)
        return name.encode('cp437').decode(ZIP_METADATA_ENCODING)
    except (UnicodeEncodeError, UnicodeDecodeError):
        # Имя не было в cp437/cp866 — оставляем как есть
        return name

def get_manager_name(source_dir):
    """Извлекает имя управляющего из пути"""
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
    """
    Извлекает информацию из имени файла внутри архива.
    Ожидает уже декодированное имя БЕЗ пути (basename).
    """
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

def build_destination_path(date_obj, rule, manager_name, portfolio_code, broker_code, filename):
    """Строит путь назначения для файла"""
    year = str(date_obj.year)
    month_folder = format_month_folder(date_obj)

    dest_path = rule["destination_template"]

    portfolio_folder = portfolio_code if portfolio_code else broker_code

    dest_path = dest_path.replace("*ГГГГ*", year)
    dest_path = dest_path.replace("*Месяц*", month_folder)
    dest_path = dest_path.replace("*Управляющий*", manager_name)
    dest_path = dest_path.replace("*Портфель*", portfolio_folder)

    full_path = os.path.join(BASE_DEST_DIR, dest_path)
    Path(full_path).mkdir(parents=True, exist_ok=True)

    return os.path.join(full_path, filename)

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

def open_zip(zip_path):
    """
    Открывает zip-архив.
    На Python 3.11+ использует metadata_encoding, на более старых версиях
    имена перекодируются вручную через decode_zip_filename.
    """
    if sys.version_info >= (3, 11):
        try:
            return zipfile.ZipFile(zip_path, 'r', metadata_encoding=ZIP_METADATA_ENCODING), False
        except ValueError:
            # Архив с UTF-8 флагом — metadata_encoding не применим
            pass
    # need_decode=True: имена нужно перекодировать вручную
    return zipfile.ZipFile(zip_path, 'r'), True

def process_archive(zip_path, manager_name):
    """Обрабатывает один архив"""
    archive_name = os.path.basename(zip_path)
    print(f"\n📦 Обработка архива: {archive_name}")

    archive_date, portfolio_code = extract_archive_info(archive_name)
    if not archive_date:
        print(f"   ⚠️ Не удалось определить дату/портфель из имени архива")
        return 0, 0

    print(f"   📅 Дата архива: {archive_date.strftime('%d.%m.%Y')}")
    print(f"   📁 Портфель: {portfolio_code}")

    processed = 0
    skipped = 0

    try:
        zip_ref, need_decode = open_zip(zip_path)

        with zip_ref:
            infos = zip_ref.infolist()
            files_count = sum(1 for i in infos if not i.is_dir())
            print(f"   📄 Найдено файлов в архиве: {files_count}")

            for info in infos:
                # Пропускаем папки внутри архива
                if info.is_dir():
                    continue

                # Декодируем имя файла
                if need_decode:
                    decoded_name = decode_zip_filename(info.filename, info.flag_bits)
                else:
                    decoded_name = info.filename

                # Работаем только с именем файла, без пути внутри архива
                base_name = os.path.basename(decoded_name.replace('\\', '/'))

                # Проверяем расширение
                ext = os.path.splitext(base_name)[1].lower()
                if ALLOWED_EXTENSIONS and ext not in ALLOWED_EXTENSIONS:
                    print(f"      ⏭️ Пропускаем (неподдерживаемое расширение): {base_name}")
                    skipped += 1
                    continue

                # Извлекаем информацию из имени файла
                file_date, file_portfolio, broker_code, rule = extract_file_info(base_name)

                if not file_date:
                    print(f"      ⚠️ Не удалось распознать файл: {base_name}")
                    skipped += 1
                    continue

                # Определяем код портфеля
                final_portfolio = portfolio_code if portfolio_code else file_portfolio

                if not final_portfolio:
                    print(f"      ⚠️ Не удалось определить портфель для: {base_name}")
                    skipped += 1
                    continue

                if rule["identifier"] in ["trades", "brok_rpt"]:
                    final_portfolio = broker_code

                # Строим путь назначения
                dest_path = build_destination_path(
                    file_date, rule, manager_name,
                    final_portfolio, broker_code, base_name
                )

                if check_file_exists(dest_path):
                    print(f"      ⏭️ Файл уже существует: {base_name}")
                    skipped += 1
                    continue

                try:
                    # Читаем по объекту info — надёжнее, чем по имени
                    file_data = zip_ref.read(info)
                    os.makedirs(os.path.dirname(dest_path), exist_ok=True)

                    with open(dest_path, 'wb') as f:
                        f.write(file_data)

                    print(f"      ✅ {base_name}")
                    print(f"         -> {dest_path}")
                    processed += 1
                except Exception as e:
                    print(f"      ❌ Ошибка при извлечении {base_name}: {e}")
                    skipped += 1

    except zipfile.BadZipFile:
        print(f"   ❌ Файл не является zip-архивом: {zip_path}")
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

    if not os.path.exists(SOURCE_DIR):
        print(f"❌ Папка-источник не найдена: {SOURCE_DIR}")
        return

    manager_name = get_manager_name(SOURCE_DIR)
    print(f"👤 Управляющий: {manager_name}")
    print("=" * 70)

    files = [f for f in os.listdir(SOURCE_DIR) if os.path.isfile(os.path.join(SOURCE_DIR, f))]
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

    print("\n" + "=" * 70)
    print("📊 ИТОГ:")
    print(f"   ✅ Извлечено файлов: {total_processed}")
    print(f"   ⏭️ Пропущено: {total_skipped}")
    print("=" * 70)

# ==================== ЗАПУСК ====================

if __name__ == "__main__":
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
