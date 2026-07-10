\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\Diadoc\diadoc_connector\Документооборот завершен\7710183778-АО УК -ПЕРВАЯ-
{Документооборот завершен}Выгрузка операций_08072026
{Документооборот завершен}I02_514833_k_d_260706
{Документооборот завершен}2026.07.08_27011_журнал учета ДС
Выгрузка операций_*ДДММГГГГ*
I02_514833_k_d_*ГГММДД*
*ГГГГ.ММ.ДД*_27011_журнал учета ДС
M:\Финансовый департамент\Treasury\Отчеты Брокера и Спецдепозитария\10.НАПФ - Ценные бумаги\2026\УК Первая\06.Июнь
M:\Финансовый департамент\Treasury\Отчеты Брокера и Спецдепозитария\2 Отчеты брокера\2026\06.Июнь 2026\УК Первая
M:\Финансовый департамент\Treasury\Отчеты Брокера и Спецдепозитария\6.Журнал учета операций\2026\06.Июнь 2026\УК Первая
M:\Финансовый департамент\Treasury\Отчеты Брокера и Спецдепозитария\10.НАПФ - Ценные бумаги\*ГГГГ*\*Управляющий*\*Месяц*
M:\Финансовый департамент\Treasury\Отчеты Брокера и Спецдепозитария\2 Отчеты брокера\*ГГГГ*\*Месяц*\*Управляющий*
M:\Финансовый департамент\Treasury\Отчеты Брокера и Спецдепозитария\6.Журнал учета операций\*ГГГГ*\*Месяц*\*Управляющий*

                                                                                                            
                                                                                                            
import os
import shutil
import re
from datetime import datetime
from pathlib import Path

# ==================== НАСТРОЙКИ ====================

# Базовый путь назначения
BASE_DEST_DIR = r"M:\Финансовый департамент\Treasury\Отчеты Брокера и Спецдепозитария"

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

# ==================== КОНФИГУРАЦИЯ ФАЙЛОВ ====================

# Каждое правило - как определить дату и куда положить файл
FILE_RULES = [
    {
        # Правило 1: Выгрузка операций
        "name_pattern": "{Документооборот завершен}Выгрузка операций_",
        "date_regex": r"(\d{8})",  # 08072026 -> 8 цифр
        "date_format": "%d%m%Y",    # ДДММГГГГ
        "destination": r"10.НАПФ - Ценные бумаги\*ГГГГ*\*Управляющий*\*Месяц*",
        "remove_prefix": True,      # Удаляем "{Документооборот завершен}"
    },
    {
        # Правило 2: Журнал учета операций (I02)
        "name_pattern": "{Документооборот завершен}I02_514833_k_d_",
        "date_regex": r"(\d{6})",   # 260706 -> 6 цифр
        "date_format": "%y%m%d",    # ГГММДД
        "destination": r"6.Журнал учета операций\*ГГГГ*\*Месяц*\*Управляющий*",
        "remove_prefix": True,
    },
    {
        # Правило 3: Отчеты брокера (журнал учета ДС)
        "name_pattern": "{Документооборот завершен}",
        "date_regex": r"(\d{4}\.\d{2}\.\d{2})",  # 2026.07.08
        "date_format": "%Y.%m.%d",               # ГГГГ.ММ.ДД
        "destination": r"2 Отчеты брокера\*ГГГГ*\*Месяц*\*Управляющий*",
        "remove_prefix": True,
        "subfolder": "журнал учета ДС",  # дополнительная папка
    },
]

# ==================== ОСНОВНАЯ ЛОГИКА ====================

def get_manager_name(source_dir):
    """Извлекает имя управляющего из пути и преобразует по маппингу"""
    # Берем последнюю папку из пути
    normalized_path = source_dir.replace('\\', '/').rstrip('/')
    raw_name = os.path.basename(normalized_path)
    
    # Применяем маппинг
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

def extract_date_from_filename(filename, rule):
    """Извлекает дату из имени файла по правилу"""
    # Проверяем, начинается ли файл с нужного паттерна
    if not filename.startswith(rule["name_pattern"]):
        return None
    
    # Убираем префикс, если нужно
    if rule.get("remove_prefix", False):
        name_part = filename[len(rule["name_pattern"]):]
    else:
        name_part = filename
    
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
    month = f"{date_obj.month:02d}"
    
    # Получаем шаблон пути
    dest_path = rule["destination"]
    
    # Заменяем переменные
    dest_path = dest_path.replace("*ГГГГ*", year)
    dest_path = dest_path.replace("*Месяц*", month_folder)
    dest_path = dest_path.replace("*Управляющий*", manager_name)
    
    # Добавляем подпапку, если есть
    if "subfolder" in rule:
        dest_path = os.path.join(dest_path, rule["subfolder"])
    
    # Полный путь
    full_path = os.path.join(BASE_DEST_DIR, dest_path)
    
    # Создаем папки
    Path(full_path).mkdir(parents=True, exist_ok=True)
    
    # Очищаем имя файла от префикса
    clean_name = original_filename
    if rule.get("remove_prefix", False):
        clean_name = clean_name.replace(rule["name_pattern"], "")
    
    return os.path.join(full_path, clean_name)

def normalize_filename(filename):
    """Нормализует имя файла для сравнения (удаляет +, пробелы и т.д.)"""
    name_without_ext, ext = os.path.splitext(filename)
    # Удаляем лишние символы
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
        
        # Ищем подходящее правило
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
