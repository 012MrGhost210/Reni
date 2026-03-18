import os
import shutil
import re
from datetime import datetime

# --- НАСТРОЙКИ ---
source_root = r"путь_к_вашей_основной_папке"  # Укажите путь к папке, где лежат 3 подпапки
destination_folder = r"путь_к_папке_X"       # Укажите путь к папке X (куда копировать)

# Соответствие папок: (путь, что_ищем, комментарий для лога)
TARGETS = [
    {
        "name": "Спутник",
        "path": "7744000951-АО -УК -СПУТНИК - УПРАВЛЕНИЕ КАПИТАЛОМ-",
        "search_pattern": "Вознаграждение",
        "log_name": "Спутнику"
    },
    {
        "name": "Райффайзен",
        "path": "7702358512-ООО -УК Райффайзен-",
        "search_pattern": "Отчет по СЧА",
        "log_name": "Райффайзен"
    },
    {
        "name": "ТКБ",
        "path": "7825489723-ТКБ Инвестмент Партнерс (АО)",
        "search_pattern": "Сводная РСА-СЧА",
        "log_name": "ТКБ"
    }
]

# --- ФУНКЦИИ ---
def remove_text_in_braces(filename):
    """Удаляет из имени файла все, что находится в фигурных скобках, включая сами скобки."""
    # Удаляем {текст} (в том числе, если скобки не закрыты, но мы предполагаем корректный формат)
    new_name = re.sub(r'\{.*?\}', '', filename)
    # Также удаляем возможные лишние пробелы, которые могли появиться
    new_name = re.sub(r'\s+', ' ', new_name).strip()
    return new_name

def is_today(date):
    """Проверяет, является ли дата сегодняшней."""
    today = datetime.now().date()
    return date.date() == today

def process_folder(target, source_root, destination_folder):
    """Ищет файл в конкретной подпапке, проверяет дату и копирует."""
    folder_path = os.path.join(source_root, target["path"])
    log_name = target["log_name"]
    search_pattern = target["search_pattern"]
    found_file = None
    found_date = None

    print(f"Проверяем папку: {folder_path}")

    if not os.path.exists(folder_path):
        print(f"  Папка {folder_path} не найдена. Пропускаем.")
        return

    # Ищем файл, содержащий нужную фразу
    for filename in os.listdir(folder_path):
        if search_pattern.lower() in filename.lower():  # регистронезависимый поиск
            file_path = os.path.join(folder_path, filename)
            if os.path.isfile(file_path):
                mod_time = os.path.getmtime(file_path)
                mod_date = datetime.fromtimestamp(mod_time)
                found_file = filename
                found_date = mod_date
                break  # берем первый подходящий файл

    # Если файл не найден
    if not found_file:
        print(f"  Файл с '{search_pattern}' не найден в папке {log_name}.")
        return

    # Проверяем дату
    if not is_today(found_date):
        print(f"  Найден файл '{found_file}', но его дата ({found_date.date()}) не сегодняшняя. Пропускаем. (Данных по {log_name} нет)")
        return

    # Копируем файл
    source_file = os.path.join(folder_path, found_file)
    new_filename = remove_text_in_braces(found_file)
    destination_file = os.path.join(destination_folder, new_filename)

    # Если в папке назначения уже есть файл с таким именем, добавляем префикс, чтобы не перезаписать
    base, ext = os.path.splitext(new_filename)
    counter = 1
    while os.path.exists(destination_file):
        destination_file = os.path.join(destination_folder, f"{base}_{counter}{ext}")
        counter += 1

    try:
        shutil.copy2(source_file, destination_file)  # copy2 сохраняет метаданные
        print(f"  УСПЕХ: Скопирован '{found_file}' -> '{os.path.basename(destination_file)}'")
    except Exception as e:
        print(f"  ОШИБКА при копировании: {e}")

# --- ЗАПУСК ---
if __name__ == "__main__":
    # Проверяем, существует ли папка назначения, если нет - создаем
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)
        print(f"Создана папка назначения: {destination_folder}")

    print(f"Начинаем обработку. Папка назначения: {destination_folder}\n")
    for target in TARGETS:
        process_folder(target, source_root, destination_folder)
        print("-" * 50)

    print("Готово!")
