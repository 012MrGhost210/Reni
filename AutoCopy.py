import os
import shutil
import re
from datetime import datetime

# --- НАСТРОЙКИ ---
source_root = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\diadoc_connector\Документооборот завершён"  # Укажите путь к папке, где лежат 3 подпапки
destination_folder = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\25.Автоматизация\NAV for DI"       # Укажите путь к папке X (куда копировать)

# Соответствие папок: (путь, что_ищем, короткое_имя_для_вывода)
TARGETS = [
    {
        "short_name": "Спутник",
        "path": "7744000951-АО -УК -СПУТНИК - УПРАВЛЕНИЕ КАПИТАЛОМ-",
        "search_pattern": "Вознаграждение"
    },
    {
        "short_name": "Райф",
        "path": "7702358512-ООО -УК Райффайзен-",
        "search_pattern": "Отчет по СЧА"
    },
    {
        "short_name": "ТКБ",
        "path": "7825489723-ТКБ Инвестмент Партнерс (АО)",
        "search_pattern": "Сводная РСА-СЧА"
    }
]

# --- ФУНКЦИИ ---
def remove_text_in_braces(filename):
    """Удаляет из имени файла все, что находится в фигурных скобках, включая сами скобки."""
    new_name = re.sub(r'\{.*?\}', '', filename)
    new_name = re.sub(r'\s+', ' ', new_name).strip()
    return new_name

def is_today(date):
    """Проверяет, является ли дата сегодняшней."""
    today = datetime.now().date()
    return date.date() == today

def clear_folder(folder_path):
    """Очищает папку: удаляет все файлы и подпапки."""
    if os.path.exists(folder_path):
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print(f"  Ошибка при очистке {file_path}: {e}")

def process_folder(target, source_root, destination_folder):
    """Ищет САМЫЙ СВЕЖИЙ файл в конкретной подпапке, проверяет дату и копирует."""
    folder_path = os.path.join(source_root, target["path"])
    short_name = target["short_name"]
    search_pattern = target["search_pattern"]
    
    matching_files = []
    found_date = None

    if not os.path.exists(folder_path):
        print(f"{short_name} не ок (папка не найдена)")
        return

    # Собираем ВСЕ файлы, содержащие нужную фразу
    for filename in os.listdir(folder_path):
        if search_pattern.lower() in filename.lower():
            file_path = os.path.join(folder_path, filename)
            if os.path.isfile(file_path):
                mod_time = os.path.getmtime(file_path)
                mod_date = datetime.fromtimestamp(mod_time)
                matching_files.append((filename, mod_date, file_path))

    # Если файлы не найдены
    if not matching_files:
        print(f"{short_name} не ок (файлы не найдены)")
        return

    # Сортируем по дате (самый свежий первый)
    matching_files.sort(key=lambda x: x[1], reverse=True)
    latest_file, latest_date, latest_path = matching_files[0]

    # Проверяем дату самого свежего файла
    if not is_today(latest_date):
        print(f"{short_name} не ок (самый свежий файл от {latest_date.date()})")
        # Для отладки можно раскомментировать:
        # print(f"  Найдены файлы: {[f[0] for f in matching_files]}")
        return

    # Копируем файл
    new_filename = remove_text_in_braces(latest_file)
    destination_file = os.path.join(destination_folder, new_filename)

    # Если в папке назначения уже есть файл с таким именем, добавляем префикс
    base, ext = os.path.splitext(new_filename)
    counter = 1
    while os.path.exists(destination_file):
        destination_file = os.path.join(destination_folder, f"{base}_{counter}{ext}")
        counter += 1

    try:
        shutil.copy2(latest_path, destination_file)
        print(f"{short_name} ок")
    except Exception as e:
        print(f"{short_name} не ок (ошибка копирования)")

# --- ЗАПУСК ---
if __name__ == "__main__":
    # Очищаем папку назначения
    print("Очищаем папку X...")
    clear_folder(destination_folder)
    
    # Создаем папку назначения, если её нет
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)
    
    print("\nРезультаты:")
    for target in TARGETS:
        process_folder(target, source_root, destination_folder)

    print("\nГотово!")
