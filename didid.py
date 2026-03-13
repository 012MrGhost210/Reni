import os
import shutil
from pathlib import Path

def move_contents_from_nested_folders(source_root, target_base, company_folders):
    """
    Переносит всё содержимое из указанных папок (company_folders),
    которые могут лежать на любом уровне вложенности внутри source_root,
    в директорию target_base.

    Args:
        source_root (str или Path): Путь к корневой папке, где лежат папки с компаниями.
        target_base (str или Path): Путь к папке Y, куда всё переносим.
        company_folders (list): Список точных названий папок компаний (например,
                                 ['7702358512-ООО -УК Райффайзен-', ...]).
    """
    source_root = Path(source_root)
    target_base = Path(target_base)

    # Создаем целевую папку, если её нет
    target_base.mkdir(parents=True, exist_ok=True)

    print(f"Поиск в: {source_root}")
    print(f"Целевая папка: {target_base}")
    print(f"Ищем папки: {company_folders}")
    print("-" * 50)

    moved_count = 0

    # Используем rglob для рекурсивного поиска по всем папкам
    # **/ означает "во всех поддиректориях на любом уровне"
    for company_folder_name in company_folders:
        # Ищем все пути, которые заканчиваются на название нужной папки
        # rglob вернет генератор всех совпадений
        found_paths = list(source_root.rglob(company_folder_name))

        if not found_paths:
            print(f"Предупреждение: Папка '{company_folder_name}' не найдена.")

        for company_path in found_paths:
            # Проверяем, что это действительно папка, а не файл с таким именем
            if not company_path.is_dir():
                print(f"Пропуск: {company_path} — это не папка.")
                continue

            print(f"\nНайдена папка: {company_path}")

            # Перебираем всё содержимое найденной папки компании
            for item_path in company_path.iterdir():
                # Формируем новое имя для элемента в целевой папке
                # Добавляем префикс с именем "компании_матрешки", чтобы избежать коллизий
                # Заменяем пробелы и спецсимволы на '_' для надежности (опционально)
                safe_company_name = company_path.name.replace(' ', '_').replace('-', '_')
                # Берем имя файла/папки
                original_name = item_path.name
                # Создаем новое имя с уникальным префиксом
                new_name = f"{safe_company_name}_{original_name}"
                destination_path = target_base / new_name

                # Если файл с таким именем уже существует, добавляем счетчик
                counter = 1
                original_stem = destination_path.stem
                original_suffix = destination_path.suffix
                while destination_path.exists():
                    new_name_with_counter = f"{original_stem}_{counter}{original_suffix}"
                    destination_path = target_base / new_name_with_counter
                    counter += 1

                # Выполняем перенос
                try:
                    shutil.move(str(item_path), str(destination_path))
                    print(f"  Перемещено: {item_path.name} -> {destination_path.name}")
                    moved_count += 1
                except Exception as e:
                    print(f"  Ошибка при перемещении {item_path.name}: {e}")

            # После переноса всего содержимого, можно удалить пустую папку компании,
            # если это требуется (раскомментируйте следующие строки при необходимости).
            # try:
            #     company_path.rmdir() # Удалит только если папка пуста
            #     print(f"Удалена пустая папка: {company_path}")
            # except OSError:
            #     pass # Папка не пуста (например, если что-то не перенеслось) или ошибка доступа

    print("-" * 50)
    print(f"Готово. Перемещено элементов: {moved_count}")

if __name__ == "__main__":
    # НАСТРОЙКИ (замените на свои пути)
    source_folder = "X:/Путь/к/вашей/матрешке"  # Папка, где лежат папки с компаниями
    target_folder = "Y:/Целевая/папка"          # Папка Y, куда всё переносим

    # Список точных названий папок, которые нужно обработать
    companies = [
        "7702358512-ООО -УК Райффайзен-",
        "7744000951-АО -УК -СПУТНИК - УПРАВЛЕНИЕ КАПИТАЛОМ-",
        "7825489723-ТКБ Инвестмент Партнерс (АО)"
    ]

    # Запуск
    move_contents_from_nested_folders(source_folder, target_folder, companies)
