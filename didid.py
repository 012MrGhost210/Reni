import os
import shutil
from pathlib import Path

def move_contents_from_nested_folders(source_root, target_base, company_folders):
    """
    Переносит всё содержимое из указанных папок (company_folders),
    которые могут лежать на любом уровне вложенности внутри source_root,
    в соответствующие подпапки внутри target_base.
    
    Args:
        source_root (str или Path): Путь к корневой папке, где лежат папки с компаниями.
        target_base (str или Path): Путь к папке Y, куда всё переносим.
        company_folders (list): Список точных названий папок компаний.
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
    
    for company_folder_name in company_folders:
        # Создаем соответствующую папку в целевой директории
        company_target_path = target_base / company_folder_name
        company_target_path.mkdir(parents=True, exist_ok=True)
        
        # Ищем все папки с таким названием в исходной директории
        found_paths = list(source_root.rglob(company_folder_name))
        
        if not found_paths:
            print(f"Предупреждение: Папка '{company_folder_name}' не найдена.")
            continue
        
        print(f"\nОбработка папок для: {company_folder_name}")
        
        for company_path in found_paths:
            # Проверяем, что это действительно папка, а не файл с таким именем
            if not company_path.is_dir():
                print(f"  Пропуск: {company_path} — это не папка.")
                continue
            
            print(f"  Найдена папка: {company_path}")
            
            # Перебираем всё содержимое найденной папки компании
            for item_path in company_path.iterdir():
                # Формируем путь назначения в соответствующей подпапке
                destination_path = company_target_path / item_path.name
                
                # Выполняем перенос (с заменой существующих файлов)
                try:
                    # Если файл уже существует, удаляем его перед перемещением
                    if destination_path.exists():
                        if destination_path.is_dir():
                            shutil.rmtree(destination_path)
                        else:
                            destination_path.unlink()
                        print(f"    Замена существующего: {destination_path.name}")
                    
                    # Перемещаем файл/папку
                    shutil.move(str(item_path), str(destination_path))
                    print(f"    Перемещено: {item_path.name} -> {company_folder_name}/")
                    moved_count += 1
                    
                except Exception as e:
                    print(f"    Ошибка при перемещении {item_path.name}: {e}")
            
            # После переноса всего содержимого, можно удалить пустую папку компании
            try:
                # Проверяем, пуста ли папка
                if not any(company_path.iterdir()):
                    company_path.rmdir()
                    print(f"    Удалена пустая папка: {company_path}")
            except (OSError, Exception):
                pass  # Папка не пуста или ошибка доступа
    
    print("-" * 50)
    print(f"Готово! Перемещено элементов: {moved_count}")

if __name__ == "__main__":
    # ВАШИ ПУТИ
    source_folder = "X:/IRI/1C"  # Путь к папке с матрешками
    target_folder = "Y:/IRI/1C"  # Путь к целевой папке Y
    
    # Список точных названий папок компаний
    companies = [
        "7702358512-ООО -УК Райффайзен-",
        "7744000951-АО -УК -СПУТНИК - УПРАВЛЕНИЕ КАПИТАЛОМ-",
        "7825489723-ТКБ Инвестмент Партнерс (АО)"
    ]
    
    # Запуск
    move_contents_from_nested_folders(source_folder, target_folder, companies)
