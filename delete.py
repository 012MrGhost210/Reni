import os
from pathlib import Path
import shutil

def clean_folder_except_results(folder_path):
    """
    Удаляет все файлы и папки из указанной директории,
    кроме тех, в названии которых есть "!_РЕЗУЛЬТАТЫ"
    """
    folder = Path(folder_path)
    
    if not folder.exists():
        print(f"❌ Папка не существует: {folder_path}")
        return
    
    print(f"🧹 Очищаем папку: {folder_path}")
    print(f"   Сохраняем файлы с '!_РЕЗУЛЬТАТЫ' в названии")
    print("-" * 60)
    
    deleted_count = 0
    kept_count = 0
    
    # Проходим по всем элементам в папке
    for item in folder.iterdir():
        # Проверяем, нужно ли сохранить этот элемент
        if "!_РЕЗУЛЬТАТЫ" in item.name:
            print(f"   ✅ СОХРАНЯЕМ: {item.name}")
            kept_count += 1
            continue
        
        # Удаляем всё остальное
        try:
            if item.is_file():
                item.unlink()  # Удаляем файл
                print(f"   ❌ Удален файл: {item.name}")
                deleted_count += 1
            elif item.is_dir():
                shutil.rmtree(item)  # Удаляем папку со всем содержимым
                print(f"   ❌ Удалена папка: {item.name}")
                deleted_count += 1
        except Exception as e:
            print(f"   ⚠️ Ошибка при удалении {item.name}: {e}")
    
    print("-" * 60)
    print(f"📊 ИТОГ: Удалено: {deleted_count}, Сохранено: {kept_count}")
    print(f"✅ Очистка завершена!")

# Использование
folder_path = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\Фонд СЧА"
clean_folder_except_results(folder_path)
