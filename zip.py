import os
import zipfile
from pathlib import Path
import shutil

# Пути
search_path = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\01.Перечень имущества Фонда (СД)"
output_path = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\Фонд СЧА"

print("="*80)
print("ПОИСК ФАЙЛОВ СЧА Фонд_ПДС")
print("="*80)
print(f"Ищем в: {search_path}")
print(f"Сохраняем в: {output_path}")
print("="*80)

# Создаем папку для сохранения
Path(output_path).mkdir(exist_ok=True)

# Счетчики
total_archives = 0
total_files = 0

# Проходим по папкам с датами
for date_folder in Path(search_path).glob("*_*_*"):
    if not date_folder.is_dir():
        continue
    
    print(f"\n📂 Папка: {date_folder.name}")
    
    # Путь к документам гаранта
    guarant_folder = date_folder / "Документы от Гаранта СД НТД"
    if not guarant_folder.exists():
        print(f"  ⚠️ Нет папки гаранта")
        continue
    
    # Ищем ZIP архивы
    zip_files = list(guarant_folder.glob("*.zip"))
    if not zip_files:
        print(f"  ⚠️ Нет ZIP архивов")
        continue
    
    print(f"  Найдено архивов: {len(zip_files)}")
    
    for zip_path in zip_files:
        total_archives += 1
        print(f"\n  📦 Архив: {zip_path.name}")
        
        try:
            with zipfile.ZipFile(zip_path, 'r') as zf:
                # Получаем список всех файлов в архиве с правильной кодировкой
                all_files = []
                for file_info in zf.infolist():
                    # Пробуем разные кодировки для русских букв
                    try:
                        # Пробуем CP866 (часто используется в Windows для русских имен)
                        filename = file_info.filename.encode('cp437').decode('cp866')
                    except:
                        try:
                            # Пробуем CP1251
                            filename = file_info.filename.encode('cp437').decode('cp1251')
                        except:
                            # Если не получилось, оставляем как есть
                            filename = file_info.filename
                    
                    all_files.append((file_info.filename, filename))
                
                print(f"     Всего файлов в архиве: {len(all_files)}")
                
                # ВЫВОДИМ ВСЕ ФАЙЛЫ ИЗ АРХИВА С РУССКИМИ НАЗВАНИЯМИ
                print(f"     Содержимое архива (первые 20):")
                for i, (orig_name, rus_name) in enumerate(all_files[:20]):
                    if not orig_name.endswith('/'):
                        print(f"       {i+1:2d}. {rus_name}")
                
                if len(all_files) > 20:
                    print(f"       ... и еще {len(all_files) - 20} файлов")
                
                # Ищем нужный файл
                found = False
                for orig_name, rus_name in all_files:
                    if orig_name.endswith('/'):
                        continue
                    
                    # Ищем фразу "СЧА Фонд_ПДС" в русском названии
                    if "СЧА Фонд_ПДС" in rus_name:
                        found = True
                        total_files += 1
                        
                        print(f"\n     ✅ НАЙДЕН: {rus_name}")
                        
                        # Сохраняем файл с правильным именем
                        new_name = f"[{date_folder.name}]_{rus_name}"
                        # Очищаем имя от недопустимых символов
                        new_name = "".join(c for c in new_name if c not in '<>:"/\\|?*')
                        
                        save_path = Path(output_path) / new_name
                        
                        # Если такой файл уже есть, добавляем номер
                        counter = 1
                        while save_path.exists():
                            name_parts = new_name.rsplit('.', 1)
                            if len(name_parts) == 2:
                                new_name = f"{name_parts[0]}_{counter}.{name_parts[1]}"
                            else:
                                new_name = f"{new_name}_{counter}"
                            new_name = "".join(c for c in new_name if c not in '<>:"/\\|?*')
                            save_path = Path(output_path) / new_name
                            counter += 1
                        
                        # Извлекаем с оригинальным именем
                        zf.extract(orig_name, output_path)
                        
                        # Переименовываем в русское название
                        extracted = Path(output_path) / orig_name
                        if extracted.exists():
                            # Создаем папки если нужно
                            save_path.parent.mkdir(exist_ok=True)
                            shutil.move(extracted, save_path)
                        
                        print(f"        💾 Сохранен как: {save_path.name}")
                
                if not found:
                    print(f"     ❌ Файл 'СЧА Фонд_ПДС' не найден в этом архиве")
                    
        except Exception as e:
            print(f"     ❌ Ошибка при открытии архива: {e}")

# Итог
print("\n" + "="*80)
print("ГОТОВО!")
print("="*80)
print(f"Проверено архивов: {total_archives}")
print(f"Найдено файлов: {total_files}")
print(f"Все файлы сохранены в: {output_path}")
print("="*80)

input("\nНажмите Enter для выхода...")
