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
                # Получаем список всех файлов в архиве
                all_files = zf.namelist()
                print(f"     Всего файлов в архиве: {len(all_files)}")
                
                # ВЫВОДИМ ВСЕ ФАЙЛЫ ИЗ АРХИВА (первые 20)
                print(f"     Содержимое архива (первые 20):")
                for i, file_path in enumerate(all_files[:20]):
                    if not file_path.endswith('/'):
                        file_name = Path(file_path).name
                        print(f"       {i+1:2d}. {file_name}")
                
                if len(all_files) > 20:
                    print(f"       ... и еще {len(all_files) - 20} файлов")
                
                # Ищем нужный файл
                found = False
                for file_in_zip in all_files:
                    if file_in_zip.endswith('/'):
                        continue
                    
                    file_name = Path(file_in_zip).name
                    
                    # ПРОСТАЯ ПРОВЕРКА - ищем фразу
                    if "СЧА Фонд_ПДС" in file_name:
                        found = True
                        total_files += 1
                        
                        print(f"\n     ✅ НАЙДЕН: {file_name}")
                        
                        # Сохраняем файл
                        new_name = f"[{date_folder.name}]_{file_name}"
                        save_path = Path(output_path) / new_name
                        
                        # Если такой файл уже есть, добавляем номер
                        counter = 1
                        while save_path.exists():
                            name_parts = new_name.rsplit('.', 1)
                            if len(name_parts) == 2:
                                new_name = f"{name_parts[0]}_{counter}.{name_parts[1]}"
                            else:
                                new_name = f"{new_name}_{counter}"
                            save_path = Path(output_path) / new_name
                            counter += 1
                        
                        # Извлекаем
                        zf.extract(file_in_zip, output_path)
                        
                        # Перемещаем если нужно
                        extracted = Path(output_path) / file_in_zip
                        if extracted != save_path:
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

if __name__ == "__main__":
    main()

 Содержимое архива (первые 20):
        1. 20.02.2026_ÅÑαÑτÑ¡∞ ¿¼πΘÑßΓóá îÇèÉÄ.xlsx
        2. 20.02.2026_ÅÑαÑτÑ¡∞ ¿¼πΘÑßΓóá ôè æÅôÆìêè - ôÅÉÇéïàìêà èÇÅêÆÇïÄî (ä.ô. 080825_1).xlsx
        3. 20.02.2026_ÅÑαÑτÑ¡∞ ¿¼πΘÑßΓóá ôè æÅôÆìêè - ôÅÉÇéïàìêà èÇÅêÆÇïÄî (ä.ô. 301024_1).xlsx
        4. 20.02.2026_ÅÑαÑτÑ¡∞ ¿¼πΘÑßΓóá ö«¡ñ_299-öç.xlsx
        5. 20.02.2026_ÅÑαÑτÑ¡∞ ¿¼πΘÑßΓóá ö«¡ñ_Åäæ.xlsx
        6. 20.02.2026_ÅÑαÑτÑ¡∞ ßñÑ½«¬ îÇèÉÄ.xls
        7. 20.02.2026_ÅÑαÑτÑ¡∞ ßñÑ½«¬ ôè æÅôÆìêè - ôÅÉÇéïàìêà èÇÅêÆÇïÄî (ä.ô. 080825_1).xls
        8. 20.02.2026_ÅÑαÑτÑ¡∞ ßñÑ½«¬ ôè æÅôÆìêè - ôÅÉÇéïàìêà èÇÅêÆÇïÄî (ä.ô. 301024_1).xls
        9. 20.02.2026_ÅÑαÑτÑ¡∞ ßñÑ½«¬ ö«¡ñ_299-öç.xls
       10. 20.02.2026_ÅÑαÑτÑ¡∞ ßñÑ½«¬ ö«¡ñ_Åäæ.xls
       11. 20.02.2026_æÉæ.xls
       12. 20.02.2026_æùÇ ôè æÅôÆìêè - ôÅÉÇéïàìêà èÇÅêÆÇïÄî (ä.ô. 080825_1) .xls
       13. 20.02.2026_æùÇ ôè æÅôÆìêè - ôÅÉÇéïàìêà èÇÅêÆÇïÄî (ä.ô. 301024_1) .xls
       14. 20.02.2026_æùÇ ö«¡ñ_299-öç.xls
       15. 20.02.2026_æùÇ ö«¡ñ_Åäæ.xls
       16. 20.02.2026_Ç¬Γ »ÑαÑαáßτÑΓá ßΓ«¿¼«ßΓ¿ á¬Γ¿ó«ó ÅÉ.xls
       17. 20.02.2026_ÉæÇ.xls
