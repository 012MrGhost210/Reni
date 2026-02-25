import os
import zipfile
import re
from pathlib import Path
import shutil
import logging
from datetime import datetime

class SCAFileFinder:
    def __init__(self, network_path, output_folder):
        self.network_path = Path(network_path)
        self.output_folder = Path(output_folder)
        
        # Настройка логирования
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('sca_finder.log', encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
        
        self.output_folder.mkdir(exist_ok=True, parents=True)
        
        self.stats = {
            'folders_checked': 0,
            'archives_checked': 0,
            'files_found': 0,
            'errors': 0
        }
        
    def find_files(self):
        """Поиск файлов содержащих 'СЧА Фонд_ПДС' в названии"""
        
        self.logger.info("="*80)
        self.logger.info("🚀 ЗАПУСК ПОИСКА ФАЙЛОВ СЧА Фонд_ПДС")
        self.logger.info("="*80)
        self.logger.info(f"📂 Ищем в: {self.network_path}")
        self.logger.info(f"📁 Сохраняем в: {self.output_folder}")
        self.logger.info("="*80)
        
        # Проверяем доступность исходного пути
        if not self.network_path.exists():
            self.logger.error(f"❌ Исходный путь не существует: {self.network_path}")
            return False
        
        # ПРОСТОЙ ПОИСК - ищем фразу в названии файла
        search_string = "СЧА Фонд_ПДС"
        self.logger.info(f"🔍 Ищем файлы содержащие: '{search_string}'")
        self.logger.info("-"*80)
        
        # Получаем все папки с датами
        date_folders = [f for f in self.network_path.glob("*_*_*") if f.is_dir()]
        date_folders.sort()
        
        self.logger.info(f"Найдено папок для проверки: {len(date_folders)}")
        
        for date_folder in date_folders:
            self.stats['folders_checked'] += 1
            
            # Путь к документам гаранта
            guarant_folder = date_folder / "Документы от Гаранта СД НТД"
            
            if not guarant_folder.exists():
                continue
            
            # Ищем ZIP архивы
            zip_files = list(guarant_folder.glob("Отчеты_*.zip"))
            
            if not zip_files:
                continue
            
            for zip_path in zip_files:
                self.stats['archives_checked'] += 1
                
                try:
                    with zipfile.ZipFile(zip_path, 'r') as zf:
                        files_in_zip = zf.namelist()
                        found_in_this_archive = False
                        
                        for file_in_zip in files_in_zip:
                            # Пропускаем папки
                            if file_in_zip.endswith('/'):
                                continue
                                
                            file_name = Path(file_in_zip).name
                            
                            # ПРОСТАЯ ПРОВЕРКА - содержит ли имя файла искомую фразу
                            if search_string in file_name:
                                found_in_this_archive = True
                                self.stats['files_found'] += 1
                                
                                # Выводим информацию о находке
                                self.logger.info(f"\n📂 Папка: {date_folder.name}")
                                self.logger.info(f"  📦 Архив: {zip_path.name}")
                                self.logger.info(f"     ✅ НАЙДЕН: {file_name}")
                                
                                # Сохраняем файл
                                self._save_file(zf, file_in_zip, date_folder.name, file_name)
                        
                except Exception as e:
                    self.stats['errors'] += 1
                    self.logger.error(f"  ❌ Ошибка при обработке {zip_path.name}: {e}")
        
        # Выводим статистику
        self._print_statistics()
        
        # Создаем файл с отчетом
        if self.stats['files_found'] > 0:
            self._create_summary_file()
        
        return self.stats['files_found'] > 0
    
    def _save_file(self, zip_file, file_in_zip, folder_name, original_filename):
        """Сохранение найденного файла"""
        try:
            # Добавляем дату папки в начало имени для уникальности
            folder_date = folder_name.replace('_', '-')
            
            # Разделяем имя и расширение
            name_parts = original_filename.rsplit('.', 1)
            if len(name_parts) == 2:
                file_base = name_parts[0]
                file_ext = name_parts[1]
                new_filename = f"[{folder_date}]_{file_base}.{file_ext}"
            else:
                new_filename = f"[{folder_date}]_{original_filename}"
            
            # Проверяем уникальность имени
            counter = 1
            save_path = self.output_folder / new_filename
            
            while save_path.exists():
                if len(name_parts) == 2:
                    new_filename = f"[{folder_date}]_{file_base}_{counter}.{file_ext}"
                else:
                    new_filename = f"[{folder_date}]_{original_filename}_{counter}"
                save_path = self.output_folder / new_filename
                counter += 1
            
            # Извлекаем файл
            zip_file.extract(file_in_zip, self.output_folder)
            
            # Если файл извлекся в подпапку, перемещаем в корень
            extracted_path = self.output_folder / file_in_zip
            if extracted_path != save_path:
                if extracted_path.exists():
                    shutil.move(extracted_path, save_path)
                
                # Удаляем пустые папки
                temp_dir = self.output_folder / Path(file_in_zip).parent
                while temp_dir != self.output_folder:
                    try:
                        temp_dir.rmdir()
                        temp_dir = temp_dir.parent
                    except:
                        break
            
            self.logger.info(f"        💾 Сохранен как: {save_path.name}")
            
        except Exception as e:
            self.logger.error(f"        ❌ Ошибка сохранения: {e}")
    
    def _print_statistics(self):
        """Вывод статистики"""
        self.logger.info("\n" + "="*80)
        self.logger.info("📊 ИТОГОВАЯ СТАТИСТИКА:")
        self.logger.info("="*80)
        self.logger.info(f"   📂 Проверено папок с датами: {self.stats['folders_checked']}")
        self.logger.info(f"   📦 Проверено архивов: {self.stats['archives_checked']}")
        self.logger.info(f"   ✅ Найдено файлов: {self.stats['files_found']}")
        self.logger.info(f"   ❌ Ошибок: {self.stats['errors']}")
        
        if self.stats['files_found'] > 0:
            self.logger.info(f"\n📁 Все файлы сохранены в:")
            self.logger.info(f"   {self.output_folder}")
            
            # Показываем список найденных файлов
            saved_files = list(self.output_folder.glob("*.xls*"))
            saved_files.extend(self.output_folder.glob("*.[0-9]*"))  # на случай если нет расширения
            saved_files = [f for f in saved_files if f.is_file()]
            saved_files.sort()
            
            if saved_files:
                self.logger.info(f"\n📋 Найденные файлы ({len(saved_files)}):")
                for i, file_path in enumerate(saved_files[:10], 1):
                    self.logger.info(f"   {i:2d}. {file_path.name}")
                if len(saved_files) > 10:
                    self.logger.info(f"       ... и еще {len(saved_files) - 10} файлов")
    
    def _create_summary_file(self):
        """Создает файл с кратким отчетом"""
        try:
            summary_file = self.output_folder / "!_ОТЧЕТ_О_НАЙДЕННЫХ_ФАЙЛАХ.txt"
            
            saved_files = list(self.output_folder.glob("*.xls*"))
            saved_files.extend(self.output_folder.glob("*.[0-9]*"))
            saved_files = [f for f in saved_files if f.is_file()]
            saved_files.sort()
            
            with open(summary_file, 'w', encoding='utf-8') as f:
                f.write("="*60 + "\n")
                f.write("ОТЧЕТ О ПОИСКЕ ФАЙЛОВ СЧА Фонд_ПДС\n")
                f.write(f"Дата поиска: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n")
                f.write("="*60 + "\n\n")
                
                f.write(f"Всего найдено файлов: {len(saved_files)}\n\n")
                f.write("Список найденных файлов:\n")
                f.write("-"*40 + "\n")
                
                for i, file_path in enumerate(saved_files, 1):
                    f.write(f"{i:3d}. {file_path.name}\n")
                
                f.write("\n" + "="*60 + "\n")
            
            self.logger.info(f"\n📄 Создан файл с отчетом: {summary_file.name}")
            
        except Exception as e:
            self.logger.error(f"Ошибка создания отчета: {e}")

def main():
    # Путь где ищем архивы
    search_path = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\01.Перечень имущества Фонда (СД)"
    
    # Путь куда сохраняем все найденные файлы
    output_path = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\Фонд СЧА"
    
    print("\n" + "="*80)
    print("🔍 ПРОГРАММА ПОИСКА ФАЙЛОВ СЧА Фонд_ПДС")
    print("="*80)
    print(f"📂 Поиск в: {search_path}")
    print(f"📁 Сохранение в: {output_path}")
    print("="*80)
    
    # Проверяем доступность путей
    search_path_obj = Path(search_path)
    
    if not search_path_obj.exists():
        print("\n❌ ОШИБКА: Не удалось подключиться к исходной папке!")
        print(f"   Путь: {search_path}")
        print("\nПроверьте:")
        print("1. Подключение к VPN")
        print("2. Откройте папку в проводнике")
        input("\nНажмите Enter для выхода...")
        return
    
    # Создаем и запускаем поисковик
    finder = SCAFileFinder(search_path, output_path)
    finder.find_files()
    
    print("\n" + "="*80)
    print("✅ РАБОТА ЗАВЕРШЕНА")
    print("="*80)
    print(f"📁 Все файлы сохранены в: {output_path}")
    print("\n📄 Подробный лог: sca_finder.log")
    
    input("\nНажмите Enter для выхода...")

if __name__ == "__main__":
    main()
