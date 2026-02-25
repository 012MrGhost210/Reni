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
        """Поиск файлов по точному паттерну: ДД.ММ.ГГГГ_СЧА Фонд_ПДС.xls"""
        
        self.logger.info("="*80)
        self.logger.info("🚀 ЗАПУСК ПОИСКА ФАЙЛОВ СЧА Фонд_ПДС")
        self.logger.info("="*80)
        self.logger.info(f"📂 Ищем в: {self.network_path}")
        self.logger.info(f"📁 Сохраняем в: {self.output_folder}")
        self.logger.info("-"*80)
        
        # Проверяем доступность исходного пути
        if not self.network_path.exists():
            self.logger.error(f"❌ Исходный путь не существует: {self.network_path}")
            self.logger.error("   Проверьте подключение к сетевому диску")
            return False
        
        # ТОЧНЫЙ ПАТТЕРН ПОИСКА - только такой формат:
        # ДД.ММ.ГГГГ_СЧА Фонд_ПДС.xls
        date_pattern = r'\d{2}\.\d{2}\.\d{4}'  # 29.12.2025
        exact_filename = f"{date_pattern}_СЧА Фонд_ПДС\\.xls"
        
        self.logger.info(f"🔍 Ищем файлы по паттерну: ДД.ММ.ГГГГ_СЧА Фонд_ПДС.xls")
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
                self.logger.info(f"📂 {date_folder.name}: пропускаем (нет папки гаранта)")
                continue
            
            # Ищем ZIP архивы
            zip_files = list(guarant_folder.glob("Отчеты_*.zip"))
            
            if not zip_files:
                self.logger.info(f"📂 {date_folder.name}: нет архивов")
                continue
            
            self.logger.info(f"\n📂 Папка: {date_folder.name} (архивов: {len(zip_files)})")
            
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
                            
                            # Проверяем точное совпадение с паттерном
                            if re.match(exact_filename, file_name):
                                found_in_this_archive = True
                                self.stats['files_found'] += 1
                                
                                self.logger.info(f"  📦 {zip_path.name}")
                                self.logger.info(f"     ✅ НАЙДЕН: {file_name}")
                                
                                # Сохраняем файл с префиксом из даты папки
                                self._save_file(zf, file_in_zip, date_folder.name, file_name)
                        
                        if not found_in_this_archive:
                            self.logger.info(f"  📦 {zip_path.name}: файл не найден")
                            
                except zipfile.BadZipFile:
                    self.stats['errors'] += 1
                    self.logger.error(f"  📦 {zip_path.name}: ❌ испорченный ZIP")
                except Exception as e:
                    self.stats['errors'] += 1
                    self.logger.error(f"  📦 {zip_path.name}: ❌ ошибка {e}")
        
        # Выводим статистику
        self._print_statistics()
        
        # Создаем файл с отчетом
        if self.stats['files_found'] > 0:
            self._create_summary_file()
        
        return self.stats['files_found'] > 0
    
    def _save_file(self, zip_file, file_in_zip, folder_name, original_filename):
        """Сохранение найденного файла с префиксом из папки"""
        try:
            # Добавляем дату папки в начало имени для уникальности
            # Папка 2026_01_12 -> префикс [2026-01-12]
            folder_date = folder_name.replace('_', '-')
            new_filename = f"[{folder_date}]_{original_filename}"
            
            # Проверяем уникальность имени
            counter = 1
            save_path = self.output_folder / new_filename
            
            while save_path.exists():
                name_parts = new_filename.rsplit('.', 1)
                if len(name_parts) == 2:
                    new_filename = f"{name_parts[0]}_{counter}.{name_parts[1]}"
                else:
                    new_filename = f"{new_filename}_{counter}"
                save_path = self.output_folder / new_filename
                counter += 1
            
            # Создаем временную папку для распаковки
            temp_extract = self.output_folder / "_temp"
            temp_extract.mkdir(exist_ok=True)
            
            # Извлекаем файл
            zip_file.extract(file_in_zip, temp_extract)
            
            # Перемещаем с новым именем
            extracted_path = temp_extract / file_in_zip
            if extracted_path.exists():
                shutil.move(extracted_path, save_path)
            
            # Удаляем временную папку
            shutil.rmtree(temp_extract, ignore_errors=True)
            
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
            
            # Показываем первые несколько файлов
            saved_files = list(self.output_folder.glob("[*]*.xls"))
            if saved_files:
                self.logger.info(f"\n📋 Примеры сохраненных файлов:")
                for i, file_path in enumerate(saved_files[:5], 1):
                    self.logger.info(f"   {i}. {file_path.name}")
        else:
            self.logger.warning("\n❌ Файлы не найдены!")
            self.logger.warning("   Проверьте вручную один архив:")
            self.logger.warning("   - Откройте любой архив")
            self.logger.warning("   - Посмотрите точное название файла")
    
    def _create_summary_file(self):
        """Создает файл с кратким отчетом"""
        try:
            summary_file = self.output_folder / "!_ОТЧЕТ_О_НАЙДЕННЫХ_ФАЙЛАХ.txt"
            
            with open(summary_file, 'w', encoding='utf-8') as f:
                f.write("="*60 + "\n")
                f.write("ОТЧЕТ О ПОИСКЕ ФАЙЛОВ СЧА Фонд_ПДС\n")
                f.write(f"Дата поиска: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n")
                f.write("="*60 + "\n\n")
                
                f.write(f"Всего найдено файлов: {self.stats['files_found']}\n\n")
                f.write("Список найденных файлов:\n")
                f.write("-"*40 + "\n")
                
                saved_files = list(self.output_folder.glob("[*]*.xls"))
                saved_files.sort()
                
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
    output_path_obj = Path(output_path)
    
    if not search_path_obj.exists():
        print("\n❌ ОШИБКА: Не удалось подключиться к исходной папке!")
        print(f"   Путь: {search_path}")
        print("\nВозможные решения:")
        print("1. Проверьте подключение к VPN")
        print("2. Откройте папку в проводнике чтобы убедиться в доступе")
        print("3. Запустите скрипт от имени другого пользователя")
        input("\nНажмите Enter для выхода...")
        return
    
    # Создаем и запускаем поисковик
    finder = SCAFileFinder(search_path, output_path)
    files_found = finder.find_files()
    
    print("\n" + "="*80)
    if files_found:
        print(f"✅ РАБОТА ЗАВЕРШЕНА. Найдено файлов: {finder.stats['files_found']}")
    else:
        print("❌ РАБОТА ЗАВЕРШЕНА. Файлы не найдены.")
    print("="*80)
    print(f"📁 Папка для сохранения: {output_path}")
    print("\nЛог работы сохранен в файл: sca_finder.log")
    
    input("\nНажмите Enter для выхода...")

if __name__ == "__main__":
    main()
