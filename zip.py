import os
import zipfile
import re
from pathlib import Path
import shutil
import logging
from datetime import datetime

class SCAFileFinder:
    def __init__(self, network_path, output_folder=None):
        self.network_path = Path(network_path)
        
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
        
        # Папка для сохранения
        if output_folder:
            self.output_folder = Path(output_folder)
        else:
            self.output_folder = Path.home() / "Desktop" / f"СЧА_файлы_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        
        self.output_folder.mkdir(exist_ok=True, parents=True)
        
        # Статистика
        self.stats = {
            'folders_checked': 0,
            'archives_checked': 0,
            'files_found': 0,
            'errors': 0
        }
        
    def check_network_connection(self):
        """Проверка доступности сетевого пути"""
        try:
            # Пробуем создать временный файл для проверки доступа
            test_file = self.network_path / 'test_write.tmp'
            test_file.touch()
            test_file.unlink()
            return True
        except:
            return False
    
    def find_files(self):
        """Основной метод поиска"""
        
        self.logger.info(f"🔍 Начинаем поиск в: {self.network_path}")
        self.logger.info(f"📁 Файлы будут сохранены в: {self.output_folder}")
        
        # Проверяем доступность сетевого пути
        if not self.network_path.exists():
            self.logger.error(f"❌ Путь не существует: {self.network_path}")
            self.logger.error("   Проверьте:")
            self.logger.error("   1. Подключение к VPN/сети")
            self.logger.error("   2. Права доступа к папке")
            return False
        
        # Паттерны для поиска
        date_patterns = [
            r'\d{2}\.\d{2}\.\d{4}',  # 29.12.2025
            r'\d{4}-\d{2}-\d{2}',    # 2025-12-29
            r'\d{2}-\d{2}-\d{4}',    # 29-12-2025
        ]
        
        file_variants = [
            "СЧА Фонд_ПДС.xls",
            "СЧА Фонд_ПДС.xlsx",
            "СЧА_Фонд_ПДС.xls",
            "СЧА_Фонд_ПДС.xlsx"
        ]
        
        # Проходим по папкам с датами
        date_folders = list(self.network_path.glob("*_*_*"))
        self.logger.info(f"Найдено папок с датами: {len(date_folders)}")
        
        for date_folder in date_folders:
            if not date_folder.is_dir():
                continue
                
            self.stats['folders_checked'] += 1
            self.logger.info(f"\n📂 Проверяем папку: {date_folder.name}")
            
            # Путь к документам гаранта
            guarant_folder = date_folder / "Документы от Гаранта СД НТД"
            
            if not guarant_folder.exists():
                self.logger.warning(f"  ⚠️  Папка 'Документы от Гаранта СД НТД' не найдена")
                continue
            
            # Ищем ZIP архивы
            zip_files = list(guarant_folder.glob("Отчеты_*.zip"))
            self.logger.info(f"  Найдено архивов: {len(zip_files)}")
            
            for zip_path in zip_files:
                self.stats['archives_checked'] += 1
                self.logger.info(f"  📦 Проверяем: {zip_path.name}")
                
                try:
                    with zipfile.ZipFile(zip_path, 'r') as zf:
                        # Проверяем все комбинации
                        for file_in_zip in zf.namelist():
                            file_name = Path(file_in_zip).name
                            
                            for date_pattern in date_patterns:
                                for file_variant in file_variants:
                                    pattern = f"{date_pattern}_{file_variant}"
                                    
                                    if re.match(pattern, file_name):
                                        self.stats['files_found'] += 1
                                        self.logger.info(f"     ✅ НАЙДЕН: {file_name}")
                                        
                                        # Сохраняем файл
                                        self._save_file(zf, file_in_zip, date_folder.name, file_name)
                                        
                except zipfile.BadZipFile:
                    self.stats['errors'] += 1
                    self.logger.error(f"     ❌ Испорченный ZIP архив: {zip_path.name}")
                except Exception as e:
                    self.stats['errors'] += 1
                    self.logger.error(f"     ❌ Ошибка: {e}")
        
        # Выводим статистику
        self._print_statistics()
        return self.stats['files_found'] > 0
    
    def _save_file(self, zip_file, file_in_zip, date_folder_name, original_filename):
        """Сохранение найденного файла"""
        try:
            # Создаем папку для этой даты
            save_dir = self.output_folder / date_folder_name
            save_dir.mkdir(exist_ok=True)
            
            # Извлекаем файл
            zip_file.extract(file_in_zip, save_dir)
            
            # Перемещаем в корень папки если был в подпапке
            extracted_path = save_dir / file_in_zip
            final_path = save_dir / original_filename
            
            if extracted_path != final_path and extracted_path.exists():
                shutil.move(extracted_path, final_path)
                
                # Удаляем пустые папки
                for parent in extracted_path.parents:
                    if parent != save_dir:
                        try:
                            parent.rmdir()
                        except:
                            pass
            
            self.logger.info(f"     💾 Сохранен в: {final_path}")
            
        except Exception as e:
            self.logger.error(f"     ❌ Ошибка сохранения: {e}")
    
    def _print_statistics(self):
        """Вывод статистики"""
        self.logger.info("\n" + "="*60)
        self.logger.info("📊 СТАТИСТИКА:")
        self.logger.info(f"   Проверено папок: {self.stats['folders_checked']}")
        self.logger.info(f"   Проверено архивов: {self.stats['archives_checked']}")
        self.logger.info(f"   Найдено файлов: {self.stats['files_found']}")
        self.logger.info(f"   Ошибок: {self.stats['errors']}")
        
        if self.stats['files_found'] > 0:
            self.logger.info(f"\n✅ Все файлы сохранены в: {self.output_folder}")
        else:
            self.logger.warning("\n❌ Файлы не найдены!")
            self.logger.warning("   Возможные причины:")
            self.logger.warning("   - Неправильная структура папок")
            self.logger.warning("   - Файлы имеют другое название")
            self.logger.warning("   - Нет доступа к архивам")

def main():
    # Путь к сетевой папке
    network_path = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\01.Перечень имущества Фонда (СД)"
    
    # Папка для сохранения на рабочем столе
    save_folder = Path.home() / "Desktop" / "СЧА_файлы_от_гаранта"
    
    print("="*60)
    print("🚀 ПОИСК ФАЙЛОВ СЧА Фонд_ПДС")
    print("="*60)
    
    # Создаем и запускаем поисковик
    finder = SCAFileFinder(network_path, save_folder)
    finder.find_files()
    
    print("\n" + "="*60)
    input("Нажмите Enter для выхода...")

if __name__ == "__main__":
    main()
