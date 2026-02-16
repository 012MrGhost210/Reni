import os
import sys
from pathlib import Path
from datetime import datetime
import time

class DiskAnalyzer:
    def __init__(self, root_path):
        self.root_path = Path(root_path)
        self.stats = {
            'total_size': 0,
            'total_files': 0,
            'total_folders': 0,
            'file_types': {},
            'largest_files': [],
            'largest_folders': []
        }
    
    def get_size_format(self, size, decimal_places=2):
        for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
            if size < 1024.0:
                return f"{size:.{decimal_places}f} {unit}"
            size /= 1024.0
        return f"{size:.{decimal_places}f} PB"
    
    def get_file_extension(self, filename):
        """Получает расширение файла"""
        ext = Path(filename).suffix.lower()
        return ext if ext else '(без расширения)'
    
    def analyze(self):
        """Основной метод анализа"""
        print(f"\n🔍 Анализ папки: {self.root_path.absolute()}")
        print("⏳ Это может занять некоторое время...\n")
        
        start_time = time.time()
        folder_sizes = {}
        
        for root, dirs, files in os.walk(self.root_path):
            current_folder = Path(root)
            folder_size = 0
            
            # Показываем прогресс
            if self.stats['total_folders'] % 100 == 0:
                print(f"\r📁 Обработано папок: {self.stats['total_folders']}", end='')
            
            self.stats['total_folders'] += 1
            
            for file in files:
                file_path = current_folder / file
                try:
                    if file_path.exists() and not file_path.is_symlink():
                        file_size = file_path.stat().st_size
                        folder_size += file_size
                        self.stats['total_size'] += file_size
                        self.stats['total_files'] += 1
                        
                        # Статистика по типам файлов
                        ext = self.get_file_extension(file)
                        self.stats['file_types'][ext] = self.stats['file_types'].get(ext, 0) + file_size
                        
                        # Сохраняем топ-10 самых больших файлов
                        if len(self.stats['largest_files']) < 10:
                            self.stats['largest_files'].append((file_path, file_size))
                            self.stats['largest_files'].sort(key=lambda x: x[1], reverse=True)
                        elif file_size > self.stats['largest_files'][-1][1]:
                            self.stats['largest_files'].append((file_path, file_size))
                            self.stats['largest_files'].sort(key=lambda x: x[1], reverse=True)
                            self.stats['largest_files'] = self.stats['largest_files'][:10]
                            
                except (PermissionError, OSError):
                    continue
            
            # Сохраняем размер папки
            if folder_size > 0:
                folder_sizes[str(current_folder.relative_to(self.root_path) or '.')] = folder_size
        
        print("\r" + " " * 50 + "\r", end='')  # Очищаем строку прогресса
        
        # Сортируем папки
        self.stats['largest_folders'] = sorted(folder_sizes.items(), 
                                              key=lambda x: x[1], reverse=True)[:20]
        
        elapsed_time = time.time() - start_time
        self.print_results(elapsed_time)
    
    def print_results(self, elapsed_time):
        """Выводит результаты анализа"""
        print("\n" + "="*70)
        print("📊 РЕЗУЛЬТАТЫ АНАЛИЗА")
        print("="*70)
        
        # Общая статистика
        print(f"\n📈 ОБЩАЯ СТАТИСТИКА:")
        print(f"   Общий размер: {self.get_size_format(self.stats['total_size'])}")
        print(f"   Всего файлов: {self.stats['total_files']:,}")
        print(f"   Всего папок: {self.stats['total_folders']:,}")
        print(f"   Время анализа: {elapsed_time:.2f} сек")
        
        # Топ-20 самых больших папок
        print(f"\n📁 ТОП-20 САМЫХ БОЛЬШИХ ПАПОК:")
        for i, (folder, size) in enumerate(self.stats['largest_folders'][:20], 1):
            if size > 1024*1024:  # Показываем только папки больше 1 MB
                print(f"   {i:2d}. {self.get_size_format(size):>10} : {folder}")
        
        # Топ-10 самых больших файлов
        print(f"\n📄 ТОП-10 САМЫХ БОЛЬШИХ ФАЙЛОВ:")
        for i, (file_path, size) in enumerate(self.stats['largest_files'], 1):
            print(f"   {i:2d}. {self.get_size_format(size):>10} : {file_path.name}")
        
        # Статистика по типам файлов
        print(f"\n🔤 СТАТИСТИКА ПО ТИПАМ ФАЙЛОВ:")
        sorted_types = sorted(self.stats['file_types'].items(), 
                            key=lambda x: x[1], reverse=True)[:15]
        for ext, size in sorted_types:
            if size > 1024*1024:  # Показываем только типы больше 1 MB
                percentage = (size / self.stats['total_size']) * 100
                print(f"   {ext:15} : {self.get_size_format(size):>10} ({percentage:.1f}%)")

def main():
    if len(sys.argv) > 1:
        folder = sys.argv[1]
    else:
        folder = input("Введите путь к папке для анализа: ").strip()
    
    if not os.path.exists(folder):
        print(f"❌ Ошибка: Папка {folder} не существует!")
        return
    
    if not os.path.isdir(folder):
        print(f"❌ Ошибка: {folder} - это не папка!")
        return
    
    analyzer = DiskAnalyzer(folder)
    analyzer.analyze()

if __name__ == "__main__":
    main()
