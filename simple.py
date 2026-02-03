import shutil
import os

# Укажите пути здесь
source_file = "C:/путь/к/исходному/файлу.txt"  # Путь к исходному файлу
destination_folder = "C:/путь/к/папке/назначения/"  # Путь к папке назначения

# Проверяем, существует ли исходный файл
if os.path.exists(source_file):
    # Получаем имя файла из исходного пути
    filename = os.path.basename(source_file)
    
    # Создаем полный путь для файла назначения
    destination_file = os.path.join(destination_folder, filename)
    
    # Копируем файл
    shutil.copy(source_file, destination_file)
    print(f"Файл успешно скопирован в: {destination_file}")
else:
    print(f"Исходный файл не найден: {source_file}")
