import os
from pathlib import Path

def get_size(size):
    for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
        if size < 1024:
            return f"{size:.1f} {unit}"
        size /= 1024

# ===== ВСТАВЬ СВОЙ ПУТЬ СЮДА =====
folder_path = r"C:\Users\Имя\Downloads"  # <--- измени это
# =================================

target = Path(folder_path)

if not target.exists():
    print("Папка не найдена!")
    exit()

print(f"\nСодержимое папки {target.name}:")
print("-" * 40)

total = 0
items = []

for item in target.iterdir():
    if item.is_dir():
        # быстрый подсчет размера папки
        size = sum(f.stat().st_size for f in item.glob('**/*') if f.is_file())
        items.append((item.name + "/", size))
    else:
        items.append((item.name, item.stat().st_size))
    total += items[-1][1]

# сортируем от больших к маленьким
for name, size in sorted(items, key=lambda x: x[1], reverse=True)[:20]:
    if size > 1024 * 1024:  # показываем только больше 1MB
        print(f"{get_size(size):>8} : {name}")

print("-" * 40)
print(f"Всего: {get_size(total)}")
