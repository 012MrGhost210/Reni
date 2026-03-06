import os

# Путь к stop.txt (должен совпадать с основным скриптом)
STOP_FILE = r"D:\мои_письма\stop.txt"

def remove_stop_file():
    """Удаляет файл stop.txt"""
    
    if os.path.exists(STOP_FILE):
        try:
            os.remove(STOP_FILE)
            print(f"✅ Файл {STOP_FILE} успешно удален")
            print("Теперь можно снова запустить отправку писем")
        except Exception as e:
            print(f"❌ Ошибка при удалении: {e}")
    else:
        print(f"❌ Файл {STOP_FILE} не найден")

if __name__ == "__main__":
    remove_stop_file()
Поиск .msg файлов в \\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\Test\Автопочта...
Ошибка: (-2147221005, 'Недопустимая строка с указанием класса', None, None)



