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
import requests
response = requests.get('http://<IP-вашего-локального-ПК>:5000/run')
print(response.json())



