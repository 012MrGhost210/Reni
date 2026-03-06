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
Traceback (most recent call last):
  File "C:\Users\Ilya.Matveev2\Project\stop.py", line 12, in remove_stop_file
    print(f"\u2705 Файл {STOP_FILE} успешно удален")
  File "C:\Users\Ilya.Matveev2\AppData\Local\Programs\Python\Python310\lib\encodings\cp1251.py", line 19, in encode
    return codecs.charmap_encode(input,self.errors,encoding_table)[0]
UnicodeEncodeError: 'charmap' codec can't encode character '\u2705' in position 0: character maps to <undefined>

During handling of the above exception, another exception occurred:

Traceback (most recent call last):
  File "C:\Users\Ilya.Matveev2\Project\stop.py", line 20, in <module>
    remove_stop_file()
  File "C:\Users\Ilya.Matveev2\Project\stop.py", line 15, in remove_stop_file
    print(f"\u274c Ошибка при удалении: {e}")
  File "C:\Users\Ilya.Matveev2\AppData\Local\Programs\Python\Python310\lib\encodings\cp1251.py", line 19, in encode
    return codecs.charmap_encode(input,self.errors,encoding_table)[0]
UnicodeEncodeError: 'charmap' codec can't encode character '\u274c' in position 0: character maps to <undefined>
