import subprocess
import sys
import os

def run_script(script_path):
    """
    Запускает Python скрипт из указанной директории
    """
    try:
        # Получаем директорию и имя файла
        script_dir = os.path.dirname(script_path)
        script_name = os.path.basename(script_path)
        
        # Запускаем скрипт
        result = subprocess.run(
            [sys.executable, script_name],
            cwd=script_dir,
            capture_output=True,
            text=True,
            timeout=30  # таймаут 30 секунд
        )
        
        # Проверяем результат
        if result.returncode == 0:
            print("✅ Скрипт выполнен успешно!")
            if result.stdout:
                print("Вывод скрипта:")
                print(result.stdout)
            return True
        else:
            print("❌ Скрипт завершился с ошибкой!")
            print(f"Код ошибки: {result.returncode}")
            if result.stderr:
                print("Ошибки:")
                print(result.stderr)
            return False
            
    except subprocess.TimeoutExpired:
        print("❌ Скрипт превысил время выполнения!")
        return False
    except FileNotFoundError:
        print(f"❌ Файл не найден: {script_path}")
        return False
    except Exception as e:
        print(f"❌ Непредвиденная ошибка: {e}")
        return False

# Использование
script_to_run = r"C:\путь\к\вашему\скрипту\script.py"
run_script(script_to_run)


Ошибка: 'charmap' codec can't encode character '\u2717' in position 0: character maps to <undefined>




import win32com.client
import pythoncom
import os

# Конфигурация
DRAFTS_FOLDER = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\Test\Автопочта"  # Папка с .msg файлами
STOP_FILE = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\Test\Автопочта\stop.txt"  # Файл-флаг

def send_msg_files():
    """Отправляет .msg файлы из указанной папки"""
    
    # Проверяем stop.txt
    if os.path.exists(STOP_FILE):
        print(f"Найден {STOP_FILE}, отправка отменена")
        return
    
    print(f"Поиск .msg файлов в {DRAFTS_FOLDER}...")
    
    try:
        # Подключаемся к Outlook
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        
        # Находим все .msg файлы
        msg_files = [f for f in os.listdir(DRAFTS_FOLDER) if f.lower().endswith('.msg')]
        
        if not msg_files:
            print("Нет .msg файлов для отправки")
            # Создаем stop.txt даже если файлов нет
            with open(STOP_FILE, 'w') as f:
                f.write("No files found")
            return
        
        print(f"Найдено файлов: {len(msg_files)}")
        
        sent_count = 0
        
        # Отправляем каждый файл
        for filename in msg_files:
            filepath = os.path.join(DRAFTS_FOLDER, filename)
            try:
                # Открываем .msg файл в Outlook
                msg = outlook.CreateItemFromTemplate(filepath)
                
                # Отправляем
                msg.Send()
                sent_count += 1
                print(f"✓ Отправлено: {filename}")
                
            except Exception as e:
                print(f"✗ Ошибка с {filename}: {e}")
        
        print(f"\nВсего отправлено: {sent_count}")
        
        # Создаем stop.txt
        with open(STOP_FILE, 'w') as f:
            f.write(f"OK\n{sent_count}")
        
        print(f"Создан {STOP_FILE}")
        
    except Exception as e:
        print(f"Ошибка: {e}")
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    send_msg_files()
