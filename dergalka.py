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
