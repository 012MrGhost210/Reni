import pandas as pd
import tkinter as tk
from tkinter import messagebox
import os
from datetime import datetime

# ФИКСИРОВАННЫЙ ПУТЬ К ФАЙЛУ 1
FILE1_PATH = r"M:\Финансовый департамент\Treasury\Базы данных(автоматизация)\DI_DATABASE\Депозит.xlsx"  # ЗАМЕНИТЕ НА ВАШ ПУТЬ

def parse_number(text):
    """Преобразует строку с числами в формате '12 122 121,31' в float"""
    return float(text.replace(' ', '').replace(',', '.'))

def format_number(number):
    """Форматирует число в строку с пробелами тысяч и запятой для десятичных"""
    return f"{number:,.2f}".replace(',', ' ').replace('.', ',')

def calculate_days_until(target_date_str):
    """Вычисляет количество дней от сегодняшней даты до целевой даты"""
    today = datetime.now().date()
    
    # Парсим дату из разных возможных форматов
    for fmt in ['%d.%m.%Y', '%d/%m/%Y', '%Y-%m-%d', '%d.%m.%y']:
        try:
            target_date = datetime.strptime(target_date_str, fmt).date()
            days = (target_date - today).days
            return max(days, 1)  # Минимум 1 день
        except ValueError:
            continue
    
    # Если не удалось распарсить, возвращаем исходную строку
    return target_date_str

def process_calculation():
    """Основная функция обработки"""
    try:
        # Проверяем существование файла 1
        if not os.path.exists(FILE1_PATH):
            messagebox.showerror("Ошибка", f"Файл не найден по пути:\n{FILE1_PATH}")
            return
        
        # 1. Читаем первый файл и находим последнее число в столбце G
        try:
            df1 = pd.read_excel(FILE1_PATH)
            if df1.empty or len(df1.columns) < 7:
                messagebox.showerror("Ошибка", "Файл 1 не содержит столбец G или пуст")
                return
            
            # Ищем последнюю непустую ячейку в столбце G (индекс 6)
            column_g = df1.iloc[:, 6].dropna()
            if column_g.empty:
                messagebox.showerror("Ошибка", "В столбце G файла 1 нет данных")
                return
            
            last_value_str = str(column_g.iloc[-1])
            initial_balance = parse_number(last_value_str)
            
            # Показываем найденный баланс
            label_balance.config(text=f"Найденный баланс: {format_number(initial_balance)}")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка чтения файла 1: {str(e)}")
            return
        
        # 2. Получаем данные из полей ввода
        try:
            payment_amount = parse_number(entry_payment.get())
            interest_rate = entry_rate.get().strip()
            target_date = entry_date.get().strip()
            
            if not payment_amount or not interest_rate or not target_date:
                messagebox.showerror("Ошибка", "Заполните все поля")
                return
                
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка в данных: {str(e)}")
            return
        
        # 3. Выполняем расчеты
        final_amount = initial_balance - payment_amount - 160000
        rounded_amount = round(final_amount / 1000) * 1000  # Округление до тысяч
        
        # 4. Форматируем сообщение
        days_info = calculate_days_until(target_date)
        
        message = f"""Добрый день!
Прошу подписать заявление на размещение депозитов в ГПБ Бизнес Онлайн

Сумма: {format_number(rounded_amount)}
Срок: до {target_date} ({days_info} дней)
Ставка: {interest_rate}"""

        # Показываем результат
        result_text.delete(1.0, tk.END)
        result_text.insert(1.0, message)
        
        # Копируем в буфер обмена
        root.clipboard_clear()
        root.clipboard_append(message)
        messagebox.showinfo("Готово", "Сообщение сформировано и скопировано в буфер обмена!")
        
    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка: {str(e)}")

# Создаем графический интерфейс
root = tk.Tk()
root.title("Расчет депозитов ГПБ")
root.geometry("500x550")

# Информация о файле
frame_info = tk.Frame(root)
frame_info.pack(pady=10, padx=20, fill=tk.X)

tk.Label(frame_info, text=f"Файл баланса: {FILE1_PATH}", 
         wraplength=450, justify=tk.LEFT, fg="blue").pack(anchor=tk.W)

label_balance = tk.Label(frame_info, text="Баланс: не загружен", fg="red")
label_balance.pack(anchor=tk.W, pady=(5, 0))

# Поля для ввода параметров
frame_inputs = tk.Frame(root)
frame_inputs.pack(pady=15, padx=20, fill=tk.X)

# Сумма платежей
tk.Label(frame_inputs, text="Сумма платежей сегодня:").pack(anchor=tk.W)
entry_payment = tk.Entry(frame_inputs, width=30)
entry_payment.pack(fill=tk.X, pady=(5, 10))

# Ставка
tk.Label(frame_inputs, text="Ставка:").pack(anchor=tk.W)
entry_rate = tk.Entry(frame_inputs, width=30)
entry_rate.pack(fill=tk.X, pady=(5, 10))

# Дата
tk.Label(frame_inputs, text="Срок до (дата):").pack(anchor=tk.W)
tk.Label(frame_inputs, text="Формат: ДД.ММ.ГГГГ (например: 31.12.2024)", 
         fg="gray", font=("Arial", 8)).pack(anchor=tk.W)
entry_date = tk.Entry(frame_inputs, width=30)
entry_date.pack(fill=tk.X, pady=(5, 10))

# Кнопка выполнения
btn_process = tk.Button(root, text="Сформировать сообщение", 
                       command=process_calculation,
                       bg="#4CAF50", fg="white", font=("Arial", 12, "bold"))
btn_process.pack(pady=15)

# Поле для результата
tk.Label(root, text="Результат:").pack(anchor=tk.W, padx=20)
result_text = tk.Text(root, height=10, width=60, font=("Arial", 10))
result_text.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

# Добавляем скроллбар для текстового поля
scrollbar = tk.Scrollbar(result_text)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
result_text.config(yscrollcommand=scrollbar.set)
scrollbar.config(command=result_text.yview)

# Инструкция
instruction = """
Инструкция:
1. Убедитесь, что файл баланса доступен по указанному пути
2. Введите данные в поля:
   • Сумма платежей сегодня (например: 22 000)
   • Ставка (например: 8,5%)
   • Дата окончания (например: 31.12.2024)
3. Нажмите "Сформировать сообщение"
4. Сообщение автоматически скопируется в буфер обмена
"""
tk.Label(root, text=instruction, justify=tk.LEFT, fg="gray", 
         font=("Arial", 8)).pack(anchor=tk.W, padx=20, pady=(0, 10))

# Запускаем приложение
if __name__ == "__main__":
    # Показываем предупреждение о необходимости настройки пути
    if FILE1_PATH == r"C:\путь\к\вашему\файлу\баланс.xlsx":
        messagebox.showwarning("Настройка", 
                              "Перед использованием замените FILE1_PATH в коде на актуальный путь к вашему файлу!")
    
    root.mainloop()  
