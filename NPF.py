import pandas as pd
import tkinter as tk
from tkinter import messagebox
import os
from datetime import datetime

# ФИКСИРОВАННЫЙ ПУТЬ К ФАЙЛУ 1
FILE1_PATH = r"C:\Users\ytggf\OneDrive\Документы\renlife\SFTTest\Депозит тест.xlsx"

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

def on_deposit_count_change(*args):
    """Обработчик изменения количества депозитов"""
    count = deposit_count.get()
    
    # Очищаем старые поля
    for widget in deposit_fields_frame.winfo_children():
        widget.destroy()
    
    if count == "1":
        # Поля для одного депозита (старая логика)
        tk.Label(deposit_fields_frame, text="Сумма платежей сегодня:").pack(anchor=tk.W)
        global entry_payment
        entry_payment = tk.Entry(deposit_fields_frame, width=30)
        entry_payment.pack(fill=tk.X, pady=(5, 10))
        
        tk.Label(deposit_fields_frame, text="Ставка:").pack(anchor=tk.W)
        global entry_rate
        entry_rate = tk.Entry(deposit_fields_frame, width=30)
        entry_rate.pack(fill=tk.X, pady=(5, 10))
        
        tk.Label(deposit_fields_frame, text="Срок до (дата):").pack(anchor=tk.W)
        tk.Label(deposit_fields_frame, text="Формат: ДД.ММ.ГГГГ (например: 31.12.2024)", 
                 fg="gray", font=("Arial", 8)).pack(anchor=tk.W)
        global entry_date
        entry_date = tk.Entry(deposit_fields_frame, width=30)
        entry_date.pack(fill=tk.X, pady=(5, 10))
        
    else:
        # Поля для нескольких депозитов (сумма, ставка и срок для каждого)
        count_num = int(count)
        tk.Label(deposit_fields_frame, text=f"Данные для {count_num} депозитов:", 
                 font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 10))
        
        # Показываем доступный остаток
        if hasattr(root, 'initial_balance'):
            available_balance = root.initial_balance - 150000
            tk.Label(deposit_fields_frame, 
                    text=f"Доступно для размещения: {format_number(available_balance)}",
                    fg="green", font=("Arial", 9, "bold")).pack(anchor=tk.W, pady=(0, 10))
        
        # Создаем поля для каждого депозита
        global deposit_entries
        deposit_entries = []
        
        for i in range(count_num):
            dep_frame = tk.Frame(deposit_fields_frame, relief=tk.GROOVE, bd=1, padx=10, pady=8)
            dep_frame.pack(fill=tk.X, pady=3)
            
            tk.Label(dep_frame, text=f"Депозит {i+1}", font=("Arial", 9, "bold")).pack(anchor=tk.W)
            
            # Сумма депозита
            tk.Label(dep_frame, text="Сумма:").pack(anchor=tk.W)
            entry_amount_dep = tk.Entry(dep_frame, width=25)
            entry_amount_dep.pack(fill=tk.X, pady=(2, 5))
            
            # Ставка
            tk.Label(dep_frame, text="Ставка:").pack(anchor=tk.W)
            entry_rate_dep = tk.Entry(dep_frame, width=25)
            entry_rate_dep.pack(fill=tk.X, pady=(2, 5))
            
            # Дата
            tk.Label(dep_frame, text="Срок до (дата):").pack(anchor=tk.W)
            tk.Label(dep_frame, text="Формат: ДД.ММ.ГГГГ", fg="gray", font=("Arial", 7)).pack(anchor=tk.W)
            entry_date_dep = tk.Entry(dep_frame, width=25)
            entry_date_dep.pack(fill=tk.X, pady=(2, 5))
            
            deposit_entries.append({
                'amount': entry_amount_dep,
                'rate': entry_rate_dep,
                'date': entry_date_dep
            })

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
            root.initial_balance = initial_balance  # Сохраняем для использования
            
            # Показываем найденный баланс
            label_balance.config(text=f"Найденный баланс: {format_number(initial_balance)}")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка чтения файла 1: {str(e)}")
            return
        
        count = deposit_count.get()
        
        if count == "1":
            # Обработка одного депозита (старая логика)
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
            
            # Выполняем расчеты
            final_amount = initial_balance - payment_amount - 150000
            rounded_amount = round(final_amount / 1000) * 1000  # Округление до тысяч
            
            # Проверяем, что сумма не отрицательная
            if rounded_amount < 0:
                messagebox.showerror("Ошибка", "Сумма к размещению не может быть отрицательной!")
                return
            
            # Форматируем сообщение
            days_info = calculate_days_until(target_date)
            
            message = f"""Добрый день!
Прошу подписать заявление на размещение депозитов в ГПБ Бизнес Онлайн

Сумма: {format_number(rounded_amount)}
Срок: до {target_date} ({days_info} дней)
Ставка: {interest_rate}"""
            
        else:
            # Обработка нескольких депозитов
            try:
                deposits_data = []
                total_amount = 0
                
                for i, deposit in enumerate(deposit_entries):
                    amount_str = deposit['amount'].get().strip()
                    rate_str = deposit['rate'].get().strip()
                    date_str = deposit['date'].get().strip()
                    
                    if not amount_str or not rate_str or not date_str:
                        messagebox.showerror("Ошибка", f"Заполните все поля для депозита {i+1}")
                        return
                    
                    amount = parse_number(amount_str)
                    total_amount += amount
                    days_info = calculate_days_until(date_str)
                    
                    deposits_data.append({
                        'amount': amount,
                        'rate': rate_str,
                        'date': date_str,
                        'days': days_info
                    })
                    
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка в данных: {str(e)}")
                return
            
            # Проверяем, что общая сумма не превышает доступный остаток
            available_balance = initial_balance - 10000
            if total_amount > available_balance:
                messagebox.showerror("Ошибка", 
                    f"Общая сумма депозитов ({format_number(total_amount)}) превышает доступный остаток!\n\n"
                    f"Доступно: {format_number(available_balance)}\n"
                    f"Превышение: {format_number(total_amount - available_balance)}")
                return
            
            # Форматируем сообщение для нескольких депозитов
            message = "Добрый день!\n"
            message += "Прошу подписать заявление на размещение депозитов в ГПБ Бизнес Онлайн\n\n"
            
            message += f"Общая сумма: {format_number(total_amount)}\n\n"
            
            for i, deposit in enumerate(deposits_data):
                message += f"Депозит {i+1}:\n"
                message += f"Сумма: {format_number(deposit['amount'])}\n"
                message += f"Ставка: {deposit['rate']}\n"
                message += f"Срок: до {deposit['date']} ({deposit['days']} дней)\n\n"

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
root.geometry("550x650")

# Информация о файле
frame_info = tk.Frame(root)
frame_info.pack(pady=10, padx=20, fill=tk.X)

tk.Label(frame_info, text=f"Файл баланса: {FILE1_PATH}", 
         wraplength=500, justify=tk.LEFT, fg="blue").pack(anchor=tk.W)

label_balance = tk.Label(frame_info, text="Баланс: не загружен", fg="red")
label_balance.pack(anchor=tk.W, pady=(5, 0))

# Выбор количества депозитов
frame_deposit_count = tk.Frame(root)
frame_deposit_count.pack(pady=10, padx=20, fill=tk.X)

tk.Label(frame_deposit_count, text="Количество депозитов:").pack(anchor=tk.W)
deposit_count = tk.StringVar(value="1")
deposit_combo = ttk.Combobox(frame_deposit_count, 
                           textvariable=deposit_count,
                           values=["1", "2", "3", "4", "5"],
                           state="readonly",
                           width=10)
deposit_combo.pack(anchor=tk.W, pady=(5, 10))
deposit_count.trace('w', on_deposit_count_change)

# Фрейм для полей ввода (будет меняться в зависимости от выбора)
deposit_fields_frame = tk.Frame(root)
deposit_fields_frame.pack(pady=10, padx=20, fill=tk.X)

# Инициализируем поля для одного депозита
on_deposit_count_change()

# Кнопка выполнения
btn_process = tk.Button(root, text="Сформировать сообщение", 
                       command=process_calculation,
                       bg="#4CAF50", fg="white", font=("Arial", 12, "bold"))
btn_process.pack(pady=15)

# Поле для результата
tk.Label(root, text="Результат:").pack(anchor=tk.W, padx=20)
result_text = tk.Text(root, height=12, width=65, font=("Arial", 10))
result_text.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

# Добавляем скроллбар для текстового поля
scrollbar = tk.Scrollbar(result_text)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
result_text.config(yscrollcommand=scrollbar.set)
scrollbar.config(command=result_text.yview)

# Инструкция
instruction = """
Инструкция:
1. Выберите количество депозитов
2. Заполните поля в зависимости от выбора:
   • 1 депозит: сумма платежей, ставка, дата
   • 2+ депозита: сумма, ставка и дата для КАЖДОГО депозита
3. Сумма всех депозитов не должна превышать доступный остаток
4. Нажмите "Сформировать сообщение"
5. Сообщение автоматически скопируется в буфер обмена
"""
tk.Label(root, text=instruction, justify=tk.LEFT, fg="gray", 
         font=("Arial", 8)).pack(anchor=tk.W, padx=20, pady=(0, 10))

# Запускаем приложение
root.mainloop()
