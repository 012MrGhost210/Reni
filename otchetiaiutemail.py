import os
import shutil
import tempfile
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client as win32

def get_date_folder(date_obj):
    """Возвращает строку с датой в формате ДД.ММ.ГГ"""
    return date_obj.strftime("%d.%m.%y")

def find_source_folder(base_path, target_date):
    """Ищет папку с отчетами за указанную дату"""
    date_str = get_date_folder(target_date)
    # Формируем имя искомой папки: M:\Финансы\Спец\26.04.15\отч_брокера+
    folder_path = os.path.join(base_path, date_str, "отч_брокера+")
    
    if os.path.exists(folder_path) and os.path.isdir(folder_path):
        return folder_path
    return None

def get_default_dates(today):
    """Возвращает список рекомендуемых дат для выбора"""
    dates = []
    # Вчера
    dates.append(today - timedelta(days=1))
    
    # Если сегодня понедельник (0 = понедельник), добавляем пятницу
    if today.weekday() == 0:
        dates.append(today - timedelta(days=3))
    
    # Позавчера
    dates.append(today - timedelta(days=2))
    
    return dates

def send_email_with_attachments(subject, body, attachments, recipient):
    """Отправляет письмо через Outlook с вложениями"""
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.Body = body
        mail.To = recipient
        
        for attachment in attachments:
            mail.Attachments.Add(attachment)
        
        mail.Send()
        return True, "Письмо успешно отправлено!"
    except Exception as e:
        return False, f"Ошибка при отправке: {str(e)}"

def copy_files_to_temp(source_folder):
    """Копирует все файлы из папки во временную директорию"""
    temp_dir = tempfile.mkdtemp()
    copied_files = []
    
    try:
        for filename in os.listdir(source_folder):
            src = os.path.join(source_folder, filename)
            if os.path.isfile(src):
                dst = os.path.join(temp_dir, filename)
                shutil.copy2(src, dst)
                copied_files.append(dst)
        return temp_dir, copied_files
    except Exception as e:
        shutil.rmtree(temp_dir, ignore_errors=True)
        return None, []

class ReportSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Отправка отчетов брокера")
        self.root.geometry("650x500")
        
        # Базовый путь - ИСПРАВЛЕНО
        self.base_path = tk.StringVar(value=r"M:\Финансовый департамент\SpecDep")
        
        # Получатель
        self.recipient = tk.StringVar(value="Ulyana.Pankratova@renlife.com")
        
        # Тема письма
        self.subject = tk.StringVar(value="Отчеты брокера")
        
        # Текущая дата
        self.today = datetime.now().date()
        self.selected_folder = None
        
        self.create_widgets()
        self.update_date_list()
    
    def create_widgets(self):
        # Базовый путь
        tk.Label(self.root, text="Базовая папка M:\Финансовый департамент\SpecDep:").pack(pady=(10,0))
        path_frame = tk.Frame(self.root)
        path_frame.pack(fill=tk.X, padx=20, pady=5)
        tk.Entry(path_frame, textvariable=self.base_path, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(path_frame, text="Обзор...", command=self.browse_folder).pack(side=tk.RIGHT, padx=5)
        
        # Выбор даты
        tk.Label(self.root, text="Выберите дату отчета (в формате ДД.ММ.ГГ):").pack(pady=(10,0))
        self.date_combo = ttk.Combobox(self.root, state="readonly", width=25)
        self.date_combo.pack(pady=5)
        self.date_combo.bind('<<ComboboxSelected>>', self.on_date_selected)
        
        # Информация о найденной папке
        self.folder_info = tk.Label(self.root, text="Папка не выбрана", fg="gray", wraplength=600, justify="left")
        self.folder_info.pack(pady=5)
        
        # Получатель
        tk.Label(self.root, text="Получатель (email):").pack(pady=(10,0))
        tk.Entry(self.root, textvariable=self.recipient, width=50).pack(pady=5)
        
        # Тема письма
        tk.Label(self.root, text="Тема письма:").pack(pady=(10,0))
        tk.Entry(self.root, textvariable=self.subject, width=60).pack(pady=5)
        
        # Текст письма
        tk.Label(self.root, text="Текст письма:").pack(pady=(10,0))
        self.body_text = tk.Text(self.root, height=8, width=70)
        self.body_text.pack(pady=5, padx=20)
        self.body_text.insert("1.0", "Добрый день!\n\nНаправляю отчеты брокера за указанную дату.\n\nС уважением,\nОтдел отчетности")
        
        # Кнопка отправки
        self.send_btn = tk.Button(self.root, text="Отправить отчеты", command=self.send_reports, 
                                  bg="green", fg="white", font=("Arial", 10, "bold"))
        self.send_btn.pack(pady=20)
        
        # Статус
        self.status_label = tk.Label(self.root, text="Готов к работе", fg="blue")
        self.status_label.pack(pady=5)
    
    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.base_path.set(folder)
            self.update_date_list()
    
    def update_date_list(self):
        """Обновляет список доступных дат"""
        default_dates = get_default_dates(self.today)
        date_options = []
        for date in default_dates:
            date_options.append(f"{date.strftime('%d.%m.%y')} ({date.strftime('%d.%m.%Y')})")
        date_options.append("Другая дата...")
        
        self.date_combo['values'] = date_options
        if date_options:
            self.date_combo.current(0)
            self.on_date_selected()
    
    def on_date_selected(self, event=None):
        """Обработчик выбора даты"""
        selection = self.date_combo.get()
        if selection == "Другая дата...":
            self.select_custom_date()
        else:
            # Извлекаем дату из строки (формат ДД.ММ.ГГ)
            date_str = selection.split(" ")[0]
            try:
                selected_date = datetime.strptime(date_str, "%d.%m.%y").date()
                self.check_folder_for_date(selected_date)
            except Exception as e:
                self.folder_info.config(text=f"Ошибка при разборе даты: {e}", fg="red")
    
    def select_custom_date(self):
        """Диалог выбора произвольной даты"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Выбор даты")
        dialog.geometry("350x180")
        dialog.grab_set()
        
        tk.Label(dialog, text="Введите дату в формате ДД.ММ.ГГ или ДД.ММ.ГГГГ:").pack(pady=10)
        date_entry = tk.Entry(dialog, width=15)
        date_entry.pack(pady=5)
        
        def confirm():
            try:
                date_str = date_entry.get().strip()
                # Пробуем разные форматы
                for fmt in ["%d.%m.%y", "%d.%m.%Y"]:
                    try:
                        selected_date = datetime.strptime(date_str, fmt).date()
                        dialog.destroy()
                        self.check_folder_for_date(selected_date)
                        # Обновляем комбобокс
                        self.date_combo.set(f"{selected_date.strftime('%d.%m.%y')} ({selected_date.strftime('%d.%m.%Y')})")
                        return
                    except:
                        continue
                raise ValueError("Неверный формат")
            except ValueError:
                messagebox.showerror("Ошибка", "Неверный формат даты. Используйте ДД.ММ.ГГ или ДД.ММ.ГГГГ")
        
        tk.Button(dialog, text="Подтвердить", command=confirm).pack(pady=10)
    
    def check_folder_for_date(self, selected_date):
        """Проверяет наличие папки с отчетами"""
        self.selected_folder = find_source_folder(self.base_path.get(), selected_date)
        
        if self.selected_folder:
            # Проверяем, что в папке есть файлы
            files = [f for f in os.listdir(self.selected_folder) if os.path.isfile(os.path.join(self.selected_folder, f))]
            if files:
                self.folder_info.config(text=f"✓ Найдена папка: {self.selected_folder}\n✓ Файлов для отправки: {len(files)}", fg="green")
                self.status_label.config(text=f"Готов к отправке отчета за {selected_date.strftime('%d.%m.%Y')}", fg="green")
            else:
                self.folder_info.config(text=f"⚠ Папка найдена, но она пуста: {self.selected_folder}", fg="orange")
                self.selected_folder = None
        else:
            expected_path = os.path.join(self.base_path.get(), get_date_folder(selected_date), "отч_брокера+")
            self.folder_info.config(text=f"✗ Папка НЕ найдена!\nОжидалось: {expected_path}\n(Обратите внимание на наличие знака '+' в конце 'отч_брокера+')", fg="red")
            self.status_label.config(text="Папка не найдена", fg="red")
            self.selected_folder = None
    
    def send_reports(self):
        """Основная логика отправки"""
        if not self.selected_folder:
            messagebox.showwarning("Предупреждение", "Сначала выберите дату, для которой существует папка 'отч_брокера+'")
            return
        
        # Проверяем получателя
        recipient = self.recipient.get().strip()
        if not recipient:
            messagebox.showwarning("Предупреждение", "Укажите email получателя")
            return
        
        # Получаем список файлов
        files = [f for f in os.listdir(self.selected_folder) if os.path.isfile(os.path.join(self.selected_folder, f))]
        
        if not files:
            messagebox.showwarning("Предупреждение", "В папке нет файлов для отправки")
            return
        
        files_list = "\n".join(f"  • {f}" for f in files)
        
        confirm_msg = f"📁 Папка: {self.selected_folder}\n\n📄 Файлы ({len(files)} шт.):\n{files_list}\n\n📧 Получатель: {recipient}\n\n✉️ Тема: {self.subject.get()}\n\nПродолжить отправку?"
        
        if not messagebox.askyesno("Подтверждение отправки", confirm_msg):
            return
        
        # Обновляем статус
        self.status_label.config(text="Копирование файлов и отправка...", fg="orange")
        self.root.update()
        
        # Копируем файлы во временную папку
        temp_dir, temp_files = copy_files_to_temp(self.selected_folder)
        
        if not temp_files:
            messagebox.showerror("Ошибка", "Не удалось скопировать файлы")
            self.status_label.config(text="Ошибка копирования", fg="red")
            return
        
        # Отправляем письмо
        success, message = send_email_with_attachments(
            subject=self.subject.get(),
            body=self.body_text.get("1.0", tk.END).strip(),
            attachments=temp_files,
            recipient=recipient
        )
        
        # Удаляем временную папку
        shutil.rmtree(temp_dir, ignore_errors=True)
        
        if success:
            messagebox.showinfo("Успех", f"Письмо отправлено!\n\nФайлов: {len(temp_files)}\nПолучатель: {recipient}")
            self.status_label.config(text=f"✓ Отправлено успешно! {datetime.now().strftime('%H:%M:%S')}", fg="green")
        else:
            messagebox.showerror("Ошибка", message)
            self.status_label.config(text="Ошибка отправки", fg="red")

def main():
    root = tk.Tk()
    app = ReportSenderApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
