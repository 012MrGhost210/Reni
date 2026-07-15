import win32com.client
import os
import re
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import pythoncom
from datetime import datetime
import sys

class EmailBatchSender:
    def __init__(self, root):
        self.root = root
        self.root.title("Массовая отправка писем с шаблоном")
        self.root.geometry("1000x700")
        
        # Инициализация Outlook
        try:
            pythoncom.CoInitialize()
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось подключиться к Outlook: {e}")
            sys.exit(1)
        
        # Данные
        self.draft_path = ""
        self.template_body = ""
        self.template_subject = ""
        self.placeholders = []
        self.email_items = []  # Список словарей с данными для каждого письма
        
        # Создаем интерфейс
        self.create_widgets()
        
    def create_widgets(self):
        # Основной контейнер с прокруткой
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Панель инструментов
        toolbar = ttk.Frame(main_frame)
        toolbar.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(toolbar, text="Загрузить шаблон", command=self.load_template).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Добавить получателя", command=self.add_recipient).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Удалить выделенного", command=self.delete_recipient).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Отправить все", command=self.send_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Открыть выбранное", command=self.open_selected).pack(side=tk.LEFT, padx=5)
        
        # Информация о шаблоне
        self.template_info = ttk.Label(main_frame, text="Шаблон не загружен", foreground="red")
        self.template_info.pack(fill=tk.X, pady=(0, 10))
        
        # Разделитель
        ttk.Separator(main_frame, orient='horizontal').pack(fill=tk.X, pady=5)
        
        # Создаем панель с двумя частями
        paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)
        
        # Левая панель - список получателей
        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=1)
        
        ttk.Label(left_frame, text="Список получателей:", font=('Arial', 10, 'bold')).pack(anchor=tk.W, pady=(0, 5))
        
        # Таблица получателей
        columns = ('#', 'To', 'CC', 'Тема')
        self.tree = ttk.Treeview(left_frame, columns=columns, show='headings', height=15)
        
        self.tree.heading('#', text='№')
        self.tree.heading('To', text='Кому')
        self.tree.heading('CC', text='Копия')
        self.tree.heading('Тема', text='Тема')
        
        self.tree.column('#', width=40)
        self.tree.column('To', width=200)
        self.tree.column('CC', width=150)
        self.tree.column('Тема', width=200)
        
        scroll_tree = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_tree.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_tree.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Привязываем событие выбора
        self.tree.bind('<<TreeviewSelect>>', self.on_recipient_select)
        
        # Правая панель - редактирование текста
        right_frame = ttk.Frame(paned)
        paned.add(right_frame, weight=2)
        
        ttk.Label(right_frame, text="Редактирование текста письма:", font=('Arial', 10, 'bold')).pack(anchor=tk.W, pady=(0, 5))
        
        # Информация о плейсхолдерах
        self.placeholder_info = ttk.Label(right_frame, text="Плейсхолдеры: {}")
        self.placeholder_info.pack(anchor=tk.W, pady=(0, 5))
        
        # Текстовое поле для редактирования
        self.text_editor = scrolledtext.ScrolledText(right_frame, wrap=tk.WORD, height=20)
        self.text_editor.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        
        # Кнопки для вставки плейсхолдеров
        placeholder_frame = ttk.Frame(right_frame)
        placeholder_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(placeholder_frame, text="Вставить плейсхолдер:").pack(side=tk.LEFT, padx=5)
        
        self.placeholder_var = tk.StringVar()
        self.placeholder_combo = ttk.Combobox(placeholder_frame, textvariable=self.placeholder_var, width=20)
        self.placeholder_combo.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(placeholder_frame, text="Вставить", command=self.insert_placeholder).pack(side=tk.LEFT, padx=5)
        
        # Кнопки управления текстом
        text_buttons_frame = ttk.Frame(right_frame)
        text_buttons_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(text_buttons_frame, text="Применить ко всем", command=self.apply_to_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(text_buttons_frame, text="Сбросить текст", command=self.reset_text).pack(side=tk.LEFT, padx=5)
        ttk.Button(text_buttons_frame, text="Предпросмотр", command=self.preview_email).pack(side=tk.LEFT, padx=5)
        
        # Статус бар
        self.status_bar = ttk.Label(self.root, text="Готов к работе", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
    def load_template(self):
        """Загрузка шаблона письма"""
        file_path = filedialog.askopenfilename(
            title="Выберите шаблон письма",
            filetypes=[("Outlook шаблоны", "*.msg *.oft"), ("Все файлы", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            # Загружаем шаблон
            mail = self.outlook.CreateItemFromTemplate(file_path)
            
            # Сохраняем данные шаблона
            self.draft_path = file_path
            self.template_body = mail.Body
            self.template_subject = mail.Subject
            
            # Находим все плейсхолдеры в формате {text}
            self.placeholders = re.findall(r'\{([^}]+)\}', self.template_body)
            self.placeholders = list(set(self.placeholders))  # Убираем дубликаты
            
            # Обновляем информацию
            self.template_info.config(
                text=f"Шаблон загружен: {os.path.basename(file_path)} | Плейсхолдеров: {len(self.placeholders)}",
                foreground="green"
            )
            
            # Обновляем комбобокс с плейсхолдерами
            self.placeholder_combo['values'] = [f"{{{p}}}" for p in self.placeholders]
            
            # Если есть плейсхолдеры, показываем информацию
            if self.placeholders:
                self.placeholder_info.config(text=f"Плейсхолдеры: {', '.join([f'{{{p}}}' for p in self.placeholders])}")
            
            # Очищаем список получателей
            self.email_items = []
            self.update_tree()
            
            # Автоматически добавляем первого получателя
            self.add_recipient()
            
            messagebox.showinfo("Успех", f"Шаблон загружен успешно!\nНайдено плейсхолдеров: {len(self.placeholders)}")
            self.status_bar.config(text=f"Загружен шаблон: {os.path.basename(file_path)}")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить шаблон: {e}")
            self.status_bar.config(text="Ошибка загрузки шаблона")
    
    def add_recipient(self):
        """Добавление нового получателя"""
        # Создаем диалог для ввода данных
        dialog = tk.Toplevel(self.root)
        dialog.title("Добавить получателя")
        dialog.geometry("400x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Email получателя (To):").pack(anchor=tk.W, padx=10, pady=(10, 5))
        to_entry = ttk.Entry(dialog, width=50)
        to_entry.pack(padx=10, pady=(0, 10))
        
        ttk.Label(dialog, text="Копия (CC):").pack(anchor=tk.W, padx=10, pady=(0, 5))
        cc_entry = ttk.Entry(dialog, width=50)
        cc_entry.pack(padx=10, pady=(0, 10))
        
        ttk.Label(dialog, text="Тема письма (оставьте пустым для использования из шаблона):").pack(anchor=tk.W, padx=10, pady=(0, 5))
        subject_entry = ttk.Entry(dialog, width=50)
        subject_entry.pack(padx=10, pady=(0, 10))
        
        def save_recipient():
            to = to_entry.get().strip()
            if not to:
                messagebox.showwarning("Предупреждение", "Email получателя обязателен!")
                return
            
            cc = cc_entry.get().strip()
            subject = subject_entry.get().strip() or self.template_subject
            
            # Создаем данные для письма
            email_data = {
                'to': to,
                'cc': cc,
                'subject': subject,
                'body': self.template_body  # Начинаем с шаблона
            }
            
            self.email_items.append(email_data)
            self.update_tree()
            
            # Если это первый получатель, автоматически выбираем его
            if len(self.email_items) == 1:
                self.tree.selection_set(self.tree.get_children()[0])
                self.on_recipient_select(None)
            
            dialog.destroy()
            self.status_bar.config(text=f"Добавлен получатель: {to}")
        
        ttk.Button(dialog, text="Добавить", command=save_recipient).pack(pady=10)
        
    def delete_recipient(self):
        """Удаление выбранного получателя"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите получателя для удаления")
            return
        
        if messagebox.askyesno("Подтверждение", "Удалить выбранного получателя?"):
            # Получаем индекс
            item = selection[0]
            index = self.tree.index(item)
            
            # Удаляем из списка
            del self.email_items[index]
            self.update_tree()
            self.status_bar.config(text="Получатель удален")
            
            # Очищаем редактор
            self.text_editor.delete('1.0', tk.END)
    
    def on_recipient_select(self, event):
        """Обработчик выбора получателя в списке"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        index = self.tree.index(item)
        
        if index < len(self.email_items):
            email_data = self.email_items[index]
            # Загружаем текст в редактор
            self.text_editor.delete('1.0', tk.END)
            self.text_editor.insert('1.0', email_data.get('body', ''))
            
            # Обновляем статус
            self.status_bar.config(text=f"Редактирование письма для: {email_data['to']}")
    
    def insert_placeholder(self):
        """Вставка плейсхолдера в текст"""
        placeholder = self.placeholder_var.get()
        if not placeholder:
            messagebox.showwarning("Предупреждение", "Выберите плейсхолдер")
            return
        
        # Вставляем в текущую позицию курсора
        self.text_editor.insert(tk.INSERT, placeholder)
    
    def apply_to_all(self):
        """Применить текущий текст ко всем получателям"""
        if not self.email_items:
            messagebox.showwarning("Предупреждение", "Нет получателей")
            return
        
        current_text = self.text_editor.get('1.0', tk.END).strip()
        if not current_text:
            messagebox.showwarning("Предупреждение", "Текст письма пуст")
            return
        
        if messagebox.askyesno("Подтверждение", "Применить текущий текст ко всем получателям?"):
            for email_data in self.email_items:
                email_data['body'] = current_text
            
            self.update_tree()
            self.status_bar.config(text="Текст применен ко всем получателям")
            messagebox.showinfo("Успех", f"Текст применен к {len(self.email_items)} получателям")
    
    def reset_text(self):
        """Сброс текста к шаблону"""
        if not self.template_body:
            messagebox.showwarning("Предупреждение", "Шаблон не загружен")
            return
        
        if messagebox.askyesno("Подтверждение", "Сбросить текст к шаблону для текущего получателя?"):
            self.text_editor.delete('1.0', tk.END)
            self.text_editor.insert('1.0', self.template_body)
            
            # Сохраняем для текущего получателя
            selection = self.tree.selection()
            if selection:
                item = selection[0]
                index = self.tree.index(item)
                if index < len(self.email_items):
                    self.email_items[index]['body'] = self.template_body
            
            self.status_bar.config(text="Текст сброшен к шаблону")
    
    def preview_email(self):
        """Предпросмотр письма в Outlook"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите получателя")
            return
        
        item = selection[0]
        index = self.tree.index(item)
        
        if index >= len(self.email_items):
            return
        
        email_data = self.email_items[index]
        
        try:
            # Создаем письмо из шаблона
            mail = self.outlook.CreateItemFromTemplate(self.draft_path)
            
            # Обновляем данные
            mail.To = email_data['to']
            if email_data['cc']:
                mail.CC = email_data['cc']
            mail.Subject = email_data['subject']
            mail.Body = email_data['body']
            
            # Отображаем для предпросмотра
            mail.Display(False)
            self.status_bar.config(text=f"Открыто окно предпросмотра для: {email_data['to']}")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть предпросмотр: {e}")
    
    def open_selected(self):
        """Открыть выбранное письмо для редактирования в Outlook"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите получателя")
            return
        
        item = selection[0]
        index = self.tree.index(item)
        
        if index >= len(self.email_items):
            return
        
        email_data = self.email_items[index]
        
        try:
            # Создаем письмо из шаблона
            mail = self.outlook.CreateItemFromTemplate(self.draft_path)
            
            # Обновляем данные
            mail.To = email_data['to']
            if email_data['cc']:
                mail.CC = email_data['cc']
            mail.Subject = email_data['subject']
            mail.Body = email_data['body']
            
            # Отображаем для редактирования
            mail.Display(False)
            self.status_bar.config(text=f"Открыто окно редактирования для: {email_data['to']}")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть письмо: {e}")
    
    def send_all(self):
        """Отправка всех писем"""
        if not self.email_items:
            messagebox.showwarning("Предупреждение", "Нет получателей для отправки")
            return
        
        # Проверяем, все ли тексты заполнены
        empty_texts = []
        for i, email_data in enumerate(self.email_items):
            if not email_data.get('body', '').strip():
                empty_texts.append(str(i + 1))
        
        if empty_texts:
            messagebox.showwarning("Предупреждение", 
                f"У следующих получателей пустой текст письма: {', '.join(empty_texts)}\n"
                "Заполните текст перед отправкой.")
            return
        
        # Подтверждение отправки
        count = len(self.email_items)
        if not messagebox.askyesno("Подтверждение", 
            f"Отправить {count} писем?\n\n"
            "Письма будут отправлены автоматически без дополнительного подтверждения."):
            return
        
        # Отправляем письма
        sent = 0
        failed = []
        
        for i, email_data in enumerate(self.email_items):
            try:
                # Создаем письмо из шаблона
                mail = self.outlook.CreateItemFromTemplate(self.draft_path)
                
                # Обновляем данные
                mail.To = email_data['to']
                if email_data['cc']:
                    mail.CC = email_data['cc']
                mail.Subject = email_data['subject']
                mail.Body = email_data['body']
                
                # Отправляем
                mail.Send()
                sent += 1
                self.status_bar.config(text=f"Отправлено {sent}/{count} писем...")
                self.root.update()
                
            except Exception as e:
                failed.append(f"{email_data['to']}: {str(e)}")
        
        # Результат
        if failed:
            messagebox.showwarning("Частичная отправка", 
                f"Отправлено: {sent}\n"
                f"Ошибок: {len(failed)}\n\n"
                f"Ошибки:\n" + "\n".join(failed))
        else:
            messagebox.showinfo("Успех", f"Все {sent} писем успешно отправлены!")
        
        self.status_bar.config(text=f"Отправлено {sent} писем")
    
    def update_tree(self):
        """Обновление списка получателей"""
        # Очищаем дерево
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Добавляем получателей
        for i, email_data in enumerate(self.email_items, 1):
            # Проверяем, есть ли текст
            has_text = bool(email_data.get('body', '').strip())
            subject = email_data.get('subject', '')
            
            # Добавляем в дерево
            self.tree.insert('', 'end', values=(
                i,
                email_data['to'],
                email_data.get('cc', ''),
                subject[:30] + '...' if len(subject) > 30 else subject
            ))
            
            # Меняем цвет если текст пустой
            if not has_text:
                item = self.tree.get_children()[-1]
                self.tree.item(item, tags=('empty',))
        
        # Настройка тегов
        self.tree.tag_configure('empty', background='#ffe6e6')
    
    def __del__(self):
        """Очистка при закрытии"""
        try:
            pythoncom.CoUninitialize()
        except:
            pass

def main():
    root = tk.Tk()
    app = EmailBatchSender(root)
    root.mainloop()

if __name__ == "__main__":
    main()
