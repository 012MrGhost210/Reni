import win32com.client
import os
import re
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import pythoncom
from datetime import datetime
import sys
import json

class EmailBatchSender:
    def __init__(self, root):
        self.root = root
        self.root.title("Массовая отправка писем с шаблоном")
        self.root.geometry("1200x700")
        
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
        self.template_parts = []  # Разбитый шаблон на части
        self.placeholders = []  # Список плейсхолдеров в порядке появления
        self.email_items = []  # Список словарей с данными для каждого письма
        self.current_selection = None  # Текущий выбранный индекс
        
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
        ttk.Button(toolbar, text="Экспорт", command=self.export_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Импорт", command=self.import_data).pack(side=tk.LEFT, padx=5)
        
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
        
        ttk.Label(right_frame, text="Редактирование значений плейсхолдеров:", font=('Arial', 10, 'bold')).pack(anchor=tk.W, pady=(0, 5))
        
        # Информация о плейсхолдерах
        self.placeholder_info = ttk.Label(right_frame, text="Плейсхолдеры не найдены")
        self.placeholder_info.pack(anchor=tk.W, pady=(0, 5))
        
        # Создаем фрейм для редактирования плейсхолдеров
        placeholder_edit_frame = ttk.Frame(right_frame)
        placeholder_edit_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        
        # Canvas для прокрутки полей ввода
        canvas = tk.Canvas(placeholder_edit_frame)
        scrollbar = ttk.Scrollbar(placeholder_edit_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Контейнер для полей ввода плейсхолдеров
        self.placeholder_widgets = {}
        
        # Кнопки управления
        text_buttons_frame = ttk.Frame(right_frame)
        text_buttons_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(text_buttons_frame, text="Предпросмотр", command=self.preview_email).pack(side=tk.LEFT, padx=5)
        ttk.Button(text_buttons_frame, text="Очистить все поля", command=self.clear_all_fields).pack(side=tk.LEFT, padx=5)
        ttk.Button(text_buttons_frame, text="Заполнить тестовыми данными", command=self.fill_test_data).pack(side=tk.LEFT, padx=5)
        
        # Статус бар
        self.status_bar = ttk.Label(self.root, text="Готов к работе", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
    def parse_template(self, body):
        """Разбор шаблона на части и извлечение плейсхолдеров"""
        # Находим все плейсхолдеры с их позициями
        pattern = r'\{([^}]+)\}'
        matches = list(re.finditer(pattern, body))
        
        if not matches:
            return [body], []
        
        # Разбиваем текст на части
        parts = []
        placeholders = []
        last_end = 0
        
        for match in matches:
            # Добавляем текст до плейсхолдера
            if match.start() > last_end:
                parts.append(body[last_end:match.start()])
            
            # Добавляем плейсхолдер как отдельную часть
            placeholder_name = match.group(1)
            parts.append(f"{{{placeholder_name}}}")
            placeholders.append(placeholder_name)
            
            last_end = match.end()
        
        # Добавляем остаток текста
        if last_end < len(body):
            parts.append(body[last_end:])
        
        return parts, placeholders
    
    def build_email_body(self, parts, values):
        """Сборка тела письма из частей и значений"""
        result = []
        value_index = 0
        
        for part in parts:
            if part.startswith('{') and part.endswith('}') and not part.startswith('{{'):
                # Это плейсхолдер
                if value_index < len(values):
                    result.append(values[value_index])
                    value_index += 1
                else:
                    result.append(part)
            else:
                # Обычный текст
                result.append(part)
        
        return ''.join(result)
    
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
            
            # Разбираем шаблон на части
            self.template_parts, self.placeholders = self.parse_template(self.template_body)
            
            # Обновляем информацию
            self.template_info.config(
                text=f"Шаблон загружен: {os.path.basename(file_path)} | Плейсхолдеров: {len(self.placeholders)}",
                foreground="green"
            )
            
            # Обновляем информацию о плейсхолдерах
            if self.placeholders:
                self.placeholder_info.config(
                    text=f"Плейсхолдеры ({len(self.placeholders)}): " + ", ".join(self.placeholders)
                )
            else:
                self.placeholder_info.config(text="Плейсхолдеры не найдены")
            
            # Очищаем список получателей
            self.email_items = []
            self.update_tree()
            
            # Очищаем поля ввода
            for widget in self.placeholder_widgets.values():
                widget.destroy()
            self.placeholder_widgets.clear()
            
            # Создаем поля для ввода значений плейсхолдеров
            if self.placeholders:
                for i, placeholder in enumerate(self.placeholders):
                    frame = ttk.Frame(self.scrollable_frame)
                    frame.pack(fill=tk.X, pady=2)
                    
                    label = ttk.Label(frame, text=f"{{{placeholder}}}:", width=20, anchor=tk.W)
                    label.pack(side=tk.LEFT, padx=5)
                    
                    entry = ttk.Entry(frame, width=50)
                    entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
                    
                    # Сохраняем ссылку на поле ввода
                    self.placeholder_widgets[placeholder] = entry
            
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
        dialog.geometry("400x250")
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
            
            # Создаем словарь для значений плейсхолдеров
            placeholder_values = {}
            for placeholder in self.placeholders:
                placeholder_values[placeholder] = ""
            
            # Создаем данные для письма
            email_data = {
                'to': to,
                'cc': cc,
                'subject': subject,
                'placeholder_values': placeholder_values
            }
            
            self.email_items.append(email_data)
            self.update_tree()
            
            # Выбираем нового получателя
            children = self.tree.get_children()
            if children:
                self.tree.selection_set(children[-1])
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
            item = selection[0]
            index = self.tree.index(item)
            
            del self.email_items[index]
            self.update_tree()
            self.status_bar.config(text="Получатель удален")
            
            # Очищаем поля ввода
            for entry in self.placeholder_widgets.values():
                entry.delete(0, tk.END)
    
    def on_recipient_select(self, event):
        """Обработчик выбора получателя в списке"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        index = self.tree.index(item)
        
        if index < len(self.email_items):
            self.current_selection = index
            email_data = self.email_items[index]
            
            # Загружаем значения плейсхолдеров
            values = email_data.get('placeholder_values', {})
            for placeholder, entry in self.placeholder_widgets.items():
                entry.delete(0, tk.END)
                if placeholder in values:
                    entry.insert(0, values[placeholder])
            
            # Обновляем статус
            self.status_bar.config(text=f"Редактирование письма для: {email_data['to']}")
    
    def get_current_values(self):
        """Получение текущих значений из полей ввода"""
        values = {}
        for placeholder, entry in self.placeholder_widgets.items():
            values[placeholder] = entry.get().strip()
        return values
    
    def save_current_values(self):
        """Сохранение текущих значений для выбранного получателя"""
        if self.current_selection is None:
            return
        
        if self.current_selection < len(self.email_items):
            values = self.get_current_values()
            self.email_items[self.current_selection]['placeholder_values'] = values
    
    def preview_email(self):
        """Предпросмотр письма в Outlook"""
        if not self.draft_path:
            messagebox.showwarning("Предупреждение", "Сначала загрузите шаблон")
            return
        
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите получателя")
            return
        
        # Сохраняем текущие значения
        self.save_current_values()
        
        item = selection[0]
        index = self.tree.index(item)
        
        if index >= len(self.email_items):
            return
        
        email_data = self.email_items[index]
        
        try:
            # Создаем письмо из шаблона
            mail = self.outlook.CreateItemFromTemplate(self.draft_path)
            
            # Собираем тело письма
            values = list(email_data['placeholder_values'].values())
            body = self.build_email_body(self.template_parts, values)
            
            # Обновляем данные
            mail.To = email_data['to']
            if email_data['cc']:
                mail.CC = email_data['cc']
            mail.Subject = email_data['subject']
            mail.Body = body
            
            # Отображаем для предпросмотра
            mail.Display(False)
            self.status_bar.config(text=f"Открыто окно предпросмотра для: {email_data['to']}")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть предпросмотр: {e}")
    
    def open_selected(self):
        """Открыть выбранное письмо для редактирования в Outlook"""
        if not self.draft_path:
            messagebox.showwarning("Предупреждение", "Сначала загрузите шаблон")
            return
        
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите получателя")
            return
        
        # Сохраняем текущие значения
        self.save_current_values()
        
        item = selection[0]
        index = self.tree.index(item)
        
        if index >= len(self.email_items):
            return
        
        email_data = self.email_items[index]
        
        try:
            # Создаем письмо из шаблона
            mail = self.outlook.CreateItemFromTemplate(self.draft_path)
            
            # Собираем тело письма
            values = list(email_data['placeholder_values'].values())
            body = self.build_email_body(self.template_parts, values)
            
            # Обновляем данные
            mail.To = email_data['to']
            if email_data['cc']:
                mail.CC = email_data['cc']
            mail.Subject = email_data['subject']
            mail.Body = body
            
            # Отображаем для редактирования
            mail.Display(False)
            self.status_bar.config(text=f"Открыто окно редактирования для: {email_data['to']}")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть письмо: {e}")
    
    def send_all(self):
        """Отправка всех писем"""
        if not self.draft_path:
            messagebox.showwarning("Предупреждение", "Сначала загрузите шаблон")
            return
        
        if not self.email_items:
            messagebox.showwarning("Предупреждение", "Нет получателей для отправки")
            return
        
        # Сохраняем текущие значения
        self.save_current_values()
        
        # Проверяем, все ли плейсхолдеры заполнены
        empty_placeholders = []
        for i, email_data in enumerate(self.email_items):
            values = email_data.get('placeholder_values', {})
            for placeholder, value in values.items():
                if not value.strip():
                    empty_placeholders.append(f"Письмо #{i+1} ({email_data['to']}): {{{placeholder}}}")
        
        if empty_placeholders:
            if not messagebox.askyesno("Предупреждение", 
                f"Следующие плейсхолдеры не заполнены:\n\n" + 
                "\n".join(empty_placeholders[:5]) + 
                ("\n..." if len(empty_placeholders) > 5 else "") +
                "\n\nПродолжить отправку?"):
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
                
                # Собираем тело письма
                values = list(email_data['placeholder_values'].values())
                body = self.build_email_body(self.template_parts, values)
                
                # Обновляем данные
                mail.To = email_data['to']
                if email_data['cc']:
                    mail.CC = email_data['cc']
                mail.Subject = email_data['subject']
                mail.Body = body
                
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
    
    def clear_all_fields(self):
        """Очистка всех полей ввода"""
        if messagebox.askyesno("Подтверждение", "Очистить все поля для текущего получателя?"):
            for entry in self.placeholder_widgets.values():
                entry.delete(0, tk.END)
            self.save_current_values()
            self.status_bar.config(text="Поля очищены")
    
    def fill_test_data(self):
        """Заполнение тестовыми данными"""
        if not self.placeholders:
            return
        
        for placeholder, entry in self.placeholder_widgets.items():
            entry.delete(0, tk.END)
            entry.insert(0, f"Тестовое значение для {placeholder}")
        
        self.save_current_values()
        self.status_bar.config(text="Поля заполнены тестовыми данными")
    
    def update_tree(self):
        """Обновление списка получателей"""
        # Очищаем дерево
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Добавляем получателей
        for i, email_data in enumerate(self.email_items, 1):
            # Проверяем, все ли плейсхолдеры заполнены
            values = email_data.get('placeholder_values', {})
            all_filled = all(value.strip() for value in values.values()) if values else True
            
            subject = email_data.get('subject', '')
            
            # Добавляем в дерево
            item = self.tree.insert('', 'end', values=(
                i,
                email_data['to'],
                email_data.get('cc', ''),
                subject[:30] + '...' if len(subject) > 30 else subject
            ))
            
            # Меняем цвет если не все поля заполнены
            if not all_filled:
                self.tree.item(item, tags=('empty',))
        
        # Настройка тегов
        self.tree.tag_configure('empty', background='#ffe6e6')
    
    def export_data(self):
        """Экспорт данных в JSON файл"""
        if not self.email_items:
            messagebox.showwarning("Предупреждение", "Нет данных для экспорта")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Сохранить данные",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            # Сохраняем текущие значения
            self.save_current_values()
            
            # Подготавливаем данные для экспорта
            export_data = {
                'template_path': self.draft_path,
                'template_subject': self.template_subject,
                'placeholders': self.placeholders,
                'recipients': self.email_items
            }
            
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(export_data, f, ensure_ascii=False, indent=2)
            
            messagebox.showinfo("Успех", f"Данные экспортированы в {file_path}")
            self.status_bar.config(text=f"Данные экспортированы")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось экспортировать данные: {e}")
    
    def import_data(self):
        """Импорт данных из JSON файла"""
        file_path = filedialog.askopenfilename(
            title="Загрузить данные",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                import_data = json.load(f)
            
            # Загружаем шаблон если путь указан
            if import_data.get('template_path') and os.path.exists(import_data['template_path']):
                self.draft_path = import_data['template_path']
                mail = self.outlook.CreateItemFromTemplate(self.draft_path)
                self.template_body = mail.Body
                self.template_subject = import_data.get('template_subject', mail.Subject)
                self.template_parts, self.placeholders = self.parse_template(self.template_body)
                
                self.template_info.config(
                    text=f"Шаблон загружен: {os.path.basename(self.draft_path)} | Плейсхолдеров: {len(self.placeholders)}",
                    foreground="green"
                )
                
                # Создаем поля ввода
                for widget in self.placeholder_widgets.values():
                    widget.destroy()
                self.placeholder_widgets.clear()
                
                for placeholder in self.placeholders:
                    frame = ttk.Frame(self.scrollable_frame)
                    frame.pack(fill=tk.X, pady=2)
                    
                    label = ttk.Label(frame, text=f"{{{placeholder}}}:", width=20, anchor=tk.W)
                    label.pack(side=tk.LEFT, padx=5)
                    
                    entry = ttk.Entry(frame, width=50)
                    entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
                    
                    self.placeholder_widgets[placeholder] = entry
            
            # Загружаем получателей
            self.email_items = import_data.get('recipients', [])
            self.update_tree()
            
            if self.email_items:
                self.tree.selection_set(self.tree.get_children()[0])
                self.on_recipient_select(None)
            
            messagebox.showinfo("Успех", f"Импортировано {len(self.email_items)} получателей")
            self.status_bar.config(text=f"Данные импортированы из {os.path.basename(file_path)}")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось импортировать данные: {e}")
    
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
