import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os
import logging

class ColumnSelector:
    """
    Менеджер выбора столбцов - полный визуальный клон correspondence_manager
    """
    
    def __init__(self, parent, df, file_path, callback):
        self.parent = parent
        self.df = df
        self.file_path = file_path
        self.callback = callback
        
    
        # Окно
        self._selector_window = None
        self._tree = None
        self._instruction_panel = None
        self._instruction_visible = False
        
        # Данные
        self.selected_name_col = -1
        self.selected_qty_col = -1
        self.full_data = []
        
        # Хранение пользовательских ширин столбцов
        self.column_widths = {}
        
        # Результат
        self._result = {"confirmed": False}
        
        # Элементы для отображения прогресса
        self.progress_frame = None
        self.progress_bar = None
        self.progress_label = None
        self.progress_percent = None
        
    def show(self):
        """Показывает окно выбора столбцов и ожидает результата"""
        self._create_selector_window()
        self._selector_window.wait_window()
        return self._get_result()
        
    def _get_result(self):
        return self._result

    def _create_selector_window(self):
        """Создает окно выбора столбцов - полный клон correspondence_manager"""
        self._selector_window = tk.Toplevel(self.parent)
        self._selector_window.title(f"Мастер импорта спецификации - {os.path.basename(self.file_path)}")
        
        # Полноэкранный режим как в correspondence_manager
        screen_width = self._selector_window.winfo_screenwidth()
        screen_height = self._selector_window.winfo_screenheight()
        self._selector_window.geometry(f"{screen_width}x{screen_height}+0+0")
        self._selector_window.state('zoomed')
        self._selector_window.configure(bg='#f5f5f5')
        self._selector_window.transient(self.parent)
        self._selector_window.grab_set()
        
        # === ВЕРХНЯЯ ИНФОРМАЦИОННАЯ ПАНЕЛЬ ===
        self._create_header_panel()
        
        # === ОСНОВНОЙ КОНТЕЙНЕР С ТАБЛИЦЕЙ ===
        main_container = ttk.Frame(self._selector_window)
        main_container.pack(fill="both", expand=True, padx=20, pady=(0, 1))
        
        content_container = ttk.Frame(main_container)
        content_container.pack(fill="both", expand=True)
        
        # Фрейм для таблицы
        table_frame = ttk.Frame(content_container)
        table_frame.pack(side="left", fill="both", expand=True)
        
        # Панель инструкции (изначально скрыта)
        self._instruction_panel = ttk.Frame(content_container, width=400, style='TFrame')
        self._instruction_panel.place(x=2000, y=0, relheight=1)
        self._instruction_panel.pack_propagate(False)
        self._instruction_visible = False
        
        # Кнопка для показа/скрытия инструкции
        self._toggle_instruction_btn = ttk.Button(
            main_container,
            text="📚 Показать инструкцию",
            command=self._toggle_instruction,
            width=25
        )
        self._toggle_instruction_btn.pack(anchor="ne", pady=(0, 10))
        
        # Создаем содержимое инструкции
        self._create_instruction_content()
        
        # === ТАБЛИЦА С ДАННЫМИ ===
        self._create_data_table(table_frame)
        
        # === ПАНЕЛЬ УПРАВЛЕНИЯ ===
        self._create_control_panel()
        
        # === ПАНЕЛЬ ПРОГРЕССА (изначально скрыта) ===
        self._create_progress_panel()
        
        # Фокус на таблицу
        self._tree.focus_set()
        
        # Инициализируем индикаторы
        self._update_indicators()

    def _create_progress_panel(self):
        """Создает панель для отображения прогресса поиска аналогов"""
        self.progress_frame = ttk.Frame(self._selector_window)
        # Изначально скрываем, показывается только при запуске поиска
        self.progress_frame.pack_forget()
        
        # Заголовок
        progress_title = ttk.Label(
            self.progress_frame,
            text="🔍 ПОИСК АНАЛОГОВ METEOR",
            font=("Segoe UI", 12, "bold"),
            foreground="#2c3e50"
        )
        progress_title.pack(anchor="w", pady=(10, 5))
        
        # Прогресс-бар
        self.progress_bar = ttk.Progressbar(
            self.progress_frame,
            mode='determinate',
            length=600
        )
        self.progress_bar.pack(fill="x", pady=5)
        
        # Метка статуса
        progress_container = ttk.Frame(self.progress_frame)
        progress_container.pack(fill="x", pady=5)
        
        self.progress_label = ttk.Label(
            progress_container,
            text="Подготовка к поиску...",
            font=("Segoe UI", 9),
            foreground="#34495e"
        )
        self.progress_label.pack(side="left")
        
        self.progress_percent = ttk.Label(
            progress_container,
            text="0%",
            font=("Segoe UI", 9, "bold"),
            foreground="#2c3e50"
        )
        self.progress_percent.pack(side="right")

    def show_progress_panel(self):
        """Показывает панель прогресса и скрывает элементы выбора"""
        # Скрываем элементы выбора столбцов
        self._toggle_instruction_btn.pack_forget()
        self.selection_info.pack_forget()
        
        # Показываем панель прогресса
        self.progress_frame.pack(fill="x", padx=20, pady=20, before=self._control_frame)
        
        # Обновляем заголовок окна
        self._selector_window.title(f"Поиск аналогов METEOR - {os.path.basename(self.file_path)}")

    def update_progress(self, value, text):
        """Обновляет прогресс бар и текст статуса"""
        if self.progress_bar and self.progress_bar.winfo_exists():
            self.progress_bar['value'] = value
            self.progress_label.config(text=text)
            self.progress_percent.config(text=f"{int(value)}%")
            self._selector_window.update_idletasks()

    def _create_header_panel(self):
        """Создает верхнюю панель"""
        header_frame = ttk.Frame(self._selector_window, style='TFrame')
        header_frame.pack(fill="x", padx=20, pady=20)
        
        # Заголовок
        title_label = ttk.Label(
            header_frame,
            text="📋 ШАГ 1: ВЫБОР СТОЛБЦОВ ДЛЯ ИМПОРТА",
            font=("Segoe UI", 16, "bold"),
            foreground="#2c3e50",
            background="#f5f5f5"
        )
        title_label.pack(anchor="w", pady=(0, 12))
        
        # Описание процесса
        desc_label = ttk.Label(
            header_frame,
            text="Программа загрузила вашу спецификацию. Теперь нужно указать какие столбцы содержат нужные данные: названия радиаторов и их колмчества",
            font=("Segoe UI", 11),
            foreground="#34495e",
            background="#f5f5f5"
        )
        desc_label.pack(anchor="w", pady=(0, 12))
        
        # Инструкция что делать дальше
        instruction_text = """Что делать дальше:
1. Просмотрите таблицу ниже
2. Нажмите на заголовок столбца с названиями радиаторов - он отметится значком 📝
3. Нажмите на заголовок столбца с количеством - он отметится значком 🔢
4. После выбора обоих столбцов нажмите \"Продолжить\""""
        
        instruction_label = ttk.Label(
            header_frame,
            text=instruction_text,
            font=("Segoe UI", 10),
            foreground="#2c3e50",
            justify="left"
        )
        instruction_label.pack(anchor="w")

    def _create_data_table(self, parent):
        """Создает таблицу с данными"""
        table_container = ttk.Frame(parent)
        table_container.pack(fill="both", expand=True)
        
        # Создаем Treeview с улучшенным стилем
        style = ttk.Style()
        style.configure("ColumnSelector.Treeview", rowheight=25, font=("Segoe UI", 9))
        style.configure("ColumnSelector.Treeview.Heading", font=("Segoe UI", 9, "bold"))
        
        columns = list(range(len(self.df.columns)))
        self._tree = ttk.Treeview(
            table_container, 
            columns=columns, 
            show="headings",
            height=20,
            selectmode="none",
            style="ColumnSelector.Treeview"
        )
        
        # Инициализируем ширины столбцов начальными значениями
        for col in columns:
            self.column_widths[col] = 150  # Начальная ширина
        
        # Настраиваем столбцы с центрированием
        for col in columns:
            col_name = f"{col+1}"
            self._tree.heading(col, text=col_name)
            self._tree.column(col, width=self.column_widths[col], anchor="center", minwidth=100)
        
        # Подготавливаем и заполняем данные
        self._prepare_data()
        self._fill_treeview()
        
        # Настраиваем заголовки с обработчиками
        for col in columns:
            self._tree.heading(col, command=lambda idx=col: self._on_header_click(idx))
        
        # Привязываем обработчик изменения ширины столбцов
        self._tree.bind('<ButtonRelease-1>', self._on_column_resize)
        
        # Скроллбары
        v_scrollbar = ttk.Scrollbar(table_container, orient="vertical", command=self._tree.yview)
        h_scrollbar = ttk.Scrollbar(table_container, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Размещаем таблицу и скроллбары
        self._tree.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        
        table_container.grid_rowconfigure(0, weight=1)
        table_container.grid_columnconfigure(0, weight=1)

    def _on_column_resize(self, event):
        """Обрабатывает изменение ширины столбцов пользователем"""
        # Сохраняем ВСЕ текущие ширины столбцов
        for col in range(len(self.df.columns)):
            current_width = self._tree.column(col, 'width')
            old_width = self.column_widths.get(col, 150)
            
            if current_width != old_width:
                self.column_widths[col] = current_width

    def _prepare_data(self):
        """Подготавливает данные для отображения"""
        self.full_data = []
        
        for r_idx in range(len(self.df)):  
            row_data = []
            for c_idx in range(len(self.df.columns)):
                cell_value = self.df.iloc[r_idx, c_idx]
                display_value = "" if pd.isna(cell_value) else str(cell_value)
                row_data.append(display_value)
            self.full_data.append(row_data)

    def _fill_treeview(self):
        """Заполняет Treeview данными"""
        for item in self._tree.get_children():
            self._tree.delete(item)
        
        for row_data in self.full_data:
            self._tree.insert("", "end", values=row_data)

    def _create_control_panel(self):
        """Создает панель управления - ТОЧНО КАК В CORRESPONDENCE_MANAGER"""
        self._control_frame = ttk.Frame(self._selector_window)
        self._control_frame.pack(fill="x", padx=20, pady=70)
        
        # Левая часть - информация о выборе
        self.selection_info = ttk.Label(
            self._control_frame,
            text="Нажмите на заголовок столбца для выбора...",
            font=("Segoe UI", 10),
            foreground="#34495e",
            background="#f5f5f5"
        )
        self.selection_info.pack(side="left", anchor="w")
        
        # Правая часть - кнопки (ТОЧНО КАК В CORRESPONDENCE_MANAGER)
        button_frame = ttk.Frame(self._control_frame)
        button_frame.pack(side="right")
        
        # Кнопка "Продолжить" - АКТИВНА ТОЛЬКО КОГДА ВЫБРАНЫ ОБА СТОЛБЦА
        self.confirm_btn = ttk.Button(
            button_frame,
            text="Продолжить",
            command=self._on_confirm,
            width=20,
            state="disabled"
        )
        self.confirm_btn.pack(side="left", padx=(0, 10))
        
        # Кнопка "Закрыть"
        ttk.Button(
            button_frame,
            text="Закрыть",
            command=self._on_cancel,
            width=15
        ).pack(side="left", padx=10)

    def _create_instruction_content(self):
        """Создает содержимое панели инструкции"""
        for widget in self._instruction_panel.winfo_children():
            widget.destroy()
        
        # Заголовок инструкции
        title_label = ttk.Label(
            self._instruction_panel,
            text="📚 ИНСТРУКЦИЯ ПО ВЫБОРУ СТОЛБЦОВ",
            font=("Segoe UI", 12, "bold"),
            foreground="#2c3e50",
            background="#f9f9f9"
        )
        title_label.pack(anchor="w", pady=(15, 10), padx=15)
        
        # Содержимое инструкции
        instruction_content = """
🎯 ЦЕЛЬ РАБОТЫ:

Это окно помогает правильно определить какие столбцы 
в вашей спецификации содержат нужные данные для подбора 
аналогов METEOR.

📊 ПОНИМАНИЕ ТАБЛИЦЫ:

• Цифры 1, 2, 3... - номера столбцов в вашем файле
• Данные ниже - предпросмотр содержимого файла
• Прокручивайте таблицу для просмотра всех данных

🎨 ЦВЕТОВАЯ ИНДИКАЦИЯ:

🟢 Зеленый - столбец успешно выбран
🔴 Красный - столбец еще не выбран

🛠️ ОСНОВНЫЕ ДЕЙСТВИЯ:

1. ВЫБОР СТОЛБЦА НАЗВАНИЙ:
- Найдите столбец с названиями радиаторов
- Нажмите на цифру заголовка этого столбца
- Столбец отметится значком 📝

2. ВЫБОР СТОЛБЦА КОЛИЧЕСТВА:
- Найдите столбец с количеством радиаторов  
- Нажмите на цифру заголовка этого столбца
- Столбец отметится значком 🔢

3. ЗАВЕРШЕНИЕ ВЫБОРА:
- Нажмите кнопку "Продолжить"
- Программа загрузит данные по выбранным столбцам

💡 ВАЖНЫЕ МОМЕНТЫ:

• ОБА столбца обязательны для выбора
• Нельзя выбрать один столбец для обоих типов данных
• Если ошиблись - нажмите на выбранный столбец еще раз

🚀 СОВЕТЫ ПО ВЫБОРУ:

• Ищите столбцы с названиями типа:
  "Радиатор Kermi FTV 22 500x800"
  "Radiator Purmo Vertical 33"
  "VC 11 300x1000"

• Для количества ищите столбцы с числами:
  "1", "2", "10" и т.д.

• Если не уверены - посмотрите предпросмотр данных
  в таблице под заголовками
"""
        
        instruction_text = tk.Text(
            self._instruction_panel,
            wrap="word",
            font=("Segoe UI", 9),
            background="#f9f9f9",
            relief="flat",
            padx=15,
            pady=10,
            height=40
        )
        instruction_text.insert("1.0", instruction_content)
        instruction_text.config(state="disabled")
        instruction_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Кнопка скрытия инструкции внутри панели
        close_btn = ttk.Button(
            self._instruction_panel,
            text="📚 Скрыть инструкцию",
            command=self._toggle_instruction,
            width=25
        )
        close_btn.pack(pady=10)

    def _toggle_instruction(self):
        """Показывает/скрывает панель инструкции"""
        if not hasattr(self, '_instruction_panel') or not self._instruction_panel.winfo_exists():
            return
        
        if self._instruction_visible:
            # Скрываем инструкцию (выезжает вправо)
            self._animate_instruction_out()
            self._toggle_instruction_btn.config(text="📚 Показать инструкцию")
            self._instruction_visible = False
        else:
            # Показываем инструкцию (выезжает слева)
            self._animate_instruction_in()
            self._toggle_instruction_btn.config(text="📚 Скрыть инструкцию")
            self._instruction_visible = True

    def _animate_instruction_in(self):
        """Анимация выезда инструкции слева"""
        if not hasattr(self, '_instruction_panel') or not self._instruction_panel.winfo_exists():
            return
        
        # Получаем родительский контейнер
        main_container = self._instruction_panel.master
        container_width = main_container.winfo_width()
        
        # Целевая позиция - прижата к правому краю
        target_x = container_width - 400
        
        current_x = self._instruction_panel.winfo_x()
        
        # Если панель уже на месте, выходим
        if current_x <= target_x:
            return
        
        # Анимация - перемещаем панель влево
        new_x = current_x - 50
        if new_x < target_x:
            new_x = target_x
        
        self._instruction_panel.place(x=new_x, y=0, relheight=1)
        
        # Продолжаем анимацию, пока не достигнем цели
        if new_x > target_x:
            self._selector_window.after(10, self._animate_instruction_in)

    def _animate_instruction_out(self):
        """Анимация заезда инструкции вправо"""
        if not hasattr(self, '_instruction_panel') or not self._instruction_panel.winfo_exists():
            return
        
        # Получаем родительский контейнер
        main_container = self._instruction_panel.master
        container_width = main_container.winfo_width()
        
        current_x = self._instruction_panel.winfo_x()
        target_x = container_width + 400
        
        # Если панель уже скрыта, выходим
        if current_x >= target_x:
            return
        
        # Анимация - перемещаем панель вправо
        new_x = current_x + 50
        if new_x > target_x:
            new_x = target_x
        
        self._instruction_panel.place(x=new_x, y=0, relheight=1)
        
        # Продолжаем анимацию, пока не достигнем цели
        if new_x < target_x:
            self._selector_window.after(10, self._animate_instruction_out)

    def _update_indicators(self):
        """Обновляет индикаторы и подсказки"""
        
        # Обновляем информацию в панели управления
        if self.selected_name_col != -1 and self.selected_qty_col != -1:
            self.selection_info.config(
                text=f"Выбрано: названия - столбец {self.selected_name_col + 1}, количество - столбец {self.selected_qty_col + 1}"
            )
            self.confirm_btn.config(state="normal")
        elif self.selected_name_col != -1:
            self.selection_info.config(text=f"Выбрано: названия - столбец {self.selected_name_col + 1}")
            self.confirm_btn.config(state="disabled")
        else:
            self.selection_info.config(text="Нажмите на заголовок столбца для выбора...")
            self.confirm_btn.config(state="disabled")
        
        # Обновляем заголовки столбцов
        self._update_headers()

    def _update_headers(self):
        """Обновляет оформление заголовков с сохранением пользовательских ширин"""
        for i in range(len(self.df.columns)):
            # Используем сохраненную ширину для этого столбца
            target_width = self.column_widths.get(i, 150)
            
            if i == self.selected_name_col:
                new_text = f"{i+1} 📝"
            elif i == self.selected_qty_col:
                new_text = f"{i+1} 🔢"
            else:
                new_text = f"{i+1}"
            
            # Обновляем текст заголовка
            self._tree.heading(i, text=new_text)
            
            # ВСЕГДА устанавливаем сохраненную ширину
            self._tree.column(i, width=target_width)

    def _on_header_click(self, col_idx):
        """Обрабатывает клик по заголовку столбца"""
        if self.selected_name_col == -1:
            self.selected_name_col = col_idx
        elif self.selected_qty_col == -1 and col_idx != self.selected_name_col:
            self.selected_qty_col = col_idx
        elif col_idx == self.selected_name_col:
            self.selected_name_col = -1
            self.selected_qty_col = -1
        elif col_idx == self.selected_qty_col:
            self.selected_qty_col = -1
        
        self._update_indicators()

    def _on_confirm(self):
        """Обрабатывает подтверждение выбора"""
        if self.selected_name_col == -1 or self.selected_qty_col == -1:
            messagebox.showwarning("Не все столбцы выбраны", "Пожалуйста, выберите ОБА столбца:\n- Столбец с названиями радиаторов\n- Столбец с количеством")
            return

        # Отключаем кнопки выбора
        self.confirm_btn.config(state="disabled")
        
        # Показываем панель прогресса
        self.show_progress_panel()
        
        # Обновляем заголовок
        title_frame = self._selector_window.winfo_children()[1]  # Получаем header_frame
        for widget in title_frame.winfo_children():
            if isinstance(widget, ttk.Label) and "ШАГ 1" in widget.cget("text"):
                widget.config(text="🔍 ШАГ 2: ПОИСК АНАЛОГОВ METEOR")
                break

        self._result = {
            "name_col": self.selected_name_col,
            "qty_col": self.selected_qty_col,
            "confirmed": True
        }

        if self.callback:
            # Передаем результат в callback вместе с self для обновления прогресса
            self.callback(self._result, self)

    def complete_processing(self):
        """Вызывается когда обработка завершена"""
        # Можно добавить финальное сообщение или автоматически закрыть окно
        self.update_progress(100, "Поиск завершен!")
        
        # Автоматически закрываем через 2 секунды
        self._selector_window.after(2000, self._selector_window.destroy)

    def _on_cancel(self):
        """Обрабатывает отмену выбора"""
        self._result = {"confirmed": False}
        
        if self.callback:
            self.callback(self._result)
        
        self._selector_window.grab_release()
        self._selector_window.destroy()