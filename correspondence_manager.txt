import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import re
from typing import Callable, Optional

class CorrespondenceManager:
    """
    Менеджер таблицы соответствия для подбора аналогов METEOR
    """
    
    def __init__(self, parent, matrix_callback: Callable, select_analog_callback: Callable):
        """
        Инициализация менеджера
        """
        self.parent = parent
        self.matrix_callback = matrix_callback
        self.select_analog_callback = select_analog_callback
        
        # Данные
        self._original_correspondence_df = None
        self._filtered_correspondence_df = None
        self._saved_correspondence_data = None
        
        # Окно
        self._correspondence_window = None
        self._correspondence_tree = None
        self._instruction_panel = None
        self._instruction_visible = False
        
    def show_correspondence_table(self, correspondence_df: pd.DataFrame) -> None:
        """Показывает таблицу соответствия - только радиаторов"""
        self._original_correspondence_df = correspondence_df.copy()
        
        # Фильтруем данные
        filtered_df = self._filter_radiators(correspondence_df)
        if filtered_df is None or len(filtered_df) == 0:
            messagebox.showinfo("Информация", "Нет радиаторов для отображения в таблице соответствия")
            return
            
        self._filtered_correspondence_df = filtered_df
        self._create_correspondence_window(filtered_df)
        
    def get_correspondence_data(self) -> Optional[pd.DataFrame]:
        """Возвращает данные таблицы соответствия как DataFrame"""
        if self._saved_correspondence_data is not None:
            return self._saved_correspondence_data.copy()
        elif self._filtered_correspondence_df is not None:
            return self._filtered_correspondence_df.copy()
        else:
            return None
            
    def _filter_radiators(self, correspondence_df: pd.DataFrame) -> Optional[pd.DataFrame]:
        """Фильтрует данные - оставляет только радиаторы"""
        # Оставляем только строки с найденными аналогами или ожидающие подбора
        filtered_df = correspondence_df[
            (correspondence_df["Артикул METEOR"] != "") | 
            (correspondence_df["Источник"] == "Ожидает ручного подбора")
        ]
        
        def is_radiator_row(row):
            try:
                name = str(row["Наименование"]).lower()
                
                # ЯВНО ИСКЛЮЧАЕМ не-радиаторы
                exclusion_keywords = [
                    'арматура', 'фитинг', 'муфта', 'переходник', 
                    "Ридан", 'PE-Xa', 'сшитого полиэтилена', "RLV",
                    'полотенцесушитель', 'работы', 
                    'гидравлическое', 'пусконаладочные', 'электрический'
                ]
                
                if any(keyword in name for keyword in exclusion_keywords):
                    return False
                
                # Ключевые слова радиаторов
                radiator_keywords = [
                    'радиатор', 'radiator', 'vc', 'vk', 'cv', 'oc', 'ov',
                    'k-profil', 'classic', 'prado', 'compact', 'ventil',
                    'тип', 'type', 'evra', 'purmo', 'royal', 'thermo', 'oasis'
                ]
                
                # Проверяем форматы названий
                has_radiator_format = (
                    bool(re.search(r'(cv|vc|oc|ov)\s*\d+\s*\d+x\d+', name)) or
                    bool(re.search(r'\d+[\-\s\/x]+\d+[\-\s\/x]+\d+', name)) or
                    ('тип' in name and any(char.isdigit() for char in name))
                )
                
                has_radiator_keyword = any(keyword in name for keyword in radiator_keywords)
                
                return has_radiator_keyword or has_radiator_format
            except Exception as e:
                print(f"[ERROR] Ошибка в фильтрации строки: {e}")
                return False
        
        # Применяем фильтр
        final_df = filtered_df[filtered_df.apply(is_radiator_row, axis=1)]
        
        print(f"[INFO] В таблицу соответствия включено {len(final_df)} радиаторов (отфильтровано {len(correspondence_df) - len(final_df)} не-радиаторов)")
        return final_df
        
    def _create_correspondence_window(self, final_df: pd.DataFrame) -> None:
        """Создает окно таблицы соответствия"""
        # Создаем окно во весь экран
        self._correspondence_window = tk.Toplevel(self.parent)
        self._correspondence_window.title("Мастер подбора аналогов METEOR")
        
        # Полноэкранный режим
        screen_width = self._correspondence_window.winfo_screenwidth()
        screen_height = self._correspondence_window.winfo_screenheight()
        self._correspondence_window.geometry(f"{screen_width}x{screen_height}+0+0")
        self._correspondence_window.state('zoomed')
        self._correspondence_window.configure(bg='#f5f5f5')
        
        # Обработчик закрытия окна (крестик)
        self._correspondence_window.protocol("WM_DELETE_WINDOW", self._on_window_close)
        
        # === ВЕРХНЯЯ ИНФОРМАЦИОННАЯ ПАНЕЛЬ ===
        self._create_header_panel(final_df)
        
        # === ОСНОВНОЙ КОНТЕЙНЕР С ТАБЛИЦЕЙ ===
        main_container = ttk.Frame(self._correspondence_window)
        main_container.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
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
        self._create_data_table(table_frame, final_df)
        
        # === ПАНЕЛЬ УПРАВЛЕНИЯ ===
        self._create_control_panel(final_df)
        
        # Фокус на таблицу
        self._correspondence_tree.focus_set()
        
        # Автоматически выбираем первую строку
        self._correspondence_window.after(100, self._select_first_row)
        
    def _create_header_panel(self, final_df: pd.DataFrame) -> None:
        """Создает верхнюю информационную панель"""
        header_frame = ttk.Frame(self._correspondence_window, style='TFrame')
        header_frame.pack(fill="x", padx=20, pady=20)
        
        # Заголовок
        title_label = ttk.Label(
            header_frame,
            text="🔍 ШАГ 2: ПРОВЕРКА И УТОЧНЕНИЕ АНАЛОГОВ",
            font=("Segoe UI", 16, "bold"),
            foreground="#2c3e50",
            background="#f5f5f5"
        )
        title_label.pack(anchor="w", pady=(0, 12))
        
        # Описание процесса
        desc_label = ttk.Label(
            header_frame,
            text="Программа обработала вашу спецификацию. Вот что получилось:",
            font=("Segoe UI", 11),
            foreground="#34495e",
            background="#f5f5f5"
        )
        desc_label.pack(anchor="w", pady=(0, 12))
        
        # Статистика
        auto_matched = len(final_df[final_df["Артикул METEOR"] != ""])
        manual_needed = len(final_df) - auto_matched
        
        stats_frame = ttk.Frame(header_frame, style='TFrame')
        stats_frame.pack(fill="x", pady=(0, 12))
        
        # Статистика в виде цветных карточек
        stats_data = [
            ("✅ Автоматически подобрано:", auto_matched, "#27ae60"),
            ("⏳ Ожидает ручного выбора:", manual_needed, "#f39c12"),
            ("📋 Всего строк с радиаторами:", len(final_df), "#3498db")
        ]
        
        for text, count, color in stats_data:
            stat_card = ttk.Frame(stats_frame, relief="solid", borderwidth=1)
            stat_card.pack(side="left", padx=(0, 20))
            
            ttk.Label(
                stat_card,
                text=text,
                font=("Segoe UI", 10, "bold"),
                foreground=color,
                background="white",
                padding=(10, 6)
            ).pack(side="left")
            
            ttk.Label(
                stat_card,
                text=str(count),
                font=("Segoe UI", 10, "bold"),
                foreground="white",
                background=color,
                padding=(10, 6)
            ).pack(side="left")
        
        # Инструкция что делать дальше
        instruction_text = """Что делать дальше:
1. Просмотрите таблицу ниже и сравните правильность автоматического подбора, если есть ошибка в подборе, тогда можно указать аналог вручную
2. Для любой строки можно выбрать/изменить аналог - выделите строку и нажмите \"Выбрать аналог METEOR\"
3. В таблицу могут попасть строки не имеющие отношения к радиаторам. Для удаления строки из таблицы - нажмите правой кнопкой мыши по строке
4. После того как убедитесь, что все аналоги подобраны верно нажмите \"Перенести в матрицу\""""
        
        instruction_label = ttk.Label(
            header_frame,
            text=instruction_text,
            font=("Segoe UI", 10),
            foreground="#2c3e50",
            justify="left"
        )
        instruction_label.pack(anchor="w")
        
    def _create_data_table(self, parent_frame: ttk.Frame, final_df: pd.DataFrame) -> None:
        """Создает таблицу с данными"""
        table_container = ttk.Frame(parent_frame)
        table_container.pack(fill="both", expand=True)
        
        # Создаем Treeview
        style = ttk.Style()
        style.configure("Correspondence.Treeview", rowheight=25, font=("Segoe UI", 9))
        style.configure("Correspondence.Treeview.Heading", font=("Segoe UI", 9, "bold"))
        
        columns = ("Наименование", "Кол-во", "Наименование METEOR", "Артикул METEOR", "Источник")
        self._correspondence_tree = ttk.Treeview(
            table_container, 
            columns=columns, 
            show="headings",
            height=20,
            selectmode="extended",  # Множественное выделение
            style="Correspondence.Treeview"
        )
        
        # Настраиваем столбцы
        column_configs = {
            "Наименование": {"width": 500, "anchor": "w"},
            "Кол-во": {"width": 100, "anchor": "center"},
            "Наименование METEOR": {"width": 400, "anchor": "w"},
            "Артикул METEOR": {"width": 150, "anchor": "center"},
            "Источник": {"width": 250, "anchor": "w"}
        }
        
        for col in columns:
            self._correspondence_tree.heading(col, text=col)
            config = column_configs.get(col, {"width": 100, "anchor": "w"})
            self._correspondence_tree.column(col, width=config["width"], anchor=config["anchor"])
        
        # Заполняем таблицу данными
        for index, row in final_df.iterrows():
            source = row["Источник"]
            meteor_art = row["Артикул METEOR"]
            
            values = (
                row["Наименование"],
                row["Кол-во"],
                row["Наименование METEOR"],
                meteor_art,
                source
            )
            
            item_id = self._correspondence_tree.insert("", "end", values=values)
            
            # Цветовая подсветка строк
            if "Ожидает" in source:
                self._correspondence_tree.item(item_id, tags=("pending",))
            elif "ручн" in source.lower() or "обучен" in source.lower() or "выбран" in source.lower():
                self._correspondence_tree.item(item_id, tags=("manual",))
            else:
                self._correspondence_tree.item(item_id, tags=("auto",))
        
        # Настраиваем цвета строк
        self._correspondence_tree.tag_configure("pending", background="#fff9e6")
        self._correspondence_tree.tag_configure("manual", background="#e6f7ff")
        self._correspondence_tree.tag_configure("auto", background="#f0f9f0")
        
        # Скроллбары
        v_scrollbar = ttk.Scrollbar(table_container, orient="vertical", command=self._correspondence_tree.yview)
        h_scrollbar = ttk.Scrollbar(table_container, orient="horizontal", command=self._correspondence_tree.xview)
        self._correspondence_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Размещаем таблицу и скроллбары
        self._correspondence_tree.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        
        table_container.grid_rowconfigure(0, weight=1)
        table_container.grid_columnconfigure(0, weight=1)
        
        # ПРОСТЫЕ обработчики событий
        self._correspondence_tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self._correspondence_tree.bind("<Button-3>", self._show_context_menu)
        self._correspondence_tree.bind("<Delete>", self._on_delete_key)
        self._correspondence_tree.bind("<Control-a>", self._select_all)
        self._correspondence_tree.bind("<Control-A>", self._select_all)
        
        # Контекстное меню
        self._create_context_menu()
        
    def _create_control_panel(self, final_df: pd.DataFrame) -> None:
        """Создает панель управления"""
        control_frame = ttk.Frame(self._correspondence_window)
        control_frame.pack(fill="x", padx=20, pady=20)
        
        # Левая часть - информация о выделенной строке
        self.selection_info = ttk.Label(
            control_frame,
            text="Выделите строку для просмотра деталей...",
            font=("Segoe UI", 10),
            foreground="#7f8c8d",
            background="#f5f5f5"
        )
        self.selection_info.pack(side="left", anchor="w")
        
        # Правая часть - кнопки
        button_frame = ttk.Frame(control_frame)
        button_frame.pack(side="right")
        
        # Кнопка "Выбрать аналог METEOR"
        self.select_analog_btn = ttk.Button(
            button_frame,
            text="Выбрать аналог METEOR",
            command=lambda: self.select_analog_callback(self._correspondence_tree, final_df),
            width=25,
            state="normal"
        )
        self.select_analog_btn.pack(side="left", padx=(0, 10))
        
        # Кнопка "Перенести в матрицу"
        ttk.Button(
            button_frame,
            text="Перенести в матрицу",
            command=lambda: self._transfer_all_to_matrix(self._correspondence_tree, final_df),
            width=20
        ).pack(side="left", padx=10)
        
        # Кнопка "Закрыть"
        ttk.Button(
            button_frame,
            text="Закрыть",
            command=self._on_window_close,
            width=15
        ).pack(side="left", padx=10)
        
    def _create_context_menu(self) -> None:
        """Создает контекстное меню"""
        self.context_menu = tk.Menu(self._correspondence_tree, tearoff=0)
        self.context_menu.add_command(label="🗑️ Удалить выделенные строки", command=self._delete_selected_rows)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="📋 Выделить все (Ctrl+A)", command=self._select_all)
        
    def _transfer_all_to_matrix(self, tree, correspondence_df) -> None:
        """Переносит все подобранные аналоги в матрицу и закрывает окно"""
        all_items = tree.get_children()
        transferred_count = 0
        
        # СОХРАНЯЕМ ДАННЫЕ ДЛЯ ТАБЛИЦЫ СООТВЕТСТВИЯ
        correspondence_data = []
        
        for item in all_items:
            values = tree.item(item, "values")
            if len(values) < 5:
                continue
                
            original_name = values[0]
            
            try:
                qty_str = values[1] if len(values) > 1 else "0"
                qty = self._parse_quantity(qty_str) if qty_str and str(qty_str).strip() else 0
            except (ValueError, TypeError):
                qty = 0
                
            meteor_art = values[3] if len(values) > 3 else ""
            meteor_name = values[2] if len(values) > 2 else ""
            source = values[4] if len(values) > 4 else ""
            
            # СОБИРАЕМ ДАННЫЕ ДЛЯ ТАБЛИЦЫ СООТВЕТСТВИЯ
            correspondence_data.append({
                'Наименование конкурента': original_name,
                'Количество': qty,
                'Наименование METEOR': meteor_name,
                'Артикул METEOR': meteor_art,
                'Источник подбора': source
            })
            
            if not meteor_art or not str(meteor_art).strip() or qty <= 0:
                continue
                
            transferred_count += 1
        
        # СОХРАНЯЕМ ДАННЫЕ ДЛЯ ТАБЛИЦЫ СООТВЕТСТВИЯ
        if correspondence_data:
            self._saved_correspondence_data = pd.DataFrame(correspondence_data)
        
        # ВЫЗЫВАЕМ CALLBACK ДЛЯ ПЕРЕНОСА В МАТРИЦУ
        if self.matrix_callback:
            transfer_data = {
                'correspondence_data': correspondence_data,
                'transferred_count': transferred_count,
                'saved_correspondence_data': self._saved_correspondence_data,
                'final_df': correspondence_df
            }
            self.matrix_callback(transfer_data)
        
        # ЗАКРЫВАЕМ ОКНО
        if self._correspondence_window:
            self._correspondence_window.destroy()
        self._correspondence_window = None
        self._correspondence_tree = None
        
    def _parse_quantity(self, qty_str: str) -> int:
        """Парсит количество из строки"""
        try:
            return int(float(str(qty_str).strip()))
        except (ValueError, TypeError):
            return 0
            
    def _create_instruction_content(self) -> None:
        """Создает содержимое панели инструкции"""
        # Очищаем панель
        for widget in self._instruction_panel.winfo_children():
            widget.destroy()
        
        # Заголовок инструкции
        title_label = ttk.Label(
            self._instruction_panel,
            text="📚 ИНСТРУКЦИЯ ПО РАБОТЕ",
            font=("Segoe UI", 12, "bold"),
            foreground="#2c3e50",
            background="#f9f9f9"
        )
        title_label.pack(anchor="w", pady=(15, 10), padx=15)
        
        # Содержимое инструкции
        instruction_content = """
Это окно предназначено для сопоставления 
радиаторов конкурентов с аналогами METEOR.

🛠️ ОСНОВНЫЕ ДЕЙСТВИЯ:

1. ВЫБОР АНАЛОГА:
- Выделите строку в таблице
- Нажмите кнопку "Выбрать аналог METEOR"
- В открывшемся окне аналог радиатора конкурента

2. УДАЛЕНИЕ СТРОК:
- Выделите одну или несколько строк 
- Нажмите Delete или правый клик → "Удалить строки"
- Используйте для удаления некорректных записей

3. ЗАВЕРШЕНИЕ РАБОТЫ:
- После подбора всех аналогов нажмите 
  "Перенести в матрицу"
- Данные будут перенесены в основную программу

💡 ПОДСКАЗКИ:

• Для выделения нескольких строк используйте:
  - Ctrl+клик - отдельные строки
  - Shift+клик - диапазон строк
  - Ctrl+A - выделить всё

• Цвета строк:
  🟡 Желтый - ожидает выбора
  🟢 Зеленый - подобран автоматически  
  🔵 Голубой - выбран вручную
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
        
        # Кнопка скрытия инструкции
        close_btn = ttk.Button(
            self._instruction_panel,
            text="📚 Скрыть инструкцию",
            command=self._toggle_instruction,
            width=25
        )
        close_btn.pack(pady=10)

    def _toggle_instruction(self) -> None:
        """Показывает/скрывает панель инструкции"""
        if not hasattr(self, '_instruction_panel') or not self._instruction_panel.winfo_exists():
            return
        
        if self._instruction_visible:
            self._instruction_panel.place(x=2000, y=0, relheight=1)
            self._toggle_instruction_btn.config(text="📚 Показать инструкцию")
            self._instruction_visible = False
        else:
            main_container = self._instruction_panel.master
            container_width = main_container.winfo_width()
            self._instruction_panel.place(x=container_width-400, y=0, relheight=1)
            self._toggle_instruction_btn.config(text="📚 Скрыть инструкцию")
            self._instruction_visible = True

    def _on_tree_select(self, event) -> None:
        """Обработчик выбора строки в таблице"""
        selected_items = self._correspondence_tree.selection()
        
        if selected_items:
            if len(selected_items) == 1:
                # Одна выделенная строка
                item = selected_items[0]
                values = self._correspondence_tree.item(item, "values")
                name = values[0] if len(values) > 0 else ""
                source = values[4] if len(values) > 4 else ""
                
                self.selection_info.config(
                    text=f"Выделено: {name[:60]}{'...' if len(name) > 60 else ''}"
                )
            else:
                # Несколько выделенных строк
                self.selection_info.config(text=f"Выделено строк: {len(selected_items)}")
        else:
            self.selection_info.config(text="Выделите строку для просмотра деталей...")

    def _on_delete_key(self, event) -> None:
        """Обработчик нажатия клавиши Delete"""
        self._delete_selected_rows()
        return "break"

    def _select_all(self, event=None) -> None:
        """Выделяет все строки в таблице"""
        all_items = self._correspondence_tree.get_children()
        if all_items:
            self._correspondence_tree.selection_set(all_items)
        return "break"

    def _show_context_menu(self, event) -> None:
        """Показывает контекстное меню"""
        item = self._correspondence_tree.identify_row(event.y)
        if item:
            # Если кликнули по строке, которая не выделена - выделяем ее
            if item not in self._correspondence_tree.selection():
                self._correspondence_tree.selection_set(item)
            try:
                self.context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.context_menu.grab_release()

    def _delete_selected_rows(self) -> None:
        """Удаляет выбранные строки"""
        selected_items = self._correspondence_tree.selection()
        if selected_items:
            if len(selected_items) == 1:
                message = f"Удалить строку?"
            else:
                message = f"Удалить {len(selected_items)} строк?"
            
            if messagebox.askyesno("Подтверждение", message):
                for item in selected_items:
                    self._correspondence_tree.delete(item)
                
                remaining = len(self._correspondence_tree.get_children())
                self.selection_info.config(text=f"Удалено: {len(selected_items)} строк. Осталось: {remaining}")

    def _select_first_row(self) -> None:
        """Выбирает первую строку в таблице"""
        items = self._correspondence_tree.get_children()
        if items:
            self._correspondence_tree.selection_set(items[0])
            self._correspondence_tree.focus(items[0])

    def _on_window_close(self) -> None:
        """Обработчик закрытия окна (крестик или кнопка Закрыть)"""
        # Сохраняем данные...
        if self._correspondence_tree:
            all_items = self._correspondence_tree.get_children()
            correspondence_data = []
            for item in all_items:
                values = self._correspondence_tree.item(item, "values")
                if len(values) < 5:
                    continue
                original_name = values[0]
                try:
                    qty_str = values[1] if len(values) > 1 else "0"
                    qty = self._parse_quantity(qty_str) if qty_str and str(qty_str).strip() else 0
                except (ValueError, TypeError):
                    qty = 0
                meteor_art = values[3] if len(values) > 3 else ""
                meteor_name = values[2] if len(values) > 2 else ""
                source = values[4] if len(values) > 4 else ""
                correspondence_data.append({
                    'Наименование конкурента': original_name,
                    'Количество': qty,
                    'Наименование METEOR': meteor_name,
                    'Артикул METEOR': meteor_art,
                    'Источник подбора': source
                })
            if correspondence_data:
                self._saved_correspondence_data = pd.DataFrame(correspondence_data)

        # Закрываем окно
        if self._correspondence_window:
            self._correspondence_window.destroy()
            self._correspondence_window = None
            self._correspondence_tree = None

        # 🔥 Вызываем callback с флагом закрытия пользователем
        if self.matrix_callback:
            self.matrix_callback({
                'correspondence_data': correspondence_data,
                'transferred_count': 0,
                'saved_correspondence_data': self._saved_correspondence_data,
                'final_df': None,
                'closed_by_user': True  # <-- важный флаг
            })