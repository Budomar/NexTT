import tkinter as tk
from tkinter import ttk
import os
import sys
import webbrowser


class InterfaceBuilder:
    """
    Класс для построения графического интерфейса с матрицей радиаторов
    """

    def __init__(self, root, app):
        self.root = root
        self.app = app

    def setup_window_geometry(self):
        """
        Комплексная настройка геометрии главного окна.
        Устанавливает окно на всю ширину экрана с центрированием по вертикали.
        Должен вызываться ПОСЛЕ create_main_interface()
        """
        # 1. Сначала полностью скрываем окно
        self.root.withdraw()

        # 2. Устанавливаем иконку
        try:
            icon_path = self.app.resource_path("icon.ico")
            self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Не удалось установить иконку: {e}")

        # 3. Ждем полного создания интерфейса
        self.root.update_idletasks()

        # 4. Получаем размеры экрана
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # 5. Рассчитываем необходимые размеры
        # Получаем фактическую требуемую высоту содержимого
        main_frame = self.root.winfo_children()[0]
        required_width = main_frame.winfo_reqwidth()
        required_height = main_frame.winfo_reqheight()
        
        # 6. Устанавливаем ширину окна равной ширине экрана
        window_width = screen_width
        
        # 7. Вычисляем высоту окна
        # Добавляем небольшие отступы для окна (рамки окна)
        window_height = required_height + 8  # Рамки + заголовок окна
        
        # Если высота окна больше 90% экрана - ограничиваем
        if window_height > int(screen_height * 0.9):
            window_height = int(screen_height * 0.9)
        
        # 8. Вычисляем позицию для центрирования по вертикали
        x = 0  # Окно начинается с левого края
        y = 0
        
        # 9. Устанавливаем геометрию окна
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 10. Устанавливаем МИНИМАЛЬНЫЕ размеры
        # Минимальная ширина - требуемая ширина контента
        # Минимальная высота - требуемая высота контента
        self.root.minsize(required_width, required_height)
        
        # 11. Устанавливаем МАКСИМАЛЬНУЮ ширину - ширину экрана
        # Это позволит окну растягиваться до ширины экрана
        self.root.maxsize(screen_width, screen_height)

        # 12. Показываем окно
        self.root.deiconify()
        self.root.focus_force()

    def setup_styles(self):
        """
        Настройка стилей виджетов
        """
        style = ttk.Style()
        style.theme_use('clam')

        # Базовые стили
        style.configure('.', background="#e1dfdf", foreground='#444141')
        style.configure('TFrame', background='#dedede')
        style.configure('TLabel', background='#dedede', font=('Segoe UI', 9))
        style.configure('TButton', font=('Segoe UI', 9), background='#955b67', foreground='white')
        style.configure('TEntry', font=('Segoe UI', 9), fieldbackground='white')
        style.configure('TLabelFrame', background="#dedede", font=('Segoe UI', 9, 'bold'))
        style.configure('TRadiobutton', background='#dedede', font=('Segoe UI', 9))
        style.configure('TMenubutton', font=('Segoe UI', 9), background='#955b67', foreground='white')
        style.configure('Highlight.TEntry', fieldbackground='#FFD700', font=('Segoe UI', 9))
        style.configure('Highlighted.TRadiobutton', background='#955b67', foreground='white', font=('Segoe UI', 9))
        style.configure('Update.TButton',
                       font=('Segoe UI', 9),
                       background='#955b67',
                       foreground='#444141')

        # Стили для Treeview
        style.configure("Correspondence.Treeview", rowheight=25, font=("Segoe UI", 9))
        style.configure("Correspondence.Treeview.Heading", font=("Segoe UI", 9, "bold"))

        # Стиль для акцентной кнопки
        style.configure("Accent.TButton",
                       background="#ffd700",
                       foreground="#000000",
                       font=('Segoe UI', 9))

        # Настройки mapping
        style.map('TRadiobutton',
                 background=[('active', '#955b67')],
                 foreground=[('active', 'white')])

        style.map('TButton',
                 background=[('active', "#263168"), ('!disabled', "#7E1A2F")],
                 foreground=[('!disabled', 'white')]
                 )

        style.configure('TMenubutton',
                       background='#955b67',
                       foreground='white',
                       arrowcolor='white')
        style.map('TMenubutton',
                 background=[('active', '#263168'), ('!active', '#7E1A2F')],
                 foreground=[('!active', 'white')],
                 arrowcolor=[('!active', 'white')]
                 )

        style.map("Accent.TButton",
                 background=[('active', '#ffed4a'), ('!disabled', '#ffd700')],
                 foreground=[('!disabled', '#000000')]
                 )

    def create_main_interface(self):
        """
        Создает полный графический интерфейс приложения.
        Каждый элемент создается последовательно с комментариями.
        """
        # 1. Создаем главный контейнер (основной фрейм)
        main_frame = ttk.Frame(self.root)
        # Убираем expand=True, чтобы фрейм не растягивался
        main_frame.pack(fill="both", padx=2, pady=2)

        # 2. Создаем верхнюю панель (меню)
        self.create_top_panel(main_frame)

        # 3. Создаем блок с настройками (3 строки друг под другом)
        self.create_settings_row(main_frame)

        # 4. Создаем контейнер для матрицы радиаторов
        self.create_matrix_container(main_frame)

        # 5. Создаем нижнюю панель с кнопками
        self.app.bottom_panel = ttk.Frame(main_frame)
        self.app.bottom_panel.pack(fill="x", pady=(5, 0))

        # Создаем нижнюю панель с кнопками
        self.create_bottom_panel(self.app.bottom_panel)

        return main_frame

    def create_settings_row(self, parent):
        """
        Создает строку с настройками:
        - Блоки "Вид подключения" и "Тип радиатора" на одном уровне (горизонтально)
        - Чекбокс "Показывать параметры" справа на том же уровне
        """
        self.app.settings_frame = ttk.Frame(parent)
        self.app.settings_frame.pack(fill="x", padx=5, pady=5)

        # 1. Создаем горизонтальный контейнер для всех блоков
        top_row_frame = ttk.Frame(self.app.settings_frame)
        top_row_frame.pack(fill="x", pady=(0, 5))

        # 2. Блок 1: Вид подключения (3 радиокнопки в один ряд) с текстом слева
        # Текст "Вид подключения:" слева
        connection_label = ttk.Label(top_row_frame, text="Вид подключения:", font=('Segoe UI', 9))
        connection_label.pack(side="left", padx=(0, 5), pady=5)

        # Обычный Frame с рамкой для радиокнопок подключения
        connection_frame = tk.Frame(top_row_frame, 
                                   relief="solid",
                                   borderwidth=1,
                                   bg='#dedede')
        connection_frame.pack(side="left", padx=0, pady=0)

        # Создаем подсказки для каждого типа подключения
        self.app.vk_right_tooltip = self.create_image_tooltip(connection_frame, self.app.resource_path("1.png"))
        self.app.vk_left_tooltip = self.create_image_tooltip(connection_frame, self.app.resource_path("2.png"))
        self.app.k_side_tooltip = self.create_image_tooltip(connection_frame, self.app.resource_path("3.png"))

        # Варианты подключения - ВСЕ В ОДИН РЯД
        connections = [
            ("VK-нижнее правое", "VK-правое"),
            ("VK-нижнее левое", "VK-левое"),
            ("K-боковое", "K-боковое")
        ]

        for text, value in connections:
            btn = ttk.Radiobutton(
                connection_frame,  # Изменено: был connection_buttons_frame, теперь connection_frame
                text=text,
                variable=self.app.connection_var,
                value=value,
                command=self.app.update_radiator_types
            )
            btn.pack(side="left", padx=10, pady=3)  # pady=3 для компактности
            self.app.connection_radio_buttons.append(btn)

            # Добавляем подсказки для всех типов подключения
            if value == "VK-правое" and self.app.vk_right_tooltip:
                btn.bind("<Enter>", lambda e, btn=btn: self.show_image_tooltip(self.app.vk_right_tooltip, btn))
                btn.bind("<Leave>", lambda e: self.hide_image_tooltip(self.app.vk_right_tooltip))
            elif value == "VK-левое" and self.app.vk_left_tooltip:
                btn.bind("<Enter>", lambda e, btn=btn: self.show_image_tooltip(self.app.vk_left_tooltip, btn))
                btn.bind("<Leave>", lambda e: self.hide_image_tooltip(self.app.vk_left_tooltip))
            elif value == "K-боковое" and self.app.k_side_tooltip:
                btn.bind("<Enter>", lambda e, btn=btn: self.show_image_tooltip(self.app.k_side_tooltip, btn))
                btn.bind("<Leave>", lambda e: self.hide_image_tooltip(self.app.k_side_tooltip))

        self.app.connection_frame = connection_frame

        # 3. Блок 2: Тип радиатора (7 радиокнопки в один ряд) с текстом слева
        # Текст "Тип радиатора:" слева
        radiator_label = ttk.Label(top_row_frame, text="Тип радиатора:", font=('Segoe UI', 9))
        radiator_label.pack(side="left", padx=(20, 5), pady=5)  # padx=(20,5) - отступ слева 20px

        # Обычный Frame с рамкой для радиокнопок типа радиатора
        self.app.radiator_frame = tk.Frame(top_row_frame,
                                          relief="solid",
                                          borderwidth=1,
                                          bg='#dedede')
        self.app.radiator_frame.pack(side="left", padx=0, pady=0)

        # Фрейм для радиокнопок типа радиатора (сохраняем для обновления)
        radiator_buttons_frame = ttk.Frame(self.app.radiator_frame)
        radiator_buttons_frame.pack(fill="x", padx=5, pady=3)  # pady=3 для компактности

        # Типы радиаторов будут заполнены в update_radiator_types
        # Здесь создаем только контейнер
        self.app.radiator_buttons_frame = radiator_buttons_frame

        # 4. Блок 3: Чекбокс "Показывать параметры" справа
        checkbox_frame = tk.Frame(top_row_frame,
                                 relief="solid",
                                 borderwidth=1,
                                 bg='#dedede')
        checkbox_frame.pack(side="right", padx=(10, 0), pady=0)

        self.app.show_tooltips_check = ttk.Checkbutton(
            checkbox_frame,  # Изменено: был checkbox_frame (ttk.Frame), теперь checkbox_frame (tk.Frame с рамкой)
            text="Показывать параметры",
            variable=self.app.show_tooltips_var,
            command=self.app.toggle_tooltips
        )
        self.app.show_tooltips_check.pack(side="right", padx=5, pady=3)

    def create_bottom_panel(self, parent):
        """
        Создает нижнюю панель с кнопками Предпросмотр, Сброс и блоком Кронштейны и скидки между ними
        """
        # Контейнер для всех элементов нижней панели
        buttons_container = ttk.Frame(parent)
        buttons_container.pack(fill="x", expand=True, pady=0)

        # 1. Кнопка "Предпросмотр" (компактная)
        self.app.preview_btn = ttk.Button(
            buttons_container,
            text="Предпросмотр",
            command=self.app.preview_spec,
            width=15
        )
        self.app.preview_btn.pack(side="left", padx=5)

        # 2. Растягивающийся отступ между кнопкой Предпросмотр и блоком Кронштейны и скидки
        spacer_left = ttk.Frame(buttons_container)
        spacer_left.pack(side="left", expand=True, fill="x")

        # 3. Блок "Кронштейны и скидки" между кнопками
        self.create_bracket_discount_row(buttons_container)

        # 4. Растягивающийся отступ между блоком Кронштейны и скидки и кнопкой Сброс
        spacer_right = ttk.Frame(buttons_container)
        spacer_right.pack(side="left", expand=True, fill="x")

        # 5. Кнопка "Сброс" (компактная) - теперь справа
        self.app.reset_btn = ttk.Button(
            buttons_container,
            text="Сброс",
            command=self.app.reset_fields,
            width=15
        )
        self.app.reset_btn.pack(side="right", padx=5)

    def create_bracket_discount_row(self, parent):
        """
        Создает блок "Кронштейны и скидки" для размещения в нижней панели между кнопками
        Теперь все элементы располагаются в одной строке
        """
        # 1. Текст "Кронштейны и скидки:" слева
        label = ttk.Label(parent, text="Кронштейны и скидки:", font=('Segoe UI', 9))
        label.pack(side="left", padx=(0, 5), pady=5)  # pady=5 для выравнивания

        # 2. Обычный Frame с рамкой вместо LabelFrame
        bracket_frame = tk.Frame(parent, 
                                relief="solid",  # Сплошная рамка
                                borderwidth=1,   # Толщина рамки
                                bg='#dedede')    # Цвет фона как у других элементов
        bracket_frame.pack(side="left", padx=0, pady=0)

        # 3. Радиокнопки кронштейнов - УПАКОВЫВАЕМ НАПРЯМУЮ В bracket_frame
        brackets = ["Настенные", "Напольные", "Без"]
        for bracket in brackets:
            full_text = f"{bracket} кронштейны" if bracket != "Без" else "Без кронштейнов"
            ttk.Radiobutton(
                bracket_frame,  # Прямо в Frame, без промежуточных фреймов
                text=full_text,
                variable=self.app.bracket_var,
                value=full_text
            ).pack(side="left", padx=5, pady=3)  # pady=3 для компактности

        # 4. Скидка на радиаторы - тоже прямо в bracket_frame
        ttk.Label(bracket_frame,
                 text="Радиаторы, %:",
                 width=12).pack(side="left", padx=(10, 0), pady=3)
        
        ttk.Entry(
            bracket_frame,
            textvariable=self.app.radiator_discount_var,
            width=5,
            validate="key",
            validatecommand=(parent.register(self.app.validate_discount), '%P')
        ).pack(side="left", padx=2, pady=3)

        # 5. Скидка на кронштейны - тоже прямо в bracket_frame
        ttk.Label(bracket_frame,
                 text="Кронштейны, %:",
                 width=12).pack(side="left", padx=(5, 0), pady=3)
        
        ttk.Entry(
            bracket_frame,
            textvariable=self.app.bracket_discount_var,
            width=5,
            validate="key",
            validatecommand=(parent.register(self.app.validate_discount), '%P')
        ).pack(side="left", padx=2, pady=3)
        
    def create_matrix_cell(self, sheet_name, data, length, height, row, col):
        """
        Создает ячейку матрицы радиаторов
        """
        pattern = f"/{height}/{length}"
        match = data[data['Наименование'].str.contains(pattern, na=False)]

        if not match.empty:
            product = match.iloc[0]
            art = str(product['Артикул']).strip()
            value = self.app.entry_values.get((sheet_name, art), "")

            entry = tk.Entry(
                self.app.scrollable_matrix_frame,
                width=5,
                justify="center",
                bg='#e6f3ff' if self.app.has_any_value() else 'white',
                relief='solid',
                borderwidth=1,
                validate='key',
                validatecommand=(self.root.register(self.app.validate_input), '%P'),
                font=('Segoe UI', 10)  # БЫЛО 11, СТАЛО 10
            )
            entry.insert(0, value)

            # Привязываем события
            entry.bind("<FocusIn>", lambda e: self.app.on_entry_focus_in(e))
            entry.bind("<FocusOut>", lambda e, s=sheet_name, a=art: self.app.on_entry_focus_out(e, s, a))
            entry.bind("<Return>", lambda e, s=sheet_name, a=art: self.app.on_entry_focus_out(e, s, a))
            entry.bind("<Tab>", lambda e, s=sheet_name, a=art: self.app.on_entry_focus_out(e, s, a))
            entry.bind("<Enter>", lambda e, p=product: self.app.show_tooltip_on_hover(p))
            entry.bind("<Leave>", lambda e: self.app.hide_tooltip_on_leave())

            # Добавляем привязки для навигации стрелками
            entry.bind("<Up>", lambda e, r=row, c=col: self.navigate_matrix(e, r, c, "up"))
            entry.bind("<Down>", lambda e, r=row, c=col: self.navigate_matrix(e, r, c, "down"))
            entry.bind("<Left>", lambda e, r=row, c=col: self.navigate_matrix(e, r, c, "left"))
            entry.bind("<Right>", lambda e, r=row, c=col: self.navigate_matrix(e, r, c, "right"))

            self.app.entries[(sheet_name, art)] = entry
            entry.grid(row=row, column=col, sticky="nsew", padx=1, pady=1)  # ВЕРНУЛИ КАК БЫЛО
        else:
            entry = tk.Entry(
                self.app.scrollable_matrix_frame,
                width=5,
                justify="center",
                bg='#f0f0f0',
                relief='solid',
                borderwidth=1,
                state='disabled',
                font=('Segoe UI', 10)  # БЫЛО 11, СТАЛО 10
            )
            entry.grid(row=row, column=col, sticky="nsew", padx=1, pady=1)  # ВЕРНУЛИ КАК БЫЛО

    def navigate_matrix(self, event, current_row, current_col, direction):
        """
        Навигация по матрице с помощью стрелок клавиатуры
        """
        # Получаем все дочерние виджеты матрицы
        children = self.app.scrollable_matrix_frame.grid_slaves()

        # Находим следующую ячейку в зависимости от направления
        if direction == "up":
            target_row = current_row - 1
            target_col = current_col
        elif direction == "down":
            target_row = current_row + 1
            target_col = current_col
        elif direction == "left":
            target_row = current_row
            target_col = current_col - 1
        elif direction == "right":
            target_row = current_row
            target_col = current_col + 1
        else:
            return "break"

        # Ищем Entry виджет на нужной позиции
        target_entry = None
        for child in children:
            info = child.grid_info()
            if info:
                row = info.get('row', -1)
                col = info.get('column', -1)
                if row == target_row and col == target_col and isinstance(child, tk.Entry) and child['state'] != 'disabled':
                    target_entry = child
                    break

        # Если нашли Entry - переходим к нему
        if target_entry:
            target_entry.focus_set()
            target_entry.select_range(0, tk.END)

        # Всегда возвращаем "break" чтобы предотвратить стандартную обработку
        return "break"

    def create_top_panel(self, parent):
        """
        Создает верхнюю панель с меню
        """
        self.app.top_panel = ttk.Frame(parent)
        self.app.top_panel.pack(fill="x", pady=(0, 10))

        # Контейнер для кнопок меню с равномерным распределением
        menu_frame = ttk.Frame(self.app.top_panel)
        menu_frame.pack(fill="x", expand=True)

        buttons = [
            ("Создать", [
                ("Спецификацию LaggarTT", lambda: self.app.generate_spec("excel")),
                ("Файл LaggarTT CSV", lambda: self.app.generate_spec("csv")),
            ]),
            ("Загрузить из", [
                ("Спецификации LaggarTT (METEOR)", self.app.load_excel_spec),
                ("Файла LaggarTT (METEOR) CSV", self.app.load_csv_spec),
                ("Иной спецификации", self.app.load_foreign_spec),
                ("Файла PDF - в разработке", self.app.load_pdf_spec)
            ]),
            ("Информация", [
                ("Лицензионное соглашение", lambda: self.app.show_info("agreement")),
                ("Инструкция по использованию программы", self.app.open_instruction_pdf),
                ("Прайс-лист от 15.12.25", self.app.open_price_list),
                ("Формуляр для регистрации проектов", self.app.open_project_form),
                ("Каталог отопительного оборудования", lambda: webbrowser.open("https://laggartt.ru/catalogs/laggartt/#p=386")),
                ("Паспорт на радиатор", lambda: webbrowser.open("https://b24.ez.meteor.ru/~yKwOz")),
                ("Сертификат соответствия", lambda: webbrowser.open("https://b24.ez.meteor.ru/~jmXyC")),
            ])
        ]

        # Создаем кнопки меню
        for i, (btn_text, menu_items) in enumerate(buttons):
            btn = ttk.Menubutton(menu_frame, text=btn_text)
            menu = tk.Menu(btn, tearoff=0)

            if btn_text == "Информация":
                for idx, (item_text, command) in enumerate(menu_items):
                    menu.add_command(label=item_text, command=command)
                    if idx == 1 or idx == 4:
                        menu.add_separator()
            else:
                for item_text, command in menu_items:
                    menu.add_command(label=item_text, command=command)

            btn["menu"] = menu
            btn.grid(row=0, column=i, sticky="ew", padx=5)

        # Добавляем кнопку "Проверить обновление" как отдельную кнопку (не меню)
        update_btn = ttk.Button(
            menu_frame,
            text="Проверить обновление",
            command=lambda: webbrowser.open("https://b24.engpx.ru/~DUCV2"),
            width=18
        )
        update_btn.grid(row=0, column=len(buttons), sticky="ew", padx=5)

        # Настраиваем равномерное распределение колонок (теперь 4 колонки)
        for i in range(len(buttons) + 1):  # +1 для новой кнопки
            menu_frame.grid_columnconfigure(i, weight=1)

    def create_matrix_container(self, parent):
        """
        Создает контейнер для матрицы радиаторов
        """
        self.app.matrix_container = ttk.Frame(parent)
        # Убираем expand=True
        self.app.matrix_container.pack(fill="x", padx=2, pady=2)

        # УБИРАЕМ ВЕРТИКАЛЬНЫЙ ТЕКСТ "высота радиаторов, мм"
        # Вместо Canvas оставляем просто пустой фрейм для отступов
        self.app.height_label_canvas = ttk.Frame(self.app.matrix_container, width=5)
        self.app.height_label_canvas.grid(row=0, column=0, sticky="ns", padx=(0, 0))

        # Внутренний фрейм матрицы радиаторов
        self.app.scrollable_matrix_frame = ttk.Frame(self.app.matrix_container, style='TFrame')
        self.app.scrollable_matrix_frame.grid(row=0, column=1, sticky="nsew")

        # Настраиваем растяжение колонок grid в matrix_container
        self.app.matrix_container.columnconfigure(0, weight=0)  # Пустой фрейм фиксирован
        self.app.matrix_container.columnconfigure(1, weight=1)  # Матрица растягивается
        self.app.matrix_container.rowconfigure(0, weight=1)

    def create_image_tooltip(self, parent, image_path):
        """
        Создает подсказку с картинкой
        """
        try:
            # Проверяем существование файла
            if not os.path.exists(image_path):
                print(f"Файл изображения не найден: {image_path}")
                return None
            tooltip = tk.Toplevel(parent)
            tooltip.wm_overrideredirect(True)
            tooltip.withdraw()
            img = tk.PhotoImage(file=image_path)
            label = ttk.Label(tooltip, image=img)
            label.image = img  # Сохраняем ссылку на изображение
            label.pack()
            return tooltip
        except Exception as e:
            print(f"Ошибка загрузки изображения для подсказки: {e}")
            return None

    def show_image_tooltip(self, tooltip, widget):
        """
        Показывает подсказку с картинкой под виджетом
        """
        if tooltip:
            # Получаем координаты виджета для вертикального позиционирования
            y = widget.winfo_rooty() + widget.winfo_height() + 80
            # Получаем координаты окна для горизонтального центрирования
            window_x = self.root.winfo_rootx()
            window_width = self.root.winfo_width()
            # Центрируем подсказку по горизонтали относительно окна
            tooltip_width = tooltip.winfo_reqwidth()
            x = window_x + (window_width - tooltip_width) // 2
            tooltip.wm_geometry(f"+{x}+{y}")
            tooltip.deiconify()

    def hide_image_tooltip(self, tooltip):
        """
        Скрывает подсказку с картинкой
        """
        if tooltip and tooltip.winfo_exists():
            tooltip.withdraw()

    def show_selected_matrix(self):
        """
        Показывает матрицу для выбранного типа подключения и радиатора
        """
        for widget in self.app.scrollable_matrix_frame.winfo_children():
            widget.destroy()

        sheet_name = f"{self.app.connection_var.get()} {self.app.radiator_type_var.get()}"

        if sheet_name not in self.app.sheets:
            tk.messagebox.showerror("Ошибка", f"Лист '{sheet_name}' не найден")
            return

        data = self.app.sheets[sheet_name]
        lengths = list(range(400, 3100, 100))
        heights = [300, 400, 500, 600, 900]

        # Создаем стиль для заголовков - ШРИФТ 10 ВМЕСТО 11
        style = ttk.Style()
        style.configure('MatrixHeader.TLabel', 
                       font=('Segoe UI', 9),  # БЫЛО 11, СТАЛО 10
                       relief='ridge',
                       borderwidth=2,
                       background='#e8e8e8',
                       foreground='#000000')

        style.configure('MatrixRowHeader.TLabel',
                       font=('Segoe UI', 9),  # БЫЛО 11, СТАЛО 10
                       relief='ridge',
                       borderwidth=2,
                       background='#e8e8e8',
                       foreground='#000000')

        # Заголовки столбцов (длины) - в первой строке
        for j, l in enumerate(lengths):
            label = ttk.Label(
                self.app.scrollable_matrix_frame,
                text=str(l),
                width=5,  # УМЕНЬШИЛИ С 6 ДО 5 (для шрифта 10)
                style='MatrixHeader.TLabel',
                anchor="center"
            )
            label.grid(row=0, column=j+1, sticky="nsew", padx=1, pady=1)

        # Заголовки строк (высоты) - во второй и последующих строках
        for i, h in enumerate(heights):
            label = ttk.Label(
                self.app.scrollable_matrix_frame,
                text=str(h),
                width=5,  # УМЕНЬШИЛИ С 6 ДО 5 (для шрифта 10)
                style='MatrixRowHeader.TLabel',
                anchor="center"
            )
            label.grid(row=i+1, column=0, sticky="nsew", padx=1, pady=1)

            # Ячейки с радиаторами
            for j, l in enumerate(lengths):
                self.create_matrix_cell(sheet_name, data, l, h, i+1, j+1)

        # Уменьшаем минимальные размеры ячеек для шрифта 10
        for col in range(len(lengths) + 1):
            self.app.scrollable_matrix_frame.columnconfigure(col, minsize=65, weight=1)  # БЫЛО 72, СТАЛО 65

        # Уменьшаем высоту строк
        for i in range(len(heights) + 1):  # +1 для заголовка столбцов
            self.app.scrollable_matrix_frame.rowconfigure(i, minsize=30)  # БЫЛО 32, СТАЛО 30

    def create_context_menu(self, tree, spec_data):
        """
        Создает контекстное меню для удаления строк
        """
        context_menu = tk.Menu(tree, tearoff=0)
        context_menu.add_command(
            label="Удалить",
            command=lambda: self.app.delete_selected_row(tree, spec_data)
        )
        context_menu.add_separator()  # Добавляем разделитель

        def show_context_menu(event):
            item = tree.identify_row(event.y)
            if item:
                tree.selection_set(item)
                context_menu.post(event.x_root, event.y_root)

        tree.bind("<Button-3>", show_context_menu)