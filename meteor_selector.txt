# meteor_selector.py
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from typing import Callable, Optional
import numpy as np
import logging

class MeteorSelector:
    """
    Менеджер выбора аналога LaggarTT с реальными данными
    """
    
    def __init__(self, parent, main_app, close_callback: Optional[Callable] = None):
        self.parent = parent
        self.main_app = main_app
        self.close_callback = close_callback
        
        # Настройка логирования
        self.logger = logging.getLogger('MeteorSelector')
        if not self.logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            self.logger.addHandler(handler)
            self.logger.setLevel(logging.DEBUG)
        
        # Данные
        self.original_name = ""
        self.current_item = None
        self.tree_reference = None
        
        # Текущий выбор
        self.selected_connection = tk.StringVar(value="VK-правое")
        self.selected_type = tk.StringVar(value="10")
        self.preview_length = None
        self.preview_height = None
        
        # Размеры матрицы
        self.heights = [300, 400, 500, 600, 900]  # 5 строк
        self.lengths = list(range(400, 3100, 100))  # 27 столбцов (400-3000)
        
        # Окно и виджеты
        self._selector_window = None
        self.matrix_frame = None
        self.matrix_cells = {}
        
    def show_analog_selector(self, tree, item, original_name: str) -> None:
        """Показывает окно выбора аналога LaggarTT"""
        self.logger.debug(f"show_analog_selector вызван для: {original_name}")
        self.tree_reference = tree
        self.current_item = item
        self.original_name = original_name
        
        self._create_selector_window()
        
    def _create_selector_window(self) -> None:
        """Создает окно выбора аналога с финальной структурой"""
        self.logger.debug("Создание окна селектора")
        self._selector_window = tk.Toplevel(self.parent)
        self._selector_window.title("Подбор аналогов LaggarTT")
        
        # Полноэкранный режим
        screen_width = self._selector_window.winfo_screenwidth()
        screen_height = self._selector_window.winfo_screenheight()
        self._selector_window.geometry(f"{screen_width}x{screen_height}+0+0")
        self._selector_window.state('zoomed')
        self._selector_window.configure(bg='#f5f5f5')
        
        # Привязываем обработчик изменения размера окна
        self._selector_window.bind("<Configure>", self._on_window_resize)
        
        # === ГЛАВНЫЙ КОНТЕЙНЕР (вертикальный) ===
        main_container = ttk.Frame(self._selector_window)
        main_container.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Распределение строк
        main_container.rowconfigure(0, weight=0)    # Заголовок
        main_container.rowconfigure(1, weight=0)    # Блок настроек
        main_container.rowconfigure(2, weight=0)    # Матрица (фиксированная высота)
        main_container.rowconfigure(3, weight=1)    # Инструкция (растягивается)
        main_container.columnconfigure(0, weight=1)
        
        # === 1. ЗАГОЛОВОК ===
        header_frame = ttk.Frame(main_container)
        header_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 15))
        
        # Заголовок
        title_text = tk.Text(
            header_frame,
            font=("Segoe UI", 16, "bold"),
            foreground="#2c3e50",
            background="#f5f5f5",
            wrap="word",
            height=3,
            padx=0,
            pady=0,
            borderwidth=0,
            highlightthickness=0,
            relief="flat"
        )
        title_text.pack(anchor="w", fill="x")
        title_text.insert("1.0", f"Подбор аналога LaggarTT для: {self.original_name}")
        title_text.config(state="disabled")
        
        # === 2. БЛОК НАСТРОЕК (горизонтальный) ===
        settings_frame = ttk.Frame(main_container)
        settings_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 20))
        
        # Контейнер для горизонтального расположения
        settings_container = ttk.Frame(settings_frame)
        settings_container.pack(fill="x", expand=True)
        
        # === ШАГ 1: ПОДКЛЮЧЕНИЕ ===
        step1_frame = ttk.LabelFrame(settings_container, text="ШАГ 1: ПОДКЛЮЧЕНИЕ", padding=10)
        step1_frame.pack(side="left", fill="y", padx=(0, 15))
        
        connections = ["VK-правое", "VK-левое", "K-боковое"]
        for conn in connections:
            rb = ttk.Radiobutton(
                step1_frame,
                text=conn,
                variable=self.selected_connection,
                value=conn,
                command=self._on_connection_changed
            )
            rb.pack(anchor="w", pady=2)
        
        # === ШАГ 2: ТИП РАДИАТОРА ===
        self.step2_frame = ttk.LabelFrame(settings_container, text="ШАГ 2: ТИП РАДИАТОРА", padding=10)
        self.step2_frame.pack(side="left", fill="y", padx=(0, 15))
        
        self._update_available_types()
        
        # === ИТОГИ ВЫБОРА ===
        summary_frame = ttk.LabelFrame(settings_container, text="ИТОГИ ВЫБОРА", padding=10)
        summary_frame.pack(side="left", fill="both", expand=True, padx=(0, 15))
        
        # Переменные для отображения
        self.analog_name_var = tk.StringVar(value="Выберите параметры")
        self.article_var = tk.StringVar(value="—")
        self.power_var = tk.StringVar(value="—")
        self.weight_var = tk.StringVar(value="—")
        self.volume_var = tk.StringVar(value="—")
        
        # Компактное отображение
        summary_grid = ttk.Frame(summary_frame)
        summary_grid.pack(fill="both", expand=True)
        
        # Название аналога
        name_label = ttk.Label(
            summary_grid,
            textvariable=self.analog_name_var,
            font=("Segoe UI", 10, "bold"),
            foreground="#2c3e50",
            wraplength=250
        )
        name_label.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 5))
        
        # Артикул
        ttk.Label(summary_grid, text="Артикул:", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky="w")
        ttk.Label(summary_grid, textvariable=self.article_var, font=("Segoe UI", 9)).grid(row=1, column=1, sticky="w", padx=(5, 0))
        
        # Мощность
        ttk.Label(summary_grid, text="Мощность:", font=("Segoe UI", 9, "bold")).grid(row=2, column=0, sticky="w", pady=(2, 0))
        ttk.Label(summary_grid, textvariable=self.power_var, font=("Segoe UI", 9)).grid(row=2, column=1, sticky="w", padx=(5, 0), pady=(2, 0))
        
        # Вес
        ttk.Label(summary_grid, text="Вес:", font=("Segoe UI", 9, "bold")).grid(row=3, column=0, sticky="w", pady=(2, 0))
        ttk.Label(summary_grid, textvariable=self.weight_var, font=("Segoe UI", 9)).grid(row=3, column=1, sticky="w", padx=(5, 0), pady=(2, 0))
        
        # Объем
        ttk.Label(summary_grid, text="Объем:", font=("Segoe UI", 9, "bold")).grid(row=4, column=0, sticky="w", pady=(2, 0))
        ttk.Label(summary_grid, textvariable=self.volume_var, font=("Segoe UI", 9)).grid(row=4, column=1, sticky="w", padx=(5, 0), pady=(2, 0))
        
        summary_grid.columnconfigure(1, weight=1)
        
        # === 3. МАТРИЦА ===
        matrix_container = ttk.LabelFrame(main_container, text="ШАГ 3: ВЫБОР РАЗМЕРА")
        matrix_container.grid(row=2, column=0, sticky="nsew", pady=(0, 20))
        
        # ФИКСИРОВАННАЯ ВЫСОТА матрицы
        matrix_container.configure(height=230)  # 6 строк × 38px + отступы
        
        # Внутренний фрейм матрицы
        self.matrix_frame = ttk.Frame(matrix_container)
        self.matrix_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Создаем матрицу
        self._create_adaptive_matrix()
        
        # === 4. НИЖНИЙ БЛОК (Инструкция и кнопка) ===
        bottom_container = ttk.Frame(main_container)
        bottom_container.grid(row=3, column=0, sticky="nsew")
        
        # Разделяем на 2 части: инструкция (80%) и кнопка (20%)
        bottom_container.columnconfigure(0, weight=80)  # Инструкция
        bottom_container.columnconfigure(1, weight=20)  # Кнопка
        
        # === ИНСТРУКЦИЯ (3 колонки) ===
        instruction_frame = ttk.LabelFrame(bottom_container, text="ИНСТРУКЦИЯ", padding=10)
        instruction_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        
        # Создаем инструкцию в 3 колонках
        self._create_three_column_instruction(instruction_frame)
        
        # === КНОПКА ЗАКРЫТЬ (отдельный блок) ===
        button_frame = ttk.Frame(bottom_container)
        button_frame.grid(row=0, column=1, sticky="nsew")
        
        # Контейнер для центрирования кнопки
        btn_container = ttk.Frame(button_frame)
        btn_container.pack(expand=True, fill="both")
        
        # Кнопка Закрыть (центрирована)
        close_btn = ttk.Button(
            btn_container,
            text="Закрыть",
            command=self._on_close,
            width=15
        )
        close_btn.pack(expand=True, pady=(50, 0))
        
        # Обновляем предпросмотр
        self._update_preview()
        
        # Фокус на окно
        self._selector_window.focus_set()
        self.logger.debug("Окно селектора создано успешно")
        
    def _create_adaptive_matrix(self) -> None:
        """Создает адаптивную матрицу 5×27 с увеличенной высотой ячеек"""
        self.logger.debug("Создание адаптивной матрицы")
        # Очищаем матрицу
        for widget in self.matrix_frame.winfo_children():
            widget.destroy()
        
        self.matrix_cells.clear()
        
        # Получаем ширину контейнера
        container_width = self.matrix_frame.winfo_width()
        if container_width < 100:
            container_width = self._selector_window.winfo_width() - 50
            
        # Рассчитываем размер ячейки
        num_columns = len(self.lengths) + 1
        min_cell_width = 60
        calculated_width = max(min_cell_width, (container_width // num_columns))
        self.logger.debug(f"Ширина ячейки: {calculated_width}px, Колонок: {num_columns}")
        
        # УВЕЛИЧЕННАЯ ВЫСОТА ЯЧЕЕК
        row_height = 38
        
        # Размер шрифта
        font_size = 9 if calculated_width >= 70 else 8 if calculated_width >= 60 else 7
        
        # === СТИЛИ ===
        style = ttk.Style()
        style.configure('MatrixHeader.TLabel',
                       font=('Segoe UI', font_size, 'bold'),
                       relief='ridge',
                       borderwidth=1,
                       background='#e8e8e8',
                       foreground='#000000',
                       anchor="center",
                       justify="center")
        
        style.configure('MatrixCell.TButton',
                       font=('Segoe UI', font_size),
                       padding=2)
        
        # === СОЗДАЕМ МАТРИЦУ ===
        
        # 1. Пустой уголок - ЦЕНТРИРОВАННЫЙ текст
        corner_label = ttk.Label(
            self.matrix_frame,
            text="ДЛИНА → \nВЫСОТА ↓ ",
            style='MatrixHeader.TLabel',
            width=10
        )
        corner_label.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        
        # 2. Заголовки столбцов (длины) - ЦЕНТРИРОВАННЫЕ
        for col_idx, length in enumerate(self.lengths, start=1):
            label = ttk.Label(
                self.matrix_frame,
                text=str(length),
                style='MatrixHeader.TLabel',
                width=6
            )
            label.grid(row=0, column=col_idx, sticky="nsew", padx=1, pady=1)
            
        # 3. Заголовки строк (высоты) - ЦЕНТРИРОВАННЫЕ
        for row_idx, height in enumerate(self.heights, start=1):
            # Заголовок строки - ЦЕНТРИРОВАННЫЙ
            label = ttk.Label(
                self.matrix_frame,
                text=str(height),
                style='MatrixHeader.TLabel',
                width=10
            )
            label.grid(row=row_idx, column=0, sticky="nsew", padx=1, pady=1)
            
            # Ячейки с кнопками
            for col_idx, length in enumerate(self.lengths, start=1):
                # Текст кнопки
                btn_text = "ВЫБРАТЬ" if calculated_width >= 65 else "ВЫБР" if calculated_width >= 50 else "В"
                
                btn = ttk.Button(
                    self.matrix_frame,
                    text=btn_text,
                    style='MatrixCell.TButton',
                    command=self._create_selection_handler(length, height)
                )
                btn.grid(row=row_idx, column=col_idx, sticky="nsew", padx=1, pady=1)
                
                # Сохраняем ссылку
                self.matrix_cells[(height, length)] = btn
                
                # Подсказка при наведении
                btn.bind("<Enter>", lambda e, l=length, h=height: self._on_cell_hover(l, h))
                btn.bind("<Leave>", lambda e: self._update_preview())
        
        # 4. Настраиваем размеры ячеек
        for col in range(num_columns):
            self.matrix_frame.columnconfigure(col, weight=1, minsize=calculated_width)
        
        # ФИКСИРОВАННАЯ ВЫСОТА
        for row in range(len(self.heights) + 1):
            self.matrix_frame.rowconfigure(row, weight=0, minsize=row_height)
        
        self.logger.debug("Матрица создана успешно")
            
    def _create_three_column_instruction(self, parent) -> None:
        """Создает инструкцию в 3 колонках"""
        # Текст инструкции (разделён на 3 колонки)
        instruction_parts = [
            # Колонка 1
            [
                "1. ВЫБЕРИТЕ ПОДКЛЮЧЕНИЕ:",
                "• VK-правое/левое —",
                "  нижнее подключение",
                "• K-боковое —",
                "  боковое подключение",
                "",
                "2. ВЫБЕРИТЕ ТИП:",
                "• VK-правое: 10, 11, 20,",
                "  21, 22, 30, 33",
                "• VK-левое: 10, 11,",
                "  30, 33",
                "• K-боковое: 10, 11, 20,",
                "  21, 22, 30, 33"
            ],
            # Колонка 2
            [
                "3. ВЫБЕРИТЕ РАЗМЕР:",
                "• Горизонтально —",
                "  ДЛИНА (мм)",
                "• Вертикально —",
                "  ВЫСОТА (мм)",
                "• Нажмите 'ВЫБРАТЬ'",
                "  в нужной ячейке",
                "",
                "4. ХАРАКТЕРИСТИКИ:",
                "• При наведении",
                "  отображаются",
                "• Фактические данные",
                "  LaggarTT",
                "• Если радиатор",
                "  не найден —",
                "  появится надпись"
            ],
            # Колонка 3
            [
                "⚡ СОВЕТЫ:",
                "",
                "• Используйте параметры",
                "  из оригинального",
                "  названия радиатора",
                "",
                "• При отсутствии",
                "  точного размера",
                "  выбирайте БОЛЬШИЙ",
                "  аналог",
                "",
                "• Программа обучается",
                "  на ваших выборах",
                "",
                "• Улучшает",
                "  автоматический",
                "  подбор"
            ]
        ]
        
        # Создаем 3 колонки
        for col_idx, column_lines in enumerate(instruction_parts):
            column_frame = ttk.Frame(parent)
            column_frame.pack(side="left", fill="both", expand=True, 
                            padx=(0, 10) if col_idx < 2 else (0, 0))
            
            # Текст виджет для колонки
            text_widget = tk.Text(
                column_frame,
                wrap="word",
                font=("Segoe UI", 9),
                background="#f9f9f9",
                relief="solid",
                borderwidth=1,
                padx=10,
                pady=5,
                height=15,
                state="disabled"
            )
            text_widget.pack(fill="both", expand=True)
            
            # Заполняем колонку
            text_widget.config(state="normal")
            for line in column_lines:
                text_widget.insert("end", line + "\n")
            text_widget.config(state="disabled")
            
    def _find_laggartt_radiator(self, connection: str, rad_type: str, height: int, length: int):
        """
        Ищет радиатор LaggarTT в базе данных по параметрам
        ТОЧНО как в interface_builder.py в методе create_matrix_cell
        """
        self.logger.debug(f"Поиск радиатора: {connection} {rad_type}, {height}×{length}")
        
        # Формируем имя листа как в основном интерфейсе
        sheet_name = f"{connection} {rad_type}"
        self.logger.debug(f"Имя листа: {sheet_name}")
        
        # Проверяем существование листа
        if sheet_name not in self.main_app.sheets:
            self.logger.error(f"Лист не найден в self.main_app.sheets: {sheet_name}")
            self.logger.debug(f"Доступные листы: {list(self.main_app.sheets.keys())[:5]}...")
            return None
        
        # Получаем данные листа
        data = self.main_app.sheets[sheet_name]
        self.logger.debug(f"Размер данных листа {sheet_name}: {data.shape}")
        self.logger.debug(f"Колонки листа: {list(data.columns)}")
        
        # ТОЧНО как в interface_builder.py create_matrix_cell
        pattern = f"/{height}/{length}"
        self.logger.debug(f"Паттерн поиска: '{pattern}'")
        
        # Логируем несколько примеров названий для проверки
        sample_names = data['Наименование'].head(3).tolist() if 'Наименование' in data.columns else []
        self.logger.debug(f"Примеры названий в листе: {sample_names}")
        
        # Используем точно такой же поиск
        try:
            match = data[data['Наименование'].str.contains(pattern, na=False)]
            self.logger.debug(f"Найдено совпадений: {len(match)}")
            
            if not match.empty:
                product = match.iloc[0]
                
                # Логируем ВСЕ найденные данные
                self.logger.info(f"=== НАЙДЕН РАДИАТОР ===")
                self.logger.info(f"Лист: {sheet_name}")
                self.logger.info(f"Размер: {height}×{length}")
                self.logger.info(f"Паттерн: {pattern}")
                
                # Логируем все столбцы
                for col in data.columns:
                    value = product.get(col, 'НЕТ')
                    if pd.isna(value):
                        value_str = "NaN"
                    else:
                        value_str = str(value)
                    self.logger.info(f"  {col}: {value_str}")
                
                self.logger.info("=" * 40)
                
                return product
            else:
                self.logger.warning(f"Радиатор не найден по паттерну: '{pattern}'")
                # Показываем несколько похожих названий для диагностики
                similar = data['Наименование'].head(5).tolist() if 'Наименование' in data.columns else []
                self.logger.debug(f"Похожие названия в листе: {similar}")
                
        except Exception as e:
            self.logger.error(f"Ошибка при поиске радиатора: {str(e)}")
            import traceback
            self.logger.error(traceback.format_exc())
        
        return None
        
    def _on_window_resize(self, event=None) -> None:
        """Обработчик изменения размера окна"""
        if self.matrix_frame and event and event.widget == self._selector_window:
            if hasattr(self, '_resize_timer'):
                self._selector_window.after_cancel(self._resize_timer)
            self._resize_timer = self._selector_window.after(200, self._create_adaptive_matrix)
            
    def _on_connection_changed(self) -> None:
        """Обработчик изменения подключения"""
        self.logger.debug(f"Изменено подключение на: {self.selected_connection.get()}")
        self._update_available_types()
        self._update_preview()
        
    def _update_available_types(self) -> None:
        """Обновляет доступные типы радиаторов"""
        self.logger.debug("Обновление доступных типов")
        for widget in self.step2_frame.winfo_children():
            widget.destroy()
            
        if self.selected_connection.get() == "VK-левое":
            available_types = ["10", "11", "30", "33"]
        else:
            available_types = ["10", "11", "20", "21", "22", "30", "33"]
        
        self.logger.debug(f"Доступные типы для {self.selected_connection.get()}: {available_types}")
            
        # Радиокнопки в ОДИН РЯД
        for rad_type in available_types:
            rb = ttk.Radiobutton(
                self.step2_frame,
                text=f"Тип {rad_type}",
                variable=self.selected_type,
                value=rad_type,
                command=self._update_preview
            )
            rb.pack(side="left", padx=5, pady=2)
            
    def _on_cell_hover(self, length: int, height: int) -> None:
        """Обработчик наведения на ячейку матрицы"""
        self.logger.debug(f"Наведение на ячейку: {height}×{length}")
        self.preview_length = length
        self.preview_height = height
        self._update_preview()
        
    def _update_preview(self) -> None:
        """Обновляет предпросмотр с реальными данными LaggarTT"""
        self.logger.debug("=== ОБНОВЛЕНИЕ ПРЕДПРОСМОТРА ===")
        
        # Сбрасываем при отсутствии параметров
        if not all([self.selected_connection.get(), 
                    self.selected_type.get(), 
                    self.preview_height, 
                    self.preview_length]):
            self.logger.debug("Не все параметры выбраны, сброс")
            self.analog_name_var.set("Выберите параметры")
            self.article_var.set("—")
            self.power_var.set("—")
            self.weight_var.set("—")
            self.volume_var.set("—")
            return
        
        conn = self.selected_connection.get()
        rad_type = self.selected_type.get()
        height = self.preview_height
        length = self.preview_length
        
        self.logger.debug(f"Параметры: {conn} {rad_type}, {height}×{length}")
        
        # Ищем реальный радиатор LaggarTT
        radiator = self._find_laggartt_radiator(conn, rad_type, height, length)
        
        if radiator is not None:
            # Формируем название аналога
            analog_name = f"LaggarTT {conn} {rad_type}"
            
            # Безопасно получаем данные с проверкой NaN
            def get_safe_value(column_name, default="—"):
                if column_name not in radiator:
                    self.logger.warning(f"Столбец '{column_name}' не найден в данных радиатора")
                    return default
                
                value = radiator[column_name]
                self.logger.debug(f"Столбец '{column_name}': сырое значение = {repr(value)}, тип = {type(value)}")
                
                # Проверяем разные типы пустых значений
                if pd.isna(value):
                    self.logger.debug(f"Столбец '{column_name}': значение NaN")
                    return default
                if value is None:
                    self.logger.debug(f"Столбец '{column_name}': значение None")
                    return default
                if str(value).strip() == '':
                    self.logger.debug(f"Столбец '{column_name}': пустая строка")
                    return default
                if str(value).strip().lower() in ['nan', 'none', 'null', 'nat']:
                    self.logger.debug(f"Столбец '{column_name}': строковое значение NaN")
                    return default
                
                # Преобразуем числа к читаемому виду
                if isinstance(value, (int, float, np.integer, np.floating)):
                    self.logger.debug(f"Столбец '{column_name}': числовое значение {value}")
                    # Для веса и объема убираем лишние нули
                    if column_name in ['Вес, кг', 'Объем, м3']:
                        result = f"{value:.2f}".rstrip('0').rstrip('.')
                        self.logger.debug(f"Столбец '{column_name}': отформатировано как {result}")
                        return result
                    # Для мощности целое число
                    elif column_name == 'Мощность, Вт':
                        result = str(int(value))
                        self.logger.debug(f"Столбец '{column_name}': отформатировано как {result}")
                        return result
                    else:
                        result = str(value)
                        self.logger.debug(f"Столбец '{column_name}': преобразовано в строку {result}")
                        return result
                
                result = str(value).strip()
                self.logger.debug(f"Столбец '{column_name}': строковое значение {result}")
                return result
            
            # Получаем значения
            article = get_safe_value('Артикул', "—")
            power = get_safe_value('Мощность, Вт', "—")
            weight = get_safe_value('Вес, кг', "—")
            volume = get_safe_value('Объем, м3', "—")
            
            self.logger.info(f"Отображаемые значения:")
            self.logger.info(f"  Артикул: {article}")
            self.logger.info(f"  Мощность: {power}")
            self.logger.info(f"  Вес: {weight}")
            self.logger.info(f"  Объем: {volume}")
            
            # Обновляем интерфейс
            self.analog_name_var.set(f"{analog_name} {height}×{length}")
            self.article_var.set(article)
            self.power_var.set(f"{power} Вт" if power != "—" else "—")
            self.weight_var.set(f"{weight} кг" if weight != "—" else "—")
            self.volume_var.set(f"{volume} м³" if volume != "—" else "—")
            
            self.logger.info("Предпросмотр обновлен с реальными данными")
        else:
            # Радиатор не найден в ассортименте
            self.logger.warning(f"Радиатор не найден: {conn} {rad_type} {height}×{length}")
            self.analog_name_var.set(f"LaggarTT {conn} {rad_type}")
            self.article_var.set("—")
            self.power_var.set("Не найден в ассортименте")
            self.weight_var.set("—")
            self.volume_var.set("—")
        
        self.logger.debug("=== ЗАВЕРШЕНО ОБНОВЛЕНИЕ ПРЕДПРОСМОТРА ===\n")
        
    def _create_selection_handler(self, length: int, height: int):
        """Создает обработчик для кнопки выбора"""
        def handler():
            self.logger.info(f"=== ВЫБОР РАДИАТОРА: {height}×{length} ===")
            # Сохраняем выбранные параметры
            self.preview_length = length
            self.preview_height = height
            
            # Находим радиатор для подтверждения
            radiator = self._find_laggartt_radiator(
                self.selected_connection.get(),
                self.selected_type.get(),
                height,
                length
            )
            
            if radiator is None:
                messagebox.showwarning(
                    "Радиатор не найден",
                    f"Радиатор LaggarTT {self.selected_connection.get()} {self.selected_type.get()} "
                    f"размером {height}×{length} не найден в ассортименте.\n"
                    f"Пожалуйста, выберите другой размер."
                )
                return
            
            # Получаем артикул для сообщения
            article = "—"
            if 'Артикул' in radiator and not pd.isna(radiator['Артикул']):
                article = str(radiator['Артикул']).strip()
            
            self.logger.info(f"Выбран радиатор: {article}")
            
            # Вызываем метод из основного приложения
            if hasattr(self.main_app, '_set_meteor_analog_and_save_pattern'):
                success = self.main_app._set_meteor_analog_and_save_pattern(
                    self.tree_reference, 
                    self.current_item, 
                    self.original_name, 
                    self.selected_connection.get(),
                    self.selected_type.get(), 
                    length, 
                    height, 
                    self._selector_window
                )
                
                if success:
                    # Применяем паттерны к таблице
                    if hasattr(self.main_app, 'apply_saved_patterns_to_table'):
                        updated_count = self.main_app.apply_saved_patterns_to_table(self.tree_reference)
                        
                        if updated_count > 0:
                            messagebox.showinfo(
                                "Аналог выбран",
                                f"Для '{self.original_name}' выбран аналог LaggarTT.\n"
                                f"Артикул: {article}\n"
                                f"Выбор сохранён в базе данных.\n"
                                f"Автоматически подобрано ещё {updated_count} радиаторов."
                            )
                        else:
                            messagebox.showinfo(
                                "Аналог выбран", 
                                f"Для '{self.original_name}' выбран аналог LaggarTT.\n"
                                f"Артикул: {article}\n"
                                f"Выбор сохранён в базе данных."
                            )
                    
                    # Закрываем окно
                    self._selector_window.destroy()
        
        return handler
        
    def _on_close(self) -> None:
        """Обработчик закрытия окна"""
        self.logger.debug("Закрытие окна селектора")
        if self.close_callback:
            self.close_callback()
        
        if self._selector_window:
            self._selector_window.destroy()