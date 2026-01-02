import tkinter as tk
from tkinter import ttk
import platform
import re
import threading
from typing import List, Optional, Set, Dict, Any
import traceback
import queue

# Попытка импорта библиотек для работы с PDF и изображениями
try:
    import fitz  # PyMuPDF
    from PIL import Image, ImageTk
    from io import BytesIO
    HAS_PDF_LIBS = True
except ImportError:
    print("[WARNING] Библиотеки для миниатюр PDF не установлены")
    HAS_PDF_LIBS = False


class ProgressDialog:
    """Диалог прогресса для загрузки миниатюр"""
    def __init__(self, parent, title, message):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("400x120")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Центрирование
        self.dialog.update_idletasks()
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = (self.dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (height // 2)
        self.dialog.geometry(f"{width}x{height}+{x}+{y}")
        
        # Сообщение
        self.message_label = ttk.Label(self.dialog, text=message)
        self.message_label.pack(pady=10)
        
        # Прогресс-бар
        self.progress_bar = ttk.Progressbar(
            self.dialog, 
            mode='determinate', 
            length=350
        )
        self.progress_bar.pack(pady=10)
        
        # Проценты
        self.percent_label = ttk.Label(self.dialog, text="0%")
        self.percent_label.pack(pady=5)
        
        self.dialog.update()
    
    def update(self, percent, message=None):
        """Обновляет прогресс"""
        if message:
            self.message_label.config(text=message)
        self.progress_bar['value'] = percent
        self.percent_label.config(text=f"{percent}%")
        self.dialog.update_idletasks()
    
    def close(self):
        """Закрывает диалог"""
        if self.dialog and self.dialog.winfo_exists():
            self.dialog.destroy()


class PdfPageSelector:
    """
    Окно для выбора страниц PDF с миниатюрами и ручным вводом диапазонов
    """
    
    def __init__(self, parent, total_pages: int, pdf_path: str, 
                 title: str = "Выбор страниц PDF", initial_selection: List[int] = None):
        """
        Инициализация окна выбора страниц
        
        Args:
            parent: родительское окно
            total_pages: общее количество страниц в PDF
            pdf_path: путь к PDF файлу
            title: заголовок окна
            initial_selection: начальный выбор страниц (по умолчанию пустой)
        """
        self.parent = parent
        self.total_pages = total_pages
        self.pdf_path = pdf_path
        self.title = title
        self.initial_selection = initial_selection or []
        
        # Результат
        self.result = {"pages": None, "confirmed": False}
        
        # Выбранные страницы
        self.selected_pages: Set[int] = set(self.initial_selection)
        
        # Миниатюры
        self.thumbnail_images: Dict[int, Any] = {}  # page_num -> PhotoImage
        self.thumbnail_widgets: Dict[int, Dict[str, Any]] = {}  # page_num -> widget_dict
        
        # Окно
        self.dialog = None
        self.thumbnail_progress_label = None
        
        # ID привязки событий для последующей очистки
        self.mousewheel_id = None
        
        # Элементы интерфейса
        self.range_entry = None
        self.range_var = None
        self.stats_label = None
        self.canvas = None
        self.thumbnails_frame = None
        
        # Таймер для автоматического применения
        self._apply_timer = None
    
    def show(self) -> Optional[List[int]]:
        """
        Показывает диалог выбора страниц
        
        Returns:
            List[int] или None: список выбранных страниц или None если отмена
        """
        self._create_dialog()
        self._initialize_dialog()
        self.dialog.wait_window()
        
        # Очистка ресурсов
        self._cleanup()
        
        if self.result["confirmed"]:
            print(f"[INFO] Выбрано {len(self.result['pages'])} страниц для обработки")
            return self.result["pages"]
        return None
    
    def _create_dialog(self):
        """Создает диалоговое окно"""
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title(f"{self.title} (всего: {self.total_pages})")
        
        # Делаем дочерним окном
        self.dialog.transient(self.parent)
        
        # Устанавливаем окно на весь экран, но не перекрывая панель задач
        self._set_fullscreen_geometry()
        
        # Делаем окно модальным
        self.dialog.grab_set()
        self.dialog.focus_set()
    
    def _set_fullscreen_geometry(self):
        """Устанавливает окно на весь экран, но не перекрывая панель задач"""
        screen_width = self.dialog.winfo_screenwidth()
        screen_height = self.dialog.winfo_screenheight()
        
        # Предполагаем высоту панели задач около 40px + 30px для подъема кнопок
        taskbar_height = 70
        dialog_height = screen_height - taskbar_height
        
        # Центрируем по горизонтали
        dialog_width = screen_width
        x = 0
        y = 0  # Начинаем от верхнего края
        
        self.dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        
        # Устанавливаем минимальный размер
        self.dialog.minsize(1000, 600)
    
    def _create_widgets(self):
        """Создает все виджеты окна"""
        # Главный фрейм
        main_frame = ttk.Frame(self.dialog)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Верхняя часть - заголовок и инструкция
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill="x", pady=(0, 15))
        
        ttk.Label(top_frame, text=f"PDF содержит {self.total_pages} страниц", 
                  font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 8))
        
        instruction_text = (
            "Выберите страницы с таблицами (клик по миниатюре или ввод диапазона). "
            "Чертежи, схемы и приложения пропускайте."
        )
        ttk.Label(top_frame, text=instruction_text,
                  wraplength=1400, font=("Segoe UI", 10)).pack(anchor="w", pady=(0, 15))
        
        # Фрейм для ручного ввода диапазона
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill="x", pady=(0, 15))
        
        ttk.Label(input_frame, text="Введите диапазон страниц:", 
                  font=("Segoe UI", 10)).pack(side="left", padx=(0, 10))
        
        # Поле для ручного ввода с StringVar для отслеживания изменений
        self.range_var = tk.StringVar()
        self.range_entry = ttk.Entry(input_frame, textvariable=self.range_var, 
                                     width=50, font=("Segoe UI", 10))
        self.range_entry.pack(side="left", padx=(0, 10))
        
        # Автоматическое применение при:
        # 1. Нажатии Enter
        self.range_entry.bind("<Return>", lambda e: self._apply_range_from_entry())
        
        # 2. Потере фокуса (пользователь закончил ввод)
        self.range_entry.bind("<FocusOut>", lambda e: self._apply_range_from_entry())
        
        # 3. Через 1 секунду после остановки ввода
        def on_text_change(*args):
            # Отменяем предыдущий таймер
            if self._apply_timer:
                self.dialog.after_cancel(self._apply_timer)
            
            # Устанавливаем новый таймер
            self._apply_timer = self.dialog.after(1000, self._apply_range_from_entry)
        
        # Отслеживаем изменения текста
        self.range_var.trace("w", on_text_change)
        
        # Пример формата
        format_label = ttk.Label(input_frame, 
                                 text="Формат: 1-5, 7, 10-15",
                                 font=("Segoe UI", 9), foreground="gray")
        format_label.pack(side="left", padx=(20, 0))
        
        # Средняя часть - фрейм с миниатюрами
        preview_container = ttk.Frame(main_frame)
        preview_container.pack(fill="both", expand=True, pady=(0, 15))
        
        # Заголовок для секции миниатюр
        preview_header_frame = ttk.Frame(preview_container)
        preview_header_frame.pack(fill="x", pady=(0, 10))
        
        preview_header = ttk.Label(preview_header_frame, text="Миниатюры страниц", 
                                   font=("Segoe UI", 10, "bold"))
        preview_header.pack(side="left")
        
        # Область для информации о прогрессе
        self.thumbnail_progress_label = ttk.Label(preview_header_frame, text="", 
                                                  font=("Segoe UI", 9))
        self.thumbnail_progress_label.pack(side="right", padx=(0, 10))
        
        # Пустой фрейм для заполнения пространства
        ttk.Frame(preview_header_frame).pack(side="left", expand=True, fill="x")
        
        # Фрейм для миниатюр с бордюром
        preview_frame = ttk.Frame(preview_container, relief="solid", borderwidth=1)
        preview_frame.pack(fill="both", expand=True, padx=0, pady=0)
        
        # Canvas для миниатюр с прокруткой
        canvas_frame = ttk.Frame(preview_frame)
        canvas_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.canvas = tk.Canvas(canvas_frame, bg="#f0f0f0", highlightthickness=0)
        vsb = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        hsb = ttk.Scrollbar(canvas_frame, orient="horizontal", command=self.canvas.xview)
        
        self.canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Размещаем элементы
        self.canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        canvas_frame.grid_rowconfigure(0, weight=1)
        canvas_frame.grid_columnconfigure(0, weight=1)
        
        # Внутренний фрейм для миниатюр
        self.thumbnails_frame = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.thumbnails_frame, anchor="nw")
        
        # Нижняя часть - статистика и кнопки (поднимаем на 30px выше)
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(side="bottom", fill="x", pady=(30, 30))  # Увеличили нижний отступ
        
        # Разделитель
        ttk.Separator(bottom_frame, orient="horizontal").pack(fill="x", pady=(0, 15))
        
        # Статистика
        stats_frame = ttk.Frame(bottom_frame)
        stats_frame.pack(fill="x", pady=(0, 15))
        
        self.stats_label = ttk.Label(stats_frame, 
                                     text=f"Выбрано страниц: 0 из {self.total_pages}",
                                     font=("Segoe UI", 10, "bold"))
        self.stats_label.pack()
        
        # Кнопки действий
        button_frame = ttk.Frame(bottom_frame)
        button_frame.pack(fill="x", pady=(0, 5))
        
        # Создаем кнопки
        self._create_action_buttons(button_frame)
        
        # Настраиваем вес строк/колонок
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=0)  # Верхняя часть
        main_frame.rowconfigure(1, weight=1)  # Средняя часть (растягивается)
        main_frame.rowconfigure(2, weight=0)  # Нижняя часть
        
        # Обработка Enter/Escape
        self.dialog.bind('<Return>', lambda e: self._confirm_selection())
        self.dialog.bind('<Escape>', lambda e: self._cancel())
        
        # Обработчик закрытия окна
        self.dialog.protocol("WM_DELETE_WINDOW", self._cancel)
    
    def _create_action_buttons(self, parent_frame):
        """Создает кнопки действий"""
        # 1. Выбрать все
        btn_select_all = ttk.Button(parent_frame, text="Выбрать все", 
                                    command=self._select_all, width=18)
        btn_select_all.pack(side="left", padx=(40, 10), expand=True)
        
        # 2. Очистить все
        btn_clear_all = ttk.Button(parent_frame, text="Очистить все", 
                                   command=self._clear_all, width=18)
        btn_clear_all.pack(side="left", padx=10, expand=True)
        
        # 3. Обработать выбранные
        btn_process = ttk.Button(parent_frame, text="Обработать выбранные", 
                                 command=self._confirm_selection, width=22)
        btn_process.pack(side="left", padx=10, expand=True)
        
        # 4. Отмена
        btn_cancel = ttk.Button(parent_frame, text="Отмена", 
                                command=self._cancel, width=18)
        btn_cancel.pack(side="left", padx=(10, 40), expand=True)
    
    def _initialize_dialog(self):
        """Инициализирует диалог с загрузкой миниатюр"""
        # Сначала создаем все виджеты
        self._create_widgets()
        
        # Показываем сообщение о загрузке
        loading_label = ttk.Label(self.dialog, text="Загрузка миниатюр...", 
                                  font=("Segoe UI", 12, "bold"))
        loading_label.place(relx=0.5, rely=0.5, anchor="center")
        self.dialog.update()
        
        # Загружаем миниатюры в отдельном потоке
        self._load_thumbnails_async(loading_label)
    
    def _load_thumbnails_async(self, loading_label):
        """Загружает миниатюры в отдельном потоке"""
        result_queue = queue.Queue()
        
        def thumbnail_worker():
            try:
                success = self._load_thumbnails()
                result_queue.put(("success", success))
            except Exception as e:
                result_queue.put(("error", str(e)))
        
        def check_result():
            try:
                status, data = result_queue.get_nowait()
                
                # Убираем сообщение о загрузке
                loading_label.destroy()
                
                if status == "success":
                    print(f"[INFO] Загружено {len(self.thumbnail_images)} миниатюр")
                    # Инициализируем отображение (ничего не выбираем по умолчанию)
                    self._create_thumbnail_widgets()
                    
                    # Настраиваем прокрутку колесиком мыши
                    self._setup_mousewheel_scroll()
                    
                    # Обновляем canvas
                    self.dialog.after(100, lambda: self.canvas.config(
                        scrollregion=self.canvas.bbox("all")))
                    
                    # Устанавливаем фокус на диалог
                    self.dialog.after(200, lambda: self.dialog.focus_set())
                else:
                    print(f"[ERROR] Ошибка загрузки миниатюр: {data}")
                    # Заполняем заглушками
                    for page_num in range(1, self.total_pages + 1):
                        self.thumbnail_images[page_num] = None
                    self._create_thumbnail_widgets()
                    
            except queue.Empty:
                self.dialog.after(100, check_result)
        
        # Запускаем поток
        thread = threading.Thread(target=thumbnail_worker, daemon=True)
        thread.start()
        
        # Запускаем проверку результата
        self.dialog.after(100, check_result)
    
    def _load_thumbnails(self) -> bool:
        """Загружает миниатюры страниц PDF"""
        if not HAS_PDF_LIBS:
            print("[INFO] Используются заглушки для миниатюр")
            return False
        
        # Создаем прогресс-бар
        progress_dialog = ProgressDialog(
            self.dialog, "Загрузка миниатюр", 
            f"Генерация миниатюр для {self.total_pages} страниц..."
        )
        
        try:
            # Открываем PDF
            pdf_document = fitz.open(self.pdf_path)
            
            for page_num in range(1, self.total_pages + 1):
                try:
                    # Обновляем прогресс
                    percent = int((page_num / self.total_pages) * 100)
                    progress_dialog.update(
                        percent,
                        f"Страница {page_num} из {self.total_pages}"
                    )
                    
                    # Обновляем информацию в интерфейсе
                    self.thumbnail_progress_label.config(
                        text=f"Загружено: {page_num}/{self.total_pages} страниц"
                    )
                    self.dialog.update_idletasks()
                    
                    # Получаем страницу (0-based индекс)
                    page = pdf_document[page_num - 1]
                    
                    # Рендерим страницу в изображение с низким разрешением
                    pix = page.get_pixmap(matrix=fitz.Matrix(0.25, 0.25))
                    
                    # Конвертируем в PIL Image
                    img_data = pix.tobytes("ppm")
                    img_pil = Image.open(BytesIO(img_data))
                    
                    # Изменяем размер для миниатюры
                    thumb_width = 140
                    thumb_height = 180
                    img_pil.thumbnail((thumb_width, thumb_height), Image.Resampling.LANCZOS)
                    
                    # Конвертируем в PhotoImage для Tkinter
                    photo_img = ImageTk.PhotoImage(img_pil)
                    self.thumbnail_images[page_num] = photo_img
                    
                except Exception as e:
                    print(f"[ERROR] Ошибка загрузки миниатюры страницы {page_num}: {e}")
                    self.thumbnail_images[page_num] = None
            
            pdf_document.close()
            progress_dialog.close()
            
            # Очищаем метку прогресса после завершения
            self.thumbnail_progress_label.config(text="")
            return True
            
        except Exception as e:
            print(f"[ERROR] Ошибка загрузки миниатюр: {e}")
            if progress_dialog:
                progress_dialog.close()
            self.thumbnail_progress_label.config(text="Ошибка загрузки")
            return False
    
    def _setup_mousewheel_scroll(self):
        """Настраивает прокрутку колесиком мыши"""
        def on_mousewheel(event):
            if self.canvas and self.canvas.winfo_exists():
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        # Привязываем событие
        self.mousewheel_id = self.canvas.bind_all("<MouseWheel>", on_mousewheel)
    
    def _create_thumbnail_widgets(self):
        """Создает все виджеты миниатюр один раз (без мигания)"""
        # Рассчитываем количество колонок
        cols = 10  # Увеличиваем количество колонок для полноэкранного режима
        
        for i, page_num in enumerate(range(1, self.total_pages + 1)):
            # Создаем виджет миниатюры
            thumb_data = self._create_single_thumbnail(page_num)
            
            # Размещаем в сетке
            row = i // cols
            col = i % cols
            thumb_data["frame"].grid(row=row, column=col, padx=10, pady=10, sticky="nw")
            
            # Сохраняем все элементы виджета
            self.thumbnail_widgets[page_num] = thumb_data
        
        # Обновляем размер canvas
        self.thumbnails_frame.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox("all"))
        
        # Обновляем статистику
        self._update_stats()
    
    def _create_single_thumbnail(self, page_num):
        """Создает один виджет миниатюры страницы и возвращает все его элементы"""
        is_selected = page_num in self.selected_pages
        
        # Основной фрейм
        thumb_frame = ttk.Frame(self.thumbnails_frame, relief="solid", borderwidth=2)
        
        # Номер страницы
        page_label = ttk.Label(thumb_frame, text=f"Страница {page_num}", 
                               font=("Segoe UI", 9, "bold"))
        page_label.pack(pady=(8, 0))
        
        # Область миниатюры (Canvas)
        thumb_canvas = tk.Canvas(thumb_frame, width=140, height=180,
                                 bg="white", highlightthickness=0)
        thumb_canvas.pack(pady=8)
        
        # Загружаем миниатюру или создаем заглушку
        if page_num in self.thumbnail_images and self.thumbnail_images[page_num] is not None:
            # Реальная миниатюра
            photo_img = self.thumbnail_images[page_num]
            thumb_canvas.create_image(70, 90, image=photo_img)
            
            # Сохраняем ссылку на изображение
            thumb_canvas.image = photo_img
        else:
            # Заглушка с цветом в зависимости от типа контента
            if page_num % 3 == 0:
                bg_color = "#ffe6e6"  # Светло-красный для чертежей
                thumb_text = "📐 Чертеж"
            elif page_num % 4 == 0:
                bg_color = "#fff0e6"  # Светло-оранжевый для приложений
                thumb_text = "📋 Приложение"
            else:
                bg_color = "#e6ffe6"  # Светло-зеленый для таблиц
                thumb_text = "📊 Таблица"
            
            thumb_canvas.config(bg=bg_color)
            thumb_canvas.create_text(70, 60, text=thumb_text, 
                                     font=("Segoe UI", 10, "bold"), fill="#333333")
            thumb_canvas.create_text(70, 90, text=f"Страница {page_num}", 
                                     font=("Segoe UI", 9), fill="#666666")
        
        # Индикатор выбора (прямоугольник и галочка)
        selection_rect = thumb_canvas.create_rectangle(5, 5, 135, 175, 
                                                      outline="#cccccc", width=1)
        checkmark = thumb_canvas.create_text(130, 20, text="", 
                                           font=("Segoe UI", 14, "bold"), fill="#0066cc")
        
        # Статус
        status_text = "⏸ Пропустить" if not is_selected else "✅ Выбрана для обработки"
        status_label = ttk.Label(thumb_frame, text=status_text, 
                                 font=("Segoe UI", 9))
        status_label.pack(pady=(0, 8))
        
        # Обработчик клика
        def on_click(event, p=page_num):
            self._toggle_page_selection_smooth(p)
        
        thumb_frame.bind("<Button-1>", on_click)
        thumb_canvas.bind("<Button-1>", on_click)
        page_label.bind("<Button-1>", on_click)
        status_label.bind("<Button-1>", on_click)
        
        # Обновляем визуальное состояние
        self._update_thumbnail_visual(thumb_canvas, selection_rect, checkmark, status_label, is_selected)
        
        # Возвращаем все элементы виджета
        return {
            "frame": thumb_frame,
            "canvas": thumb_canvas,
            "selection_rect": selection_rect,
            "checkmark": checkmark,
            "status_label": status_label,
            "page_num": page_num
        }
    
    def _update_thumbnail_visual(self, canvas, selection_rect, checkmark, status_label, is_selected):
        """Обновляет визуальное состояние одной миниатюры (без пересоздания)"""
        if is_selected:
            canvas.itemconfig(selection_rect, outline="#0066cc", width=3)
            canvas.itemconfig(checkmark, text="✓")
            status_label.config(text="✅ Выбрана\nдля обработки")
        else:
            canvas.itemconfig(selection_rect, outline="#cccccc", width=1)
            canvas.itemconfig(checkmark, text="")
            status_label.config(text="⏸ Пропустить")
    
    def _toggle_page_selection_smooth(self, page_num):
        """Переключает выбор страницы без мигания"""
        if page_num in self.selected_pages:
            self.selected_pages.remove(page_num)
        else:
            self.selected_pages.add(page_num)
        
        # Обновляем только визуальное состояние этой миниатюры
        if page_num in self.thumbnail_widgets:
            thumb_data = self.thumbnail_widgets[page_num]
            is_selected = page_num in self.selected_pages
            self._update_thumbnail_visual(
                thumb_data["canvas"],
                thumb_data["selection_rect"],
                thumb_data["checkmark"],
                thumb_data["status_label"],
                is_selected
            )
        
        # Обновляем статистику и поле ввода
        self._update_stats()
        self._update_range_entry()
    
    def _toggle_page_selection(self, page_num):
        """Алиас для обратной совместимости"""
        self._toggle_page_selection_smooth(page_num)
    
    def _update_stats(self):
        """Обновляет статистику выбора"""
        if self.stats_label:
            self.stats_label.config(
                text=f"Выбрано страниц: {len(self.selected_pages)} из {self.total_pages}"
            )
    
    def _update_range_entry(self):
        """Обновляет поле ввода на основе выбранных страниц"""
        if not self.range_entry:
            return
        
        if not self.selected_pages:
            self.range_entry.delete(0, tk.END)
            return
        
        # Сортируем страницы
        sorted_pages = sorted(list(self.selected_pages))
        
        # Формируем диапазоны
        ranges = []
        start = sorted_pages[0]
        end = start
        
        for page in sorted_pages[1:]:
            if page == end + 1:
                end = page
            else:
                if start == end:
                    ranges.append(str(start))
                else:
                    ranges.append(f"{start}-{end}")
                start = page
                end = page
        
        # Добавляем последний диапазон
        if start == end:
            ranges.append(str(start))
        else:
            ranges.append(f"{start}-{end}")
        
        # Обновляем поле ввода
        self.range_entry.delete(0, tk.END)
        self.range_entry.insert(0, ", ".join(ranges))
    
    def _apply_range_from_entry(self):
        """Применяет диапазон из поля ввода (автоматически)"""
        if not self.range_entry:
            return
        
        # Отменяем таймер, если он активен
        if self._apply_timer:
            self.dialog.after_cancel(self._apply_timer)
            self._apply_timer = None
        
        range_str = self.range_entry.get().strip()
        
        # Если поле пустое - очищаем выбор
        if not range_str:
            self._clear_all()
            return
        
        # Парсим диапазон
        pages = self._parse_page_range(range_str)
        
        if pages is None:
            # Некорректный формат - просто игнорируем
            return
        
        # Устанавливаем выбранные страницы
        new_selection = set(pages)
        
        # Если выбор не изменился - ничего не делаем
        if new_selection == self.selected_pages:
            return
        
        self.selected_pages = new_selection
        
        # Обновляем все миниатюры (но без пересоздания)
        for page_num, thumb_data in self.thumbnail_widgets.items():
            is_selected = page_num in self.selected_pages
            self._update_thumbnail_visual(
                thumb_data["canvas"],
                thumb_data["selection_rect"],
                thumb_data["checkmark"],
                thumb_data["status_label"],
                is_selected
            )
        
        # Обновляем статистику
        self._update_stats()
    
    def _parse_page_range(self, range_str: str) -> Optional[List[int]]:
        """
        Парсит строку диапазона страниц (например: "1-5, 7, 9-12")
        
        Returns:
            List[int] или None: список номеров страниц или None при ошибке
        """
        pages = []
        parts = range_str.split(',')
        
        for part in parts:
            part = part.strip()
            if not part:
                continue
                
            if '-' in part:
                # Диапазон
                try:
                    start, end = part.split('-')
                    start = int(start.strip())
                    end = int(end.strip())
                    
                    if 1 <= start <= self.total_pages and 1 <= end <= self.total_pages and start <= end:
                        pages.extend(range(start, end + 1))
                    else:
                        return None
                except (ValueError, TypeError):
                    return None
            else:
                # Одиночная страница
                try:
                    page = int(part)
                    if 1 <= page <= self.total_pages:
                        pages.append(page)
                    else:
                        return None
                except (ValueError, TypeError):
                    return None
        
        # Удаляем дубликаты и сортируем
        return sorted(list(set(pages)))
    
    def _select_all(self):
        """Выбрать все страницы"""
        self.selected_pages = set(range(1, self.total_pages + 1))
        
        # Обновляем все миниатюры
        for page_num, thumb_data in self.thumbnail_widgets.items():
            self._update_thumbnail_visual(
                thumb_data["canvas"],
                thumb_data["selection_rect"],
                thumb_data["checkmark"],
                thumb_data["status_label"],
                True
            )
        
        self._update_stats()
        self._update_range_entry()
    
    def _clear_all(self):
        """Очистить все выборы"""
        self.selected_pages.clear()
        
        # Обновляем все миниатюры
        for page_num, thumb_data in self.thumbnail_widgets.items():
            self._update_thumbnail_visual(
                thumb_data["canvas"],
                thumb_data["selection_rect"],
                thumb_data["checkmark"],
                thumb_data["status_label"],
                False
            )
        
        self._update_stats()
        self._update_range_entry()
    
    def _confirm_selection(self):
        """Обработать выбранные страницы"""
        if not self.selected_pages:
            # Показываем только это одно предупреждение
            from tkinter import messagebox
            messagebox.showwarning(
                "Предупреждение", 
                "Вы не выбрали ни одной страницы для обработки.\n"
                "Нажмите на миниатюры таблиц, введите диапазон или нажмите 'Выбрать все'."
            )
            return
        
        self.result["pages"] = sorted(list(self.selected_pages))
        self.result["confirmed"] = True
        self.dialog.destroy()
    
    def _cancel(self):
        """Отмена выбора"""
        self.result["confirmed"] = False
        self.dialog.destroy()
    
    def _cleanup(self):
        """Очистка ресурсов"""
        # Очищаем миниатюры из памяти
        self.thumbnail_images.clear()
        self.thumbnail_widgets.clear()
        
        # Отменяем таймер, если он активен
        if self._apply_timer:
            self.dialog.after_cancel(self._apply_timer)
            self._apply_timer = None
        
        # Удаляем привязку события прокрутки
        if self.mousewheel_id and self.canvas:
            try:
                self.canvas.unbind_all("<MouseWheel>")
            except:
                pass