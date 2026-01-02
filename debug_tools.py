import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
import os
import re
import json
from datetime import datetime

class DebugTools:
    def __init__(self, root, pattern_manager):
        self.root = root
        self.pattern_manager = pattern_manager
        
    def show_pattern_debug_window(self):
        """Окно отладки шаблонов"""
        debug_window = tk.Toplevel(self.root)
        debug_window.title("Отладка шаблонов")
        debug_window.geometry("1200x700")
        
        main_frame = ttk.Frame(debug_window)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill="x", pady=(0, 10))
        
        debug_text = scrolledtext.ScrolledText(
            main_frame,
            wrap="word",
            font=("Consolas", 9)
        )
        debug_text.pack(fill="both", expand=True)
        
        scrollbar = ttk.Scrollbar(debug_text)
        scrollbar.pack(side="right", fill="y")
        debug_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=debug_text.yview)
        
        ttk.Button(control_frame, text="Обновить данные", 
                  command=lambda: self.update_debug_info(debug_text)).pack(side="left", padx=5)
        ttk.Button(control_frame, text="Проверить все шаблоны", 
                  command=lambda: self.test_all_patterns(debug_text)).pack(side="left", padx=5)
        ttk.Button(control_frame, text="Экспорт в файл", 
                  command=lambda: self.export_debug_info(debug_text)).pack(side="left", padx=5)
        
        self.update_debug_info(debug_text)

    def update_debug_info(self, debug_text):
        """Обновление информации в отладочном окне"""
        debug_text.delete(1.0, tk.END)
        
        debug_text.insert(tk.END, "="*80 + "\n")
        debug_text.insert(tk.END, "ИНФОРМАЦИЯ О ФАЙЛЕ ШАБЛОНОВ\n")
        debug_text.insert(tk.END, "="*80 + "\n")
        
        patterns_file = "patterns.json"
        if os.path.exists(patterns_file):
            file_size = os.path.getsize(patterns_file)
            debug_text.insert(tk.END, f"Файл: {patterns_file}\n")
            debug_text.insert(tk.END, f"Размер: {file_size} байт\n")
            debug_text.insert(tk.END, f"Последнее изменение: {datetime.fromtimestamp(os.path.getmtime(patterns_file))}\n")
        else:
            debug_text.insert(tk.END, f"Файл {patterns_file} не существует!\n")
        
        debug_text.insert(tk.END, "\n" + "="*80 + "\n")
        debug_text.insert(tk.END, "ЗАГРУЖЕННЫЕ ШАБЛОНЫ\n")
        debug_text.insert(tk.END, "="*80 + "\n")
        
        patterns = self.pattern_manager.patterns
        debug_text.insert(tk.END, f"Всего шаблонов: {len(patterns)}\n\n")
        
        for i, pattern_data in enumerate(patterns, 1):
            debug_text.insert(tk.END, f"Шаблон #{i}:\n")
            debug_text.insert(tk.END, f"  Regex: {pattern_data.get('pattern', 'НЕТ')}\n")
            debug_text.insert(tk.END, f"  Подключение: {pattern_data.get('connection', 'НЕТ')}\n")
            debug_text.insert(tk.END, f"  Тип: {pattern_data.get('rad_type', 'НЕТ')}\n")
            debug_text.insert(tk.END, f"  Высота: {pattern_data.get('height', 'НЕТ')}\n")
            debug_text.insert(tk.END, f"  Длина: {pattern_data.get('length', 'НЕТ')}\n")
            
            try:
                re.compile(pattern_data.get('pattern', ''))
                debug_text.insert(tk.END, "  Regex: ВАЛИДНЫЙ\n")
            except re.error as e:
                debug_text.insert(tk.END, f"  Regex: ОШИБКА - {e}\n")
            
            debug_text.insert(tk.END, "-"*40 + "\n")

    def test_all_patterns(self, debug_text):
        """Тестирование всех шаблонов"""
        debug_text.insert(tk.END, "\n" + "="*80 + "\n")
        debug_text.insert(tk.END, "ТЕСТИРОВАНИЕ ШАБЛОНОВ\n")
        debug_text.insert(tk.END, "="*80 + "\n")
        debug_text.insert(tk.END, "Функция тестирования шаблонов\n")

    def export_debug_info(self, debug_text):
        """Экспорт отладочной информации в файл"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")],
            title="Сохранить отладочную информацию"
        )
        
        if file_path:
            try:
                content = debug_text.get(1.0, tk.END)
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                debug_text.insert(tk.END, f"\nИнформация сохранена в: {file_path}\n")
            except Exception as e:
                debug_text.insert(tk.END, f"\nОшибка сохранения: {e}\n")