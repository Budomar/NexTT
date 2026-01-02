import json
import re
import os
import sys
from datetime import datetime
from typing import Dict, List, Optional, Any

class PatternManager:
    """Управляет шаблонами для автоматического определения параметров радиаторов конкурентов."""

    def __init__(self, patterns_file: str = "patterns.json"):
        self.patterns_file = patterns_file
        # Определяем путь к файлу РЯДОМ с .exe (или скриптом)
        self.external_file_path = os.path.join(os.path.dirname(sys.argv[0]), self.patterns_file)
        self.patterns: List[Dict[str, Any]] = []
        self.load_patterns()

    def resource_path(self, relative_path):
        """Get absolute path to resource, works for dev and for PyInstaller"""
        try:
            # PyInstaller создаёт временную папку _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def load_patterns(self):
        """Загружает паттерны из внешнего файла (patterns.json рядом с .exe).
        Если файла нет — создаёт его на основе встроенного ресурса."""
        
        # 1. Если внешний файл существует — грузим из него
        if os.path.exists(self.external_file_path):
            try:
                with open(self.external_file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.patterns = self._ensure_list_format(data)
                    print(f"[PATTERNS] Загружено {len(self.patterns)} правил из внешнего файла: {self.external_file_path}")
                    return
            except Exception as e:
                print(f"[ERROR] Ошибка загрузки внешнего {self.external_file_path}: {e}")

        # 2. Если внешнего файла нет — создаём его из ВСТРОЕННОГО ресурса
        try:
            internal_path = self.resource_path("patterns.json")
            if os.path.exists(internal_path):
                # Читаем встроенный файл
                with open(internal_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.patterns = self._ensure_list_format(data)
                # Копируем встроенный файл во внешний
                with open(self.external_file_path, 'w', encoding='utf-8') as f:
                    json.dump(self.patterns, f, ensure_ascii=False, indent=4)
                print(f"[PATTERNS] Создан внешний файл: {self.external_file_path} на основе встроенного")
            else:
                print(f"[PATTERNS] Встроенный файл patterns.json не найден. Создаём пустой внешний файл.")
                self.patterns = []
                with open(self.external_file_path, 'w', encoding='utf-8') as f:
                    json.dump(self.patterns, f, ensure_ascii=False, indent=4)
                    
        except Exception as e:
            print(f"[ERROR] Ошибка инициализации паттернов: {e}")
            self.patterns = []

    def _ensure_list_format(self, data):
        """Преобразует данные в формат списка, даже если пришёл словарь (старый формат)."""
        if isinstance(data, list):
            return data
        elif isinstance(data, dict):
            # Старый формат — конвертируем
            print("[PATTERNS] Конвертация старого формата (dict) в новый (list)...")
            return self._convert_old_format(data)
        else:
            return []        

    def _convert_old_format(self, old_data: Dict[str, Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Конвертирует старый формат (словарь полных названий) в новый (список паттернов)."""
        new_patterns = []
        for original_name, params in old_data.items():
             # Создаем паттерн, который ищет точное совпадение всей строки
             # re.escape экранирует специальные символы
             escaped_name = re.escape(original_name.lower())
             # Заменяем экранированные пробелы на \s+ (один или более пробельных символов)
             # Это делает правило более устойчивым к незначительным изменениям форматирования
             pattern_for_exact_match = escaped_name.replace(r'\ ', r'\s+')
             new_rule = {
                 "pattern": pattern_for_exact_match,
                 "connection": params.get("connection", "VK-правое"),
                 "rad_type": params.get("rad_type", "10"),
                 "height": params.get("height", 500),
                 "length": params.get("length", 1000),
                 "source": "converted" # Метка для отладки
             }
             new_patterns.append(new_rule)
        return new_patterns

    def save_pattern(self, name, connection, rad_type, height, length, art, met_name):
        """Сохраняет новое правило в ВНЕШНИЙ файл patterns.json (рядом с .exe)"""
        try:
            # Создаём паттерн
            new_pattern = {
                "pattern": self._create_smart_pattern(name.lower(), rad_type, height, length),
                "connection": connection,
                "rad_type": rad_type,
                "height": height, 
                "length": length,
                "source": "learned",
                "original_example": name,
                "learned_art": art,
                "learned_meteor_name": met_name,
                "learned_timestamp": datetime.now().isoformat()
            }

            # Загружаем существующие правила из ВНЕШНЕГО файла
            if os.path.exists(self.external_file_path):
                with open(self.external_file_path, 'r', encoding='utf-8') as f:
                    patterns = json.load(f)
                if not isinstance(patterns, list):
                    patterns = []
            else:
                patterns = []

            # Добавляем новый паттерн в начало
            patterns.insert(0, new_pattern)

            # Сохраняем ВО ВНЕШНИЙ ФАЙЛ
            with open(self.external_file_path, 'w', encoding='utf-8') as f:
                json.dump(patterns, f, ensure_ascii=False, indent=4)

            print(f"[SUCCESS] Сохранено новое правило в: {self.external_file_path}")
            self.load_patterns()  # Перезагружаем

        except Exception as e:
            print(f"[ERROR] Ошибка сохранения паттерна: {e}")
            import traceback
            traceback.print_exc()

    def _create_smart_pattern(self, name_lower, rad_type, height, length):
        """Создает умный regex-паттерн на основе названия радиатора."""
        # НОРМАЛИЗУЕМ входную строку: убираем управляющие символы, лишние пробелы
        import re
        # Удаляем \r, \n, \t и заменяем любые пробельные последовательности на один пробел
        normalized = re.sub(r'\s+', ' ', name_lower.strip())
        
        # Экранируем специальные символы, но оставляем буквы, цифры и пробелы
        escaped_name = re.escape(normalized)
        
        # Заменяем конкретные числа на общие паттерны
        # Тип радиатора
        pattern = escaped_name.replace(re.escape(str(rad_type)), r'(?P<type>\d+)')
        # Высота
        pattern = pattern.replace(re.escape(str(height)), r'(?P<height>\d{2,4})')
        # Длина
        pattern = pattern.replace(re.escape(str(length)), r'(?P<length>\d{3,4})')
        
        # Делаем пробелы более гибкими
        pattern = pattern.replace(r'\ ', r'\s+')
        # Делаем дефисы и тире более гибкими
        pattern = pattern.replace(r'\-', r'[\-\s]*')
        
        return pattern

    def test_pattern_on_series(self, pattern_index: int, test_names: List[str]):
        """
        Тестирует паттерн на серии названий для отладки.
        """
        if 0 <= pattern_index < len(self.patterns):
            pattern_data = self.patterns[pattern_index]
            regex_pattern = pattern_data["pattern"]
            
            print(f"\n=== ТЕСТ ПАТТЕРНА #{pattern_index} ===")
            print(f"Паттерн: {regex_pattern}")
            print(f"Пример: {pattern_data.get('original_example', 'N/A')}")
            
            compiled_regex = re.compile(regex_pattern, re.IGNORECASE)
            
            for test_name in test_names:
                match = compiled_regex.search(test_name.lower())
                if match:
                    print(f"✅ СОВПАДЕНИЕ: {test_name}")
                    if "type" in match.groupdict():
                        print(f"   Тип: {match.group('type')}")
                    if "height" in match.groupdict():
                        print(f"   Высота: {match.group('height')}")
                    if "length" in match.groupdict():
                        print(f"   Длина: {match.group('length')}")
                else:
                    print(f"❌ НЕТ СОВПАДЕНИЯ: {test_name}")        

    def find_match(self, name: str) -> Optional[Dict[str, Any]]:
        """
        Ищет подходящий шаблон для названия и извлекает параметры.
        ПРИОРИТЕТЫ (от высшего к низшему):
        1. Обученные правила (source == "learned") - ВЫСШИЙ ПРИОРИТЕТ
        2. Специфичные паттерны PRADO
        3. Специфичные паттерны OASIS  
        4. Специфичные паттерны форматов "Н" и "С"
        5. Универсальные паттерны
        6. Остальные правила
        """
        name_lower = name.lower().strip()
        
        print(f"[DEBUG] Поиск паттерна для: '{name}'")
        print(f"[DEBUG] Всего паттернов в базе: {len(self.patterns)}")
        
        # 1. ВЫСШИЙ ПРИОРИТЕТ: обученные правила (source == "learned")
        print(f"[DEBUG] Поиск в обученных правилах...")
        for i, pattern_data in enumerate(self.patterns):
            if pattern_data.get("source") != "learned":
                continue
                
            try:
                regex_pattern = pattern_data["pattern"]
                compiled_regex = re.compile(regex_pattern, re.IGNORECASE)
                match = compiled_regex.search(name_lower)
                
                if match:
                    print(f"[DEBUG] ✓ Найдено совпадение с ОБУЧЕННЫМ паттерном #{i}: {regex_pattern}")
                    
                    # БАЗОВЫЕ ПАРАМЕТРЫ ИЗ ПАТТЕРНА
                    base_connection = pattern_data.get("connection", "VK-правое")
                    base_type = pattern_data.get("rad_type", "10")
                    base_height = pattern_data.get("height", 500)
                    base_length = pattern_data.get("length", 1000)
                    
                    # ОПРЕДЕЛЯЕМ ФАКТИЧЕСКОЕ ПОДКЛЮЧЕНИЕ ИЗ НАЗВАНИЯ
                    actual_connection = self._determine_actual_connection(name_lower, base_connection)
                    
                    # ИЗВЛЕКАЕМ ПАРАМЕТРЫ ИЗ СОВПАДЕНИЯ
                    extracted_type = base_type
                    extracted_height = base_height
                    extracted_length = base_length
                    
                    # Проверяем, есть ли именованные группы
                    if match.groupdict():
                        # Паттерн с именованными группами
                        if "type" in match.groupdict():
                            extracted_type_raw = match.group("type")
                            if extracted_type_raw:
                                type_match = re.search(r'(\d+)', extracted_type_raw)
                                if type_match:
                                    extracted_type = type_match.group(1)
                                else:
                                    extracted_type = base_type
                            
                            # ПРЕОБРАЗОВАНИЕ ТИПА ДЛЯ KERMI
                            if 'kermi' in name_lower and extracted_type == '12':
                                extracted_type = '21'
                                print(f"[DEBUG] Преобразован тип Kermi 12 -> METEOR 21")
                            
                            print(f"[DEBUG] Извлечен тип: '{extracted_type_raw}' -> {extracted_type}")
                        
                        if "height" in match.groupdict():
                            try:
                                extracted_height = int(match.group("height"))
                                print(f"[DEBUG] Извлечена высота: {extracted_height}")
                            except (ValueError, TypeError):
                                extracted_height = base_height
                        
                        if "length" in match.groupdict():
                            try:
                                extracted_length = int(match.group("length"))
                                print(f"[DEBUG] Извлечена длина: {extracted_length}")
                            except (ValueError, TypeError):
                                extracted_length = base_length
                    else:
                        # Паттерн с обычными группами - анализируем количество групп
                        groups = match.groups()
                        print(f"[DEBUG] Найдены группы: {groups}")
                        
                        if len(groups) >= 3:
                            # Предполагаем формат: тип, высота, длина
                            try:
                                extracted_type = str(groups[0])
                                extracted_height = int(groups[1])
                                extracted_length = int(groups[2])
                                
                                # ПРЕОБРАЗОВАНИЕ ТИПА ДЛЯ KERMI
                                if 'kermi' in name_lower and extracted_type == '12':
                                    extracted_type = '21'
                                    print(f"[DEBUG] Преобразован тип Kermi 12 -> METEOR 21")
                                    
                                print(f"[DEBUG] Извлечено из групп: тип={extracted_type}, высота={extracted_height}, длина={extracted_length}")
                                
                            except (ValueError, TypeError) as e:
                                print(f"[DEBUG] Ошибка извлечения параметров из групп: {e}")
                                # Используем базовые значения при ошибке
                                extracted_type = base_type
                                extracted_height = base_height
                                extracted_length = base_length
                    
                    result = {
                        "connection": actual_connection,
                        "recognized": True,
                        "matched_by": f"learned_pattern #{i}",
                        "pattern_source": "learned",
                        "original_example": pattern_data.get("original_example", ""),
                        "brand_info": name_lower,
                        "rad_type": extracted_type,
                        "height": extracted_height,
                        "length": extracted_length
                    }
                        
                    print(f"[DEBUG] Финальное подключение: {actual_connection} (было: {base_connection})")
                        
                    return result
                        
            except re.error as e:
                print(f"[WARNING PatternManager] Пропущено обученное правило #{i} с некорректным регулярным выражением: {e}")
                continue
            except Exception as e:
                print(f"[ERROR PatternManager] Неожиданная ошибка в обученном правиле #{i}: {e}")
                continue
        
        # 2. ВЫСОКИЙ ПРИОРИТЕТ: Специфичные паттерны для PRADO
        print(f"[DEBUG] Поиск специфичных паттернов PRADO...")
        prado_patterns = [
            # PRADO Classic: "тип 22, высота H = 300 мм, Длина L=400 мм"
            r'prado classic.*?тип\s*(\d+).*?высота\s*h\s*=\s*(\d+)\s*мм.*?длина\s*l\s*=\s*(\d+)\s*мм',
            # PRADO Universal: "тип 11, высота H = 500 мм, Длина L=600 мм"  
            r'prado universal.*?тип\s*(\d+).*?высота\s*h\s*=\s*(\d+)\s*мм.*?длина\s*l\s*=\s*(\d+)\s*мм',
            # PRADO общий формат с мм
            r'prado.*?тип\s*(\d+).*?высота\s*(\d+)\s*мм.*?длина\s*(\d+)\s*мм',
            # PRADO общий формат с H/L
            r'prado.*?тип\s*(\d+).*?h\s*[=:\s]*(\d+).*?l\s*[=:\s]*(\d+)',
            # PRADO с тремя числами
            r'prado.*?(\d+).*?(\d{3}).*?(\d{3,4})'
        ]
        
        for i, pattern in enumerate(prado_patterns):
            match = re.search(pattern, name_lower)
            if match:
                print(f"[DEBUG] ✓ Найден специфичный паттерн PRADO #{i}: {pattern}")
                groups = match.groups()
                if len(groups) >= 3:
                    try:
                        rad_type = groups[0]
                        height = int(groups[1])
                        length = int(groups[2])
                        
                        # ПРЕОБРАЗОВАНИЕ ТИПА ДЛЯ KERMI
                        if 'kermi' in name_lower and rad_type == '12':
                            rad_type = '21'
                            print(f"[DEBUG] Преобразован тип Kermi 12 -> METEOR 21")
                        
                        # Определяем подключение по контексту
                        if 'боковым подкл' in name_lower or 'classic' in name_lower:
                            connection = "K-боковое"
                            print(f"[DEBUG] Определено подключение: K-боковое (PRADO Classic)")
                        elif 'нижним подкл' in name_lower or 'universal' in name_lower:
                            connection = "VK-правое" 
                            print(f"[DEBUG] Определено подключение: VK-правое (PRADO Universal)")
                        else:
                            connection = "K-боковое"  # по умолчанию для PRADO
                            print(f"[DEBUG] Определено подключение по умолчанию: K-боковое")
                        
                        print(f"[DEBUG] Извлечено из PRADO: тип={rad_type}, высота={height}, длина={length}")
                        
                        return {
                            "connection": connection,
                            "recognized": True,
                            "matched_by": f"prado_specific_pattern_{i}",
                            "pattern_source": "prado_specific",
                            "rad_type": rad_type,
                            "height": height,
                            "length": length
                        }
                        
                    except (ValueError, TypeError) as e:
                        print(f"[DEBUG] Ошибка извлечения параметров PRADO: {e}")
                        continue
        
        # 3. ВЫСОКИЙ ПРИОРИТЕТ: Специфичные паттерны для OASIS
        print(f"[DEBUG] Поиск специфичных паттернов OASIS...")
        oasis_patterns = [
            # OASIS формат: "Oasis Pro PB 22-4-04"
            r'oasis.*?(?:pb|pn|oc|ov)[\s\-]*(\d+)[\s\-]*(\d+)[\s\-]*(\d+)',
            # OASIS с тире
            r'oasis.*?(\d+)[\-\s](\d+)[\-\s](\d+)',
            # OASIS Pro формат
            r'oasis pro.*?(\d+)[\-\s](\d+)[\-\s](\d+)'
        ]
        
        for i, pattern in enumerate(oasis_patterns):
            match = re.search(pattern, name_lower)
            if match:
                print(f"[DEBUG] ✓ Найден специфичный паттерн OASIS #{i}: {pattern}")
                groups = match.groups()
                if len(groups) >= 3:
                    try:
                        rad_type = groups[0]
                        height_raw = int(groups[1])
                        length_raw = int(groups[2])
                        
                        # Преобразование высоты и длины для OASIS формата
                        height = height_raw * 100 if height_raw < 10 else height_raw
                        length = length_raw * 100 if length_raw < 100 else length_raw
                        
                        # ПРЕОБРАЗОВАНИЕ ТИПА ДЛЯ KERMI
                        if 'kermi' in name_lower and rad_type == '12':
                            rad_type = '21'
                            print(f"[DEBUG] Преобразован тип Kermi 12 -> METEOR 21")
                        
                        # Определяем подключение по префиксу
                        if 'pb' in name_lower or 'oc' in name_lower:
                            connection = "K-боковое"
                            print(f"[DEBUG] Определено подключение: K-боковое (OASIS PB/OC)")
                        elif 'pn' in name_lower or 'ov' in name_lower:
                            connection = "VK-правое"
                            print(f"[DEBUG] Определено подключение: VK-правое (OASIS PN/OV)")
                        else:
                            connection = "VK-правое"  # по умолчанию для OASIS
                            print(f"[DEBUG] Определено подключение по умолчанию: VK-правое")
                        
                        print(f"[DEBUG] Извлечено из OASIS: тип={rad_type}, высота={height_raw}->{height}, длина={length_raw}->{length}")
                        
                        return {
                            "connection": connection,
                            "recognized": True,
                            "matched_by": f"oasis_specific_pattern_{i}",
                            "pattern_source": "oasis_specific",
                            "rad_type": rad_type,
                            "height": height,
                            "length": length
                        }
                        
                    except (ValueError, TypeError) as e:
                        print(f"[DEBUG] Ошибка извлечения параметров OASIS: {e}")
                        continue
        
        # 4. ВЫСОКИЙ ПРИОРИТЕТ: Специфичные паттерны для форматов "Н20-300-700" и "С33-500-1100"
        print(f"[DEBUG] Поиск специфичных паттернов формата Н/С...")
        
        # Паттерн для формата "Н20-300-700" (нижнее подключение)
        h_pattern = r'н\s*(\d+)[\-\s]*(\d+)[\-\s]*(\d+)'
        h_match = re.search(h_pattern, name_lower)
        if h_match:
            print(f"[DEBUG] ✓ Найден специфичный паттерн Н-формата: {h_match.groups()}")
            try:
                rad_type = h_match.group(1)
                height = int(h_match.group(2))
                length = int(h_match.group(3))
                
                # Для формата "Н" - всегда нижнее подключение
                connection = "VK-правое"
                print(f"[DEBUG] Извлечено из Н-формата: тип={rad_type}, высота={height}, длина={length}, подключение={connection}")
                
                return {
                    "connection": connection,
                    "recognized": True,
                    "matched_by": "h_specific_pattern",
                    "pattern_source": "h_specific",
                    "rad_type": rad_type,
                    "height": height,
                    "length": length
                }
                
            except (ValueError, TypeError) as e:
                print(f"[DEBUG] Ошибка извлечения параметров Н-формата: {e}")
        
        # Паттерн для формата "С33-500-1100" (боковое подключение)  
        c_pattern = r'с\s*(\d+)[\-\s]*(\d+)[\-\s]*(\d+)'
        c_match = re.search(c_pattern, name_lower)
        if c_match:
            print(f"[DEBUG] ✓ Найден специфичный паттерн С-формата: {c_match.groups()}")
            try:
                rad_type = c_match.group(1)
                height = int(c_match.group(2))
                length = int(c_match.group(3))
                
                # Для формата "С" - всегда боковое подключение
                connection = "K-боковое"
                print(f"[DEBUG] Извлечено из С-формата: тип={rad_type}, высота={height}, длина={length}, подключение={connection}")
                
                return {
                    "connection": connection,
                    "recognized": True,
                    "matched_by": "c_specific_pattern",
                    "pattern_source": "c_specific",
                    "rad_type": rad_type,
                    "height": height,
                    "length": length
                }
                
            except (ValueError, TypeError) as e:
                print(f"[DEBUG] Ошибка извлечения параметров С-формата: {e}")
        
        # 5. НИЗКИЙ ПРИОРИТЕТ: универсальные паттерны
        print(f"[DEBUG] Поиск в универсальных паттернах...")
        
        # --- УНИВЕРСАЛЬНЫЕ ПАТТЕРНЫ ДЛЯ РАЗНЫХ ФОРМАТОВ ---
        universal_patterns = [
            # Формат с явным указанием параметров
            r'(?:тип|type)\s*[=:\-\s]*(\d+).*?(?:высота|h|height)[\s=:\-]*(\d+)\s*мм.*?(?:длина|l|length)[\s=:\-]*(\d+)\s*мм',
            # OASIS формат: "PN 22-2-04" (тип-высота-длина)
            r'(?:pn|oc|ov)[\s\-]*(\d+)[\s\-]*(\d+)[\s\-]*(\d+)',
            # KERMI формат: "FTV 12 300/1600" 
            r'(?:ftv|тип|type)[\s\-]*(\d+)[\s\-]*(\d+)[/x](\d+)',
            # Общий формат: "22-300-400" или "22 300 400"
            r'\b(\d+)[\s\-]*(\d{2,3})[\s\-]*(\d{2,4})\b',
            # Формат с высотой и длиной: "300/400" или "300x400"
            r'\b(\d{2,3})[/x](\d{2,4})\b',
            # Формат с типом и размерами: "тип 22 300 400"
            r'(?:тип|type)\s*(\d+)\s+(\d+)\s+(\d+)',
            # Самый общий формат - ищем три числа подряд
            r'(\d+)[\s\-/x]+(\d+)[\s\-/x]+(\d+)'
        ]
        
        for pattern in universal_patterns:
            match = re.search(pattern, name_lower)
            if match:
                print(f"[DEBUG] Найден универсальный паттерн: {pattern} - {match.groups()}")
                
                groups = match.groups()
                if len(groups) == 3:
                    # Формат "тип-высота-длина" или "тип высота длина"
                    rad_type = groups[0]
                    height_str = groups[1]
                    length_str = groups[2]
                elif len(groups) == 2:
                    # Формат "высота/длина" - определяем тип по умолчанию
                    rad_type = "10"  # тип по умолчанию
                    height_str = groups[0]
                    length_str = groups[1]
                else:
                    continue
                    
                try:
                    # Преобразуем высоту и длину в числа
                    height = int(height_str)
                    length = int(length_str)
                    
                    # ПРЕОБРАЗОВАНИЕ ТИПА ДЛЯ KERMI
                    if 'kermi' in name_lower and rad_type == '12':
                        rad_type = '21'
                        print(f"[DEBUG] Преобразован тип Kermi 12 -> METEOR 21")
                    
                    # Преобразование для OASIS формата (2 -> 200, 3 -> 300, 5 -> 500)
                    if height in [2, 3, 4, 5, 6, 7, 8, 9]:
                        height = height * 100
                        print(f"[DEBUG] Преобразована высота OASIS: {height_str} -> {height}")
                    
                    # Преобразование длины для OASIS формата (04 -> 400, 08 -> 800, 09 -> 900)
                    if length < 100:
                        length = length * 100
                        print(f"[DEBUG] Преобразована длина OASIS: {length_str} -> {length}")
                        
                except (ValueError, TypeError) as e:
                    print(f"[DEBUG] Ошибка преобразования чисел: {e}")
                    continue
                
                # Определяем подключение по ключевым словам
                if 'нижнее' in name_lower and 'правое' in name_lower:
                    connection = "VK-правое"
                elif 'нижнее' in name_lower and 'левое' in name_lower:
                    connection = "VK-левое" 
                elif 'нижнее' in name_lower or 'pn' in name_lower or 'ov' in name_lower or 'universal' in name_lower:
                    connection = "VK-правое"  # по умолчанию для нижнего
                elif 'боковое' in name_lower or 'oc' in name_lower or 'classic' in name_lower:
                    connection = "K-боковое"  # по умолчанию для бокового
                else:
                    # Автоопределение по L/R
                    if ' l' in name_lower or '-l' in name_lower or 'левое' in name_lower:
                        connection = "VK-левое"
                    elif ' r' in name_lower or '-r' in name_lower or 'правое' in name_lower:
                        connection = "VK-правое"
                    else:
                        connection = "VK-правое"  # по умолчанию
                
                result = {
                    "connection": connection,
                    "recognized": True,
                    "matched_by": "universal_pattern",
                    "pattern_source": "auto_detected",
                    "rad_type": rad_type,
                    "height": height,
                    "length": length
                }
                    
                return result
        
        # 6. САМЫЙ НИЗКИЙ ПРИОРИТЕТ: остальные правила (не обученные)
        print(f"[DEBUG] Поиск в остальных правилах...")
        for i, pattern_data in enumerate(self.patterns):
            if pattern_data.get("source") == "learned":
                continue  # Уже проверили обученные правила
                
            try:
                regex_pattern = pattern_data["pattern"]
                compiled_regex = re.compile(regex_pattern, re.IGNORECASE)
                match = compiled_regex.search(name_lower)
                
                if match:
                    print(f"[DEBUG] ✓ Найдено совпадение с паттерном #{i}: {regex_pattern}")
                    
                    base_connection = pattern_data.get("connection", "VK-правое")
                    base_type = pattern_data.get("rad_type", "10")
                    base_height = pattern_data.get("height", 500)
                    base_length = pattern_data.get("length", 1000)
                    
                    actual_connection = self._determine_actual_connection(name_lower, base_connection)
                    
                    result = {
                        "connection": actual_connection,
                        "recognized": True,
                        "matched_by": f"pattern #{i}",
                        "pattern_source": pattern_data.get("source", "unknown"),
                        "original_example": pattern_data.get("original_example", ""),
                        "brand_info": name_lower
                    }
                    
                    extracted_type = base_type
                    extracted_height = base_height
                    extracted_length = base_length
                    
                    if "type" in match.groupdict():
                        extracted_type_raw = match.group("type")
                        if extracted_type_raw:
                            type_match = re.search(r'(\d+)', extracted_type_raw)
                            if type_match:
                                extracted_type = type_match.group(1)
                            else:
                                extracted_type = base_type
                        
                        # ПРЕОБРАЗОВАНИЕ ТИПА ДЛЯ KERMI
                        if 'kermi' in name_lower and extracted_type == '12':
                            extracted_type = '21'
                            print(f"[DEBUG] Преобразован тип Kermi 12 -> METEOR 21")
                        
                        print(f"[DEBUG] Извлечен тип: '{extracted_type_raw}' -> {extracted_type}")
                    
                    if "height" in match.groupdict():
                        try:
                            extracted_height = int(match.group("height"))
                            if extracted_height < 100:
                                extracted_height = extracted_height * 100
                                print(f"[DEBUG] Преобразована высота: {match.group('height')} -> {extracted_height}")
                            else:
                                print(f"[DEBUG] Извлечена высота: {extracted_height}")
                        except (ValueError, TypeError):
                            extracted_height = base_height
                    
                    if "length" in match.groupdict():
                        try:
                            extracted_length = int(match.group("length"))
                            if extracted_length < 100:
                                extracted_length = extracted_length * 100
                                print(f"[DEBUG] Преобразована длина: {match.group('length')} -> {extracted_length}")
                            else:
                                print(f"[DEBUG] Извлечена длина: {extracted_length}")
                        except (ValueError, TypeError):
                            extracted_length = base_length
                    
                    result["rad_type"] = extracted_type
                    result["height"] = extracted_height
                    result["length"] = extracted_length
                        
                    print(f"[DEBUG] Финальное подключение: {actual_connection} (было: {base_connection})")
                        
                    return result
                    
            except re.error as e:
                print(f"[WARNING PatternManager] Пропущено правило #{i} с некорректным регулярным выражением: {e}")
                continue
            except Exception as e:
                print(f"[ERROR PatternManager] Неожиданная ошибка в правиле #{i}: {e}")
                continue
        
        print(f"[DEBUG] Не найдено подходящих паттернов")
        return None
    
    def _determine_actual_connection(self, name_lower: str, base_connection: str) -> str:
        """
        Определяет фактическое подключение на основе содержания названия.
        Учитывает симметричные типы радиаторов METEOR (20,21,22).
        """
        # ЯВНЫЕ УКАЗАНИЯ В НАЗВАНИИ - ВЫСШИЙ ПРИОРИТЕТ
        # Ключевые слова для НИЖНЕГО подключения (VENTIL COMPACT)
        lower_keywords = [
            'ventil', 'vc', 'vk', 'valve', 'нижн', 'нижним', 'нижнее', 'нижнее подключение',
            'universal', 'vh', 'VH', 'pn', 'ov', 'vk-profil'
        ]
        
        # Ключевые слова для БОКОВОГО подключения  
        side_keywords = [
            'боков', 'боковое', 'боковой', 'боковое подключение', 'k-profil', 'k-боковое', 
            'k-профиль', 'oc', 'pb'
        ]
        
        # Одиночные буквы - только в контексте
        single_letters = [' c', '-c', '(c', ' c ', '-c ', '(c ']  # Боковое
        single_letters_lower = [' v', '-v', '(v', ' v ', '-v ', '(v ']  # Нижнее
        
        left_keywords = ['левое', 'левой', 'левом', ' l', '-l', '(l']
        right_keywords = ['правое', 'правой', 'правом', ' r', '-r', '(r']
        
        # Проверяем ключевые слова в названии
        has_lower = any(keyword in name_lower for keyword in lower_keywords)
        has_side = any(keyword in name_lower for keyword in side_keywords)
        has_single_c = any(letter in name_lower for letter in single_letters)
        has_single_v = any(letter in name_lower for letter in single_letters_lower)
        has_left = any(keyword in name_lower for keyword in left_keywords)
        has_right = any(keyword in name_lower for keyword in right_keywords)
        
        print(f"[CONNECTION DEBUG] '{name_lower}' -> lower: {has_lower}, side: {has_side}, single_c: {has_single_c}, single_v: {has_single_v}")
        
        # ОПРЕДЕЛЯЕМ ПОДКЛЮЧЕНИЕ НА ОСНОВЕ НАЗВАНИЯ (ВЫСШИЙ ПРИОРИТЕТ)
        
        # СЛУЧАЙ 1: Явно указано нижнее подключение (ВЫСШИЙ ПРИОРИТЕТ)
        if 'нижнее подключение' in name_lower or 'vk-profil' in name_lower:
            print(f"[CONNECTION DEBUG] Обнаружено явное указание нижнего подключения")
            # ДЛЯ СИММЕТРИЧНЫХ ТИПОВ (20,21,22) - ВСЕГДА VK-правое
            # Извлекаем тип из названия для проверки
            type_match = re.search(r'(?:тип|type)?\s*(\d+)[\-\s]', name_lower)
            if type_match:
                rad_type = type_match.group(1)
                if rad_type in ['20', '21', '22']:
                    print(f"[CONNECTION DEBUG] Симметричный тип {rad_type} - принудительно VK-правое")
                    return "VK-правое"
            
            if has_left:
                return "VK-левое"
            elif has_right:
                return "VK-правое"
            else:
                return "VK-правое"
        
        # СЛУЧАЙ 2: Явно указано боковое подключение (ВЫСШИЙ ПРИОРИТЕТ)
        elif 'боковое подключение' in name_lower or 'k-profil' in name_lower:
            print(f"[CONNECTION DEBUG] Обнаружено явное указание бокового подключения")
            return "K-боковое"
        
        # СЛУЧАЙ 3: COMPACT C22 - анализируем контекст
        elif 'compact c' in name_lower:
            # COMPACT C22 - это БОКОВОЕ подключение (Royal Thermo Compact)
            print(f"[CONNECTION DEBUG] Обнаружен COMPACT C - боковое подключение")
            return "K-боковое"
        
        # СЛУЧАЙ 4: VENTIL COMPACT VC22 - это нижнее подключение
        elif 'ventil compact' in name_lower or 'vc' in name_lower:
            print(f"[CONNECTION DEBUG] Обнаружен VENTIL COMPACT/VC - нижнее подключение")
            # ДЛЯ СИММЕТРИЧНЫХ ТИПОВ (20,21,22) - ВСЕГДА VK-правое
            type_match = re.search(r'(?:тип|type)?\s*(\d+)[\-\s]', name_lower)
            if type_match:
                rad_type = type_match.group(1)
                if rad_type in ['20', '21', '22']:
                    print(f"[CONNECTION DEBUG] Симметричный тип {rad_type} - принудительно VK-правое")
                    return "VK-правое"
            
            if has_left:
                return "VK-левое"
            else:
                return "VK-правое"
        
        # СЛУЧАЙ 5: OASIS PN - нижнее подключение
        elif ('oasis' in name_lower and 'pn' in name_lower) or ('ov' in name_lower):
            print(f"[CONNECTION DEBUG] Обнаружен OASIS PN/OV - нижнее подключение")
            # ДЛЯ СИММЕТРИЧНЫХ ТИПОВ (20,21,22) - ВСЕГДА VK-правое
            type_match = re.search(r'(?:тип|type)?\s*(\d+)[\-\s]', name_lower)
            if type_match:
                rad_type = type_match.group(1)
                if rad_type in ['20', '21', '22']:
                    print(f"[CONNECTION DEBUG] Симметричный тип {rad_type} - принудительно VK-правое")
                    return "VK-правое"
            
            if has_left:
                return "VK-левое"
            else:
                return "VK-правое"
        
        # СЛУЧАЙ 6: Общие ключевые слова (нижний приоритет)
        elif has_lower and not has_side:
            # ДЛЯ СИММЕТРИЧНЫХ ТИПОВ (20,21,22) - ВСЕГДА VK-правое
            type_match = re.search(r'(?:тип|type)?\s*(\d+)[\-\s]', name_lower)
            if type_match:
                rad_type = type_match.group(1)
                if rad_type in ['20', '21', '22']:
                    print(f"[CONNECTION DEBUG] Симметричный тип {rad_type} - принудительно VK-правое")
                    return "VK-правое"
            
            if has_left:
                return "VK-левое"
            elif has_right:
                return "VK-правое"
            else:
                return "VK-правое"
        
        elif has_side and not has_lower:
            return "K-боковое"
        
        # СЛУЧАЙ 7: Конфликт - анализируем что доминирует
        elif has_lower and has_side:
            vc_count = name_lower.count('vc') + name_lower.count('ventil') + name_lower.count('valve') + name_lower.count('vk-profil')
            oc_count = name_lower.count('oc') + name_lower.count('боков') + name_lower.count('k-profil')
            if vc_count > oc_count:
                # ДЛЯ СИММЕТРИЧНЫХ ТИПОВ (20,21,22) - ВСЕГДА VK-правое
                type_match = re.search(r'(?:тип|type)?\s*(\d+)[\-\s]', name_lower)
                if type_match:
                    rad_type = type_match.group(1)
                    if rad_type in ['20', '21', '22']:
                        print(f"[CONNECTION DEBUG] Симметричный тип {rad_type} - принудительно VK-правое")
                        return "VK-правое"
                return "VK-правое"
            else:
                return "K-боковое"
        
        # СЛУЧАЙ 8: Только одиночные буквы
        elif has_single_c and not has_single_v:
            return "K-боковое"
        elif has_single_v and not has_single_c:
            # ДЛЯ СИММЕТРИЧНЫХ ТИПОВ (20,21,22) - ВСЕГДА VK-правое
            type_match = re.search(r'(?:тип|type)?\s*(\d+)[\-\s]', name_lower)
            if type_match:
                rad_type = type_match.group(1)
                if rad_type in ['20', '21', '22']:
                    print(f"[CONNECTION DEBUG] Симметричный тип {rad_type} - принудительно VK-правое")
                    return "VK-правое"
            return "VK-правое"
        
        else:
            # Нет явных указаний - используем из паттерна
            print(f"[CONNECTION DEBUG] Используем подключение из паттерна: {base_connection}")
            return base_connection

    def add_pattern(self, name, connection, rad_type, height, length, art, met_name):
        # Создаём новое правило
        new_pattern = {
            "pattern": self.create_pattern_from_name(name),
            "connection": connection,
            "rad_type": rad_type,
            "height": height,
            "length": length,
            "source": "learned"
        }
        self.patterns.append(new_pattern)
        # Вызов save_pattern с 7 аргументами
        # ПЕРЕДАЁМ ВСЕ 7 АРГУМЕНТОВ
        self.save_pattern(name, connection, rad_type, height, length, art, met_name)

    def learn_pattern(self, original_name, meteor_params, meteor_art, meteor_name):
        """
        Обучает новому правилу на основе названия конкурента и выбранного аналога METEOR.
        """
        try:
            print(f"[LEARNING] Обучение: Оригинальное имя '{original_name}'")
            print(f"[LEARNING] Обучение: Параметры METEOR {meteor_params}")
            print(f"[LEARNING] Обучение: Выбранный METEOR арт: {meteor_art}, имя: {meteor_name}")

            # Извлекаем параметры METEOR из словаря
            connection = meteor_params['connection']
            rad_type = meteor_params['rad_type']
            height = meteor_params['height']
            length = meteor_params['length']

            # Вызов save_pattern с правильными аргументами
            self.save_pattern(original_name, connection, rad_type, height, length, meteor_art, meteor_name)
            
            print(f"[LEARNING] Обучение завершено для: {original_name}")
            
        except Exception as e:
            print(f"[ERROR PatternManager] Ошибка при обучении правила для '{original_name}': {e}")
            import traceback
            traceback.print_exc()
            
    def get_all_patterns(self) -> List[Dict[str, Any]]:
        """Возвращает список всех текущих правил."""
        return self.patterns

    def remove_pattern(self, index: int):
        """Удаляет правило по индексу."""
        if 0 <= index < len(self.patterns):
            removed_rule = self.patterns.pop(index)
            self.save_patterns()
            print(f"[INFO] Удалено правило: {removed_rule}")
        else:
            print(f"[WARNING] Попытка удаления несуществующего правила с индексом {index}")
            
    def update_pattern(self, index: int, new_rule: Dict[str, Any]):
        """Обновляет правило по индексу."""
        if 0 <= index < len(self.patterns):
            self.patterns[index] = new_rule
            self.save_patterns()
            print(f"[INFO] Обновлено правило с индексом {index}")
        else:
            print(f"[WARNING] Попытка обновления несуществующего правила с индексом {index}")