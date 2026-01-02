import re
from typing import Dict, Any

class RadiatorNameParser:
    """
    Универсальный парсер названий радиаторов конкурентов.
    Извлекает параметры: тип, высоту, длину и вид подключения.
    """
    
    @staticmethod
    def advanced_parse_radiator_name(name: str) -> Dict[str, Any]:
        """
        УНИВЕРСАЛЬНЫЙ ПАРСЕР - улучшенная версия с приоритетом для формата "тип\высота\длина"
        """
        result = {
            'recognized': False,
            'connection': 'VK-правое',  # По умолчанию
            'type': '10',               # Тип по умолчанию
            'height': None,
            'length': None
        }
        
        try:
            if not isinstance(name, str):
                return result

            name_lower = name.lower().strip()
            print(f"[DEBUG] Парсинг названия: '{name}'")

            # --- НОВЫЙ КОД: ВЫСОКИЙ ПРИОРИТЕТ для формата ЛИДЕЯ "ЛК XX-YYY" или "ЛУ XX-YYY" ---
            lidea_pattern = r'(л[кkуy])\s*(\d{2})[\-\s]*(\d{3,4})'
            lidea_match = re.search(lidea_pattern, name_lower)
            if lidea_match:
                try:
                    connection_type = lidea_match.group(1)  # лк или лу
                    rad_type = lidea_match.group(2)  # 11, 22, 33
                    lidea_code = int(lidea_match.group(3))  # 504, 304 и т.д.
                    
                    # Определяем высоту
                    height = RadiatorNameParser._detect_lidea_height(lidea_code)
                    
                    # Определяем длину
                    length = RadiatorNameParser._convert_lidea_length(lidea_code)
                    
                    # Определяем подключение
                    if 'лу' in connection_type:
                        connection = 'VK-правое'
                    else:
                        connection = 'K-боковое'
                    
                    result.update({
                        'recognized': True,
                        'connection': connection,
                        'type': rad_type,
                        'height': height,
                        'length': length
                    })
                    
                    print(f"[DEBUG] Распознан ЛИДЕЯ: {connection_type}{rad_type}-{lidea_code} -> тип={rad_type}, высота={height}, длина={length}, подключение={connection}")
                    return result
                    
                except (ValueError, IndexError) as e:
                    print(f"[DEBUG] Ошибка парсинга Лидея: {e}")
                    # Продолжаем парсить другими методами

            # --- НОВЫЙ КОД: ВЫСОКИЙ ПРИОРИТЕТ для формата KERMI ---
            # Паттерны для Kermi: "KERMI Profil-V FTV 11 500/1200", "KERMI Profil-K FK0 22 500/1000"
            kermi_patterns = [
                r'kermi.*?profil[-\s]?[vk]\s+ft[vk]\s*(\d{2})\s+(\d{3})/(\d{3,4})',
                r'kermi.*?ft[vk]\s*(\d{2})\s+(\d{3})/(\d{3,4})',
                r'kermi.*?profil[-\s]?[vk].*?(\d{2})\s+(\d{3})/(\d{3,4})',
                r'kermi.*?(\d{2})\s+(\d{3})/(\d{3,4})'
            ]
            
            for pattern in kermi_patterns:
                match = re.search(pattern, name_lower)
                if match and len(match.groups()) == 3:
                    try:
                        rad_type = match.group(1)  # 11, 12, 22 и т.д.
                        height = int(match.group(2))  # 500
                        length = int(match.group(3))  # 1200 и т.д.
                        
                        # Определяем подключение для Kermi
                        connection = RadiatorNameParser._determine_kermi_connection(name_lower)
                        
                        # Специальное преобразование типа для Kermi: 12 -> 21
                        if rad_type == '12':
                            rad_type = '21'
                        
                        result.update({
                            'recognized': True,
                            'connection': connection,
                            'type': rad_type,
                            'height': height,
                            'length': length
                        })
                        
                        print(f"[DEBUG] Распознан KERMI: тип={rad_type}, высота={height}, длина={length}, подключение={connection}")
                        return result
                        
                    except (ValueError, IndexError) as e:
                        print(f"[DEBUG] Ошибка парсинга Kermi: {e}")
                        continue

            # --- ВЫСОКИЙ ПРИОРИТЕТ: Формат с обратными слешами "тип\высота\длина" ---
            backslash_pattern = r'(\d{2})[\\\/](\d{2,4})[\\\/](\d{3,4})'
            backslash_match = re.search(backslash_pattern, name)
            if backslash_match:
                try:
                    rad_type = backslash_match.group(1)
                    height = int(backslash_match.group(2))
                    length = int(backslash_match.group(3))
                    
                    connection = RadiatorNameParser._determine_connection(name_lower)
                    
                    result.update({
                        'recognized': True,
                        'connection': connection,
                        'type': rad_type,
                        'height': height,
                        'length': length
                    })
                    
                    print(f"[DEBUG] Распознан формат тип\\высота\\длина: тип={rad_type}, высота={height}, длина={length}, подключение={connection}")
                    return result
                    
                except ValueError as e:
                    print(f"[DEBUG] Ошибка преобразования чисел в формате тип\\высота\\длина: {e}")

            # --- ВЫСОКИЙ ПРИОРИТЕТ: Формат с дефисами "тип-высота-длина" ---
            hyphen_pattern = r'(\d{2})[\-\s](\d{2,4})[\-\s](\d{3,4})'
            hyphen_match = re.search(hyphen_pattern, name)
            if hyphen_match:
                try:
                    rad_type = hyphen_match.group(1)
                    height = int(hyphen_match.group(2))
                    length = int(hyphen_match.group(3))
                    
                    connection = RadiatorNameParser._determine_connection(name_lower)
                    
                    result.update({
                        'recognized': True,
                        'connection': connection,
                        'type': rad_type,
                        'height': height,
                        'length': length
                    })
                    
                    print(f"[DEBUG] Распознан формат тип-высота-длина: тип={rad_type}, высота={height}, длина={length}, подключение={connection}")
                    return result
                    
                except ValueError as e:
                    print(f"[DEBUG] Ошибка преобразования чисел в формате тип-высота-длина: {e}")

            # --- ПАТТЕРНЫ ДЛЯ METEOR ---
            meteor_patterns = [
                r'радиатор\s+meteor\s+(?:classic|universal)\s+(?:vk|k)[\-\s]*(\d+)[/\s]*(\d+)[/\s]*(\d+)',
                r'meteor.*?(?:vk|k)[\-\s]*(\d+)[/\s]*(\d+)[/\s]*(\d+)',
                r'meteor.*?(\d+)[/\s]*(\d+)[/\s]*(\d+)',
                r'\b(\d+)[/\s]*(\d+)[/\s]*(\d+)\s*(?:ra|ls|re)\b'
            ]
            
            for pattern in meteor_patterns:
                match = re.search(pattern, name_lower)
                if match and len(match.groups()) == 3:
                    try:
                        rad_type = match.group(1)
                        height = int(match.group(2))
                        length = int(match.group(3))
                        
                        connection = RadiatorNameParser._determine_meteor_connection(name_lower)
                        
                        result.update({
                            'recognized': True,
                            'connection': connection,
                            'type': rad_type,
                            'height': height,
                            'length': length
                        })
                        
                        print(f"[DEBUG] Распознан METEOR: тип={rad_type}, высота={height}, длина={length}, подключение={connection}")
                        return result
                        
                    except ValueError as e:
                        print(f"[DEBUG] Ошибка преобразования чисел METEOR: {e}")
                        continue

            # --- ПАТТЕРНЫ ДЛЯ LIDEA (старая версия - оставлена для обратной совместимости) ---
            lidea_patterns = [
                r'радиаторы\s+стальные\s+штампованные\s+["]?лидея["]?\s+л[кkуy]\s*(\d+)[\-\s]*(\d+)',
                r'л[кkуy]\s*(\d+)[\-\s]*(\d+)',
                r'радиатор\s+лидея\s+л[кkуy]\s*(\d+)[\-\s]*(\d+)'
            ]
            
            for pattern in lidea_patterns:
                match = re.search(pattern, name_lower)
                if match:
                    try:
                        if len(match.groups()) == 2:
                            rad_type = match.group(1)
                            length_raw = int(match.group(2))
                            
                            height = RadiatorNameParser._detect_lidea_height(length_raw)
                            length = RadiatorNameParser._convert_lidea_length(length_raw)
                            connection = RadiatorNameParser._determine_lidea_connection(name_lower)
                            
                            result.update({
                                'recognized': True,
                                'connection': connection,
                                'type': rad_type,
                                'height': height,
                                'length': length
                            })
                            
                            print(f"[DEBUG] Распознан LIDEA (старый метод): тип={rad_type}, высота={height}, длина={length_raw}->{length}, подключение={connection}")
                            return result
                            
                    except ValueError as e:
                        print(f"[DEBUG] Ошибка преобразования чисел LIDEA: {e}")
                        continue

            # --- УЛУЧШЕННАЯ ЛОГИКА ДЛЯ ОБЩИХ СЛУЧАЕВ ---
            
            # Определяем подключение
            result['connection'] = RadiatorNameParser._determine_connection(name_lower)

            # Улучшенный поиск трех чисел подряд (тип, высота, длина)
            numbers_found = re.findall(r'\b\d{2,4}\b', name)
            print(f"[DEBUG] Найдены все числа: {numbers_found}")

            if len(numbers_found) >= 3:
                nums = []
                for num_str in numbers_found:
                    try:
                        num = int(num_str)
                        nums.append(num)
                    except ValueError:
                        continue

                print(f"[DEBUG] Числа после преобразования: {nums}")
                
                # Ищем три числа в правильном порядке
                for i in range(len(nums) - 2):
                    potential_type = nums[i]
                    potential_height = nums[i+1]
                    potential_length = nums[i+2]
                    
                    # Проверяем валидность комбинации
                    if (10 <= potential_type <= 33 and 
                        potential_height in [300, 400, 500, 600, 900] and
                        400 <= potential_length <= 2000):
                        
                        result.update({
                            'recognized': True,
                            'type': str(potential_type),
                            'height': potential_height,
                            'length': potential_length
                        })
                        print(f"[DEBUG] Найдена валидная комбинация: тип={potential_type}, высота={potential_height}, длина={potential_length}")
                        return result

            # Альтернативная логика для случаев с 2 числами
            if len(numbers_found) >= 2:
                nums = []
                for num_str in numbers_found:
                    try:
                        num = int(num_str)
                        nums.append(num)
                    except ValueError:
                        continue

                print(f"[DEBUG] Числа после преобразования: {nums}")
                
                # Сортируем числа по категориям
                potential_types = []
                potential_heights = []
                potential_lengths = []

                for num in nums:
                    if 10 <= num <= 33 and num in [10, 11, 20, 21, 22, 30, 33]:
                        potential_types.append(num)
                    elif num in [300, 400, 500, 600, 900]:
                        potential_heights.append(num)
                    elif 400 <= num <= 2000:
                        potential_lengths.append(num)

                print(f"[DEBUG] После сортировки - типы: {potential_types}, высоты: {potential_heights}, длины: {potential_lengths}")

                # Логика для заполнения недостающих параметров
                if potential_types and potential_heights and potential_lengths:
                    result.update({
                        'recognized': True,
                        'type': str(potential_types[0]),
                        'height': potential_heights[0],
                        'length': potential_lengths[0]
                    })
                elif potential_types and potential_heights and not potential_lengths:
                    result.update({
                        'recognized': True,
                        'type': str(potential_types[0]),
                        'height': potential_heights[0],
                        'length': 1000  # значение по умолчанию
                    })
                elif potential_types and not potential_heights and potential_lengths:
                    result.update({
                        'recognized': True,
                        'type': str(potential_types[0]),
                        'height': 500,  # по умолчанию
                        'length': potential_lengths[0]
                    })
                elif not potential_types and potential_heights and potential_lengths:
                    result.update({
                        'recognized': True,
                        'type': '10',  # по умолчанию
                        'height': potential_heights[0],
                        'length': potential_lengths[0]
                    })

            print(f"[DEBUG] Финальный результат для '{name}': recognized={result['recognized']}, type={result['type']}, height={result['height']}, length={result['length']}, connection={result['connection']}")
            return result

        except Exception as e:
            print(f"[ERROR] Ошибка в advanced_parse_radiator_name: {e}")
            return result

    @staticmethod
    def _determine_connection(name_lower: str) -> str:
        """Определяет тип подключения по ключевым словам"""
        if 'vk' in name_lower or 'vc' in name_lower or 'нижн' in name_lower or 'universal' in name_lower or 'ventil' in name_lower:
            return 'VK-правое'
        elif 'k' in name_lower or 'боков' in name_lower or 'classic' in name_lower:
            return 'K-боковое'
        else:
            return 'VK-правое'  # по умолчанию

    @staticmethod
    def _determine_meteor_connection(name_lower: str) -> str:
        """Определяет подключение для METEOR"""
        if 'vk' in name_lower or 'vl' in name_lower:
            return 'VK-правое'
        elif 'k' in name_lower:
            return 'K-боковое'
        else:
            if 'classic' in name_lower:
                return 'K-боковое'
            elif 'universal' in name_lower:
                return 'VK-правое'
            else:
                return 'VK-правое'

    @staticmethod
    def _determine_kermi_connection(name_lower: str) -> str:
        """Определяет подключение для KERMI - НОВЫЙ МЕТОД"""
        # По вашим данным: Profil-V FTV = VK (нижнее), Profil-K FK0 = K (боковое)
        if 'profil-v' in name_lower or 'ftv' in name_lower or 'нижнее' in name_lower:
            # Определяем сторону для VK
            if 'левое' in name_lower or 'la' in name_lower:
                return 'VK-левое'
            else:
                return 'VK-правое'  # по умолчанию для VK
        elif 'profil-k' in name_lower or 'fk0' in name_lower or 'боковое' in name_lower:
            return 'K-боковое'
        else:
            return 'VK-правое'  # по умолчанию

    @staticmethod
    def _determine_lidea_connection(name_lower: str) -> str:
        """Определяет подключение для LIDEA - ИСПРАВЛЕННЫЙ"""
        if 'лу' in name_lower or 'ly' in name_lower or 'lу' in name_lower:
            return 'VK-правое'
        else:
            return 'K-боковое'  # по умолчанию для ЛК

    @staticmethod
    def _detect_lidea_height(lidea_code: int) -> int:
        """Определяет высоту радиатора Лидея по коду модели - ИСПРАВЛЕННЫЙ"""
        if lidea_code < 100:
            return 500  # по умолчанию для коротких кодов
            
        try:
            first_digit = int(str(lidea_code)[0])
            if first_digit == 3:
                return 300
            elif first_digit == 4:
                return 400
            elif first_digit == 5:
                return 500
            elif first_digit == 6:
                return 600
            elif first_digit == 7:
                return 700
            elif first_digit == 9:
                return 900
            else:
                return 500  # по умолчанию
        except (ValueError, IndexError):
            return 500  # по умолчанию при ошибке

    @staticmethod
    def _convert_lidea_length(lidea_code: int) -> int:
        """Конвертирует длину Лидея в стандартную длину METEOR - ИСПРАВЛЕННЫЙ"""
        if lidea_code < 100:
            # Если код двухзначный (например 12), умножаем на 100
            return lidea_code * 100
            
        try:
            # Берем последние 2 цифры кода
            code_str = str(lidea_code)
            if len(code_str) >= 2:
                last_two_digits = int(code_str[-2:])
                return last_two_digits * 100  # 04 → 400, 12 → 1200, 16 → 1600
            else:
                return lidea_code * 100  # если код однозначный
        except (ValueError, IndexError):
            # Возвращаем исходный код, если не удалось преобразовать
            return lidea_code

    @staticmethod
    def parse_evra_name(name: str) -> Dict[str, Any]:
        """Парсит наименование радиатора EVRA"""
        result = {
            'recognized': False,
            'connection': None,
            'type': None,
            'height': None,
            'length': None
        }
        
        try:
            name_lower = ' '.join(name.split()).lower()
            
            patterns = [
                r'valve compact.*?v[сc]\s*(\d+)[-\s](\d+)[-\s](\d+)',
                r'v[сc]\s*(\d+)[-\s](\d+)[-\s](\d+)'
            ]
            
            for pattern in patterns:
                match = re.search(pattern, name_lower)
                if match:
                    rad_type = match.group(1)
                    height = int(match.group(2))
                    length = int(match.group(3))
                    
                    connection = 'VK-правое'  # Valve Compact всегда VK-правое
                    
                    result.update({
                        'recognized': True,
                        'connection': connection,
                        'type': rad_type,
                        'height': height,
                        'length': length
                    })
                    break
                    
        except Exception as e:
            print(f"[ERROR] Ошибка парсинга имени EVRA: {str(e)}")
            
        return result

    @staticmethod
    def parse_foreign_radiator_name_flexibly(self, name: str, designation: str = None) -> dict:
        """
        Улучшенный парсер для гибкого извлечения параметров из наименования и обозначения радиатора.
        Приоритет отдается информации из 'designation', если она содержит структурированные данные.
        """
        # Инициализация переменных
        brand = ""
        model = ""
        type_ = ""
        height = None
        length = None
        connection = "VK-правое" # Значение по умолчанию
        connection_side = ""

        # Словари для поиска
        brands = ['Prado', 'Kermi', 'Purmo', 'Royal Thermo', 'Buderus', 'Stout', 'Global', 'Oasis', 'Itap', 'Sira']
        models = ['Classic', 'FTV', 'K2', 'R3000', 'Eco', 'Vogue', 'Style', 'Compact', 'Universal', 'Oasis', 'Cento']
        types = ['10', '11', '20', '21', '22', '23', '30', '33']
        connections_vk = ['vk', 'нижнее']
        connections_k = ['k', 'боковое']

        # --- ПАРСИНГ ОБОЗНАЧЕНИЯ (артикула) ---
        if designation:
            designation_clean = designation.strip().lower()
            # Паттерн для артикулов вроде "RA 22 500 1000" или "Kermi FTV 22 500 1000"
            # Извлекаем brand, model, type, height, length
            # Примеры: RA 22 500 900, Kermi FTV 22 500 1000, Purmo K2 22 500 1000
            designation_pattern = r'(?:([A-Za-z]+)\s+)?([A-Za-z0-9]+)?\s*(\d{2})\s+(\d{3,4})\s+(\d{3,4})'
            match = re.search(designation_pattern, designation, re.IGNORECASE)
            if match:
                # Извлекаем компоненты
                designation_brand = match.group(1) # может быть None
                designation_model = match.group(2) # может быть None
                designation_type = match.group(3)
                designation_height = match.group(4)
                designation_length = match.group(5)

                # Обновляем переменные, если они соответствуют ожидаемым форматам
                if designation_brand and any(designation_brand.lower() == b.lower() for b in brands):
                    brand = designation_brand
                if designation_model and any(designation_model.lower() == m.lower() for m in models):
                    model = designation_model
                if designation_type in types:
                    type_ = designation_type
                if designation_height.isdigit():
                    height = int(designation_height)
                if designation_length.isdigit():
                    length = int(designation_length)

        # --- ПАРСИНГ НАИМЕНОВАНИЯ ---
        # Очищаем и приводим к нижнему регистру для поиска
        name_clean = name.strip().lower()
        original_name_lower = name.lower()

        # 1. Извлечение connection_side (ra, re, ls, rs) из конца строки наименования
        # Это может перезаписать connection_side, если он был извлечен из designation
        # Это может быть не всегда корректно, но часто суффикс в наименовании точнее
        if original_name_lower.endswith('ra'):
            connection_side = 'ra'
        elif original_name_lower.endswith('re'):
            connection_side = 're'
        elif original_name_lower.endswith('ls'):
            connection_side = 'ls'
        elif original_name_lower.endswith('rs'):
            connection_side = 'rs'

        # 2. Извлечение connection (VK, K) из наименования
        # Используем информацию из наименования для уточнения или если connection не был установлен через designation
        for conn in connections_vk:
            if conn in name_clean:
                connection = "VK-правое" # по умолчанию для VK
                if 'правое' in name_clean or connection_side == 'ra':
                    connection = "VK-правое"
                elif 'левое' in name_clean or connection_side == 're':
                    connection = "VK-левое"
                break # выходим, если нашли VK
        for conn in connections_k:
            if conn in name_clean:
                connection = "K-боковое" # по умолчанию для K
                if 'правое' in name_clean or connection_side == 'rs':
                    connection = "K-правое"
                elif 'левое' in name_clean or connection_side == 'ls':
                    connection = "K-левое"
                break # выходим, если нашли K

        # 3. Извлечение brand и model из наименования
        # Используем, если они не были извлечены из designation
        if not brand:
            for b in brands:
                if re.search(rf'\b{re.escape(b)}\b', name, re.IGNORECASE):
                    brand = b
                    break
        if not model:
            for m in models:
                if re.search(rf'\b{re.escape(m)}\b', name, re.IGNORECASE):
                    model = m
                    break

        # 4. Извлечение type, height, length из наименования
        # Паттерн для формата /высота/длина, например VK 22/500/900
        # Используем, если height/length не были извлечены из designation
        if not height or not length:
            height_length_pattern = r'/(\d{3,4})/(\d{3,4})'
            height_length_match = re.search(height_length_pattern, name)
            if height_length_match:
                found_height = height_length_match.group(1)
                found_length = height_length_match.group(2)
                if not height and found_height.isdigit():
                    height = int(found_height)
                if not length and found_length.isdigit():
                    length = int(found_length)
            # Также ищем тип рядом с формата /высота/длина
            if not type_:
                type_h_l_pattern_with_type = r'([VKvkKk])\s*(\d{2})/(\d{3,4})/(\d{3,4})'
                type_h_l_match = re.search(type_h_l_pattern_with_type, name)
                if type_h_l_match:
                    # connection_part = type_h_l_match.group(1) # уже обработано выше
                    found_type = type_h_l_match.group(2)
                    if found_type in types:
                        type_ = found_type

        # Если тип не найден в /высота/длина, ищем отдельно
        if not type_:
            for t in types:
                if re.search(rf'\b{t}\b', name):
                    type_ = t
                    break

        # --- ПРЕОБРАЗОВАНИЯ ---
        # Преобразование типа для Kermi 12 -> METEOR 21
        if brand.lower() == 'kermi' and type_ == '12':
            type_ = '21'

        # --- ВОЗВРАТ РЕЗУЛЬТАТА ---
        # Определяем финальное подключение на основе connection_side, если connection не было определено по VK/K или если side указывает иначе
        if connection_side == 'ra' and connection.startswith("VK"):
            connection = "VK-правое"
        elif connection_side == 're' and connection.startswith("VK"):
            connection = "VK-левое"
        elif connection_side == 'ls' and connection.startswith("K"):
            connection = "K-левое"
        elif connection_side == 'rs' and connection.startswith("K"):
            connection = "K-правое"

        # Если connection_side задан, но connection по умолчанию, используем side для уточнения
        if connection_side and connection == "VK-правое": # по умолчанию
            if connection_side == 're':
                connection = "VK-левое"
        elif connection_side and connection.startswith("K"): # по умолчанию K-боковое или K-правое/левое
            if connection_side == 'ls':
                connection = "K-левое"
            elif connection_side == 'rs':
                connection = "K-правое"
            # 'ra' и 're' для K обычно не используются, но если встречаются, можно оставить как есть или уточнить логику

        return {
            'brand': brand,
            'model': model,
            'type': type_,
            'height': height,
            'length': length,
            'connection': connection,
            'connection_side': connection_side,
            'original_name': name,
            'designation': designation
        }

    @staticmethod
    def extract_params_flexibly(name: str) -> Dict[str, Any]:
        """
        Старая версия парсера для обратной совместимости
        """
        result = {
            'recognized': False,
            'connection': None,
            'type': None,
            'height': None,
            'length': None
        }
        
        name_lower = name.lower()
        
        # Попытка извлечь тип, высоту, длину — три числа подряд
        match = re.search(r"(10|11|20|22|30|33)[-_/x\s]+(300|400|500|600|900)[-_/x\s]+(\d{3,4})", name_lower)
        if match:
            rad_type = match.group(1)
            height = int(match.group(2))
            length = int(match.group(3))
            
            # Определение подключения
            if "vk" in name_lower or "valve" in name_lower or "compact" in name_lower:
                connection = "VK-правое"
            else:
                connection = "K-боковое"
                
            result.update({
                'recognized': True,
                'connection': connection,
                'type': rad_type,
                'height': height,
                'length': length
            })
            
        return result