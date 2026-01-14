import json
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Tuple


def check_templates_exist() -> List[str]:
    """Проверяет существование шаблонов документов перед генерацией."""
    excel_template_path = os.path.join('templates_docs', 'template.xlsx')
    word_template_path = os.path.join('templates_docs', 'template.docx')

    errors = []
    if not os.path.exists(excel_template_path):
        errors.append(f'Excel template not found: {excel_template_path}')
    if not os.path.exists(word_template_path):
        errors.append(f'Word template not found: {word_template_path}')

    return errors


def get_safe_filename(company_name: str) -> str:
    """
    Создаёт безопасное имя файла из названия компании.
    
    Защищает от path traversal атак и других небезопасных символов.
    Удаляет пути (.., /, \), нормализует имя файла.
    
    Args:
        company_name: Исходное название компании
    
    Returns:
        Безопасное имя файла (максимум 50 символов)
    """
    if not company_name:
        return 'file'
    
    # Защита от path traversal: удаляем .., /, \
    safe_name = company_name.replace('..', '').replace('/', '').replace('\\', '')
    
    # Удаляем все символы кроме букв, цифр, пробелов и дефисов
    safe_name = re.sub(r'[^\w\s-]', '', safe_name)
    
    # Удаляем ведущие/завершающие пробелы и дефисы
    safe_name = safe_name.strip(' \t\n\r-_')
    
    # Заменяем последовательности пробелов и дефисов на одно подчеркивание
    safe_name = re.sub(r'[-\s]+', '_', safe_name)
    
    # Убеждаемся, что имя не пустое
    if not safe_name:
        return 'file'
    
    # Дополнительная защита: используем только имя файла (без пути)
    # Это защищает от случаев, когда имя начинается с пути
    safe_name = Path(safe_name).name
    
    # Ограничиваем длину
    return safe_name[:50]


def extract_positions_from_form(
    form_data: Dict[str, Any],
    include_field_keys: bool = False
) -> List[Dict[str, Any]] | Tuple[List[Dict[str, Any]], List[Dict[str, str]]]:
    """Извлекает множественные позиции из данных формы."""

    positions: List[Dict[str, Any]] = []
    field_keys: List[Dict[str, str]] = []

    # Сначала проверяем positions_payload (JSON строка с позициями)
    positions_payload = form_data.get('positions_payload')
    if positions_payload:
        try:
            if isinstance(positions_payload, str):
                parsed_positions = json.loads(positions_payload)
            else:
                parsed_positions = positions_payload
            
            if isinstance(parsed_positions, list) and len(parsed_positions) > 0:
                # Создаем field_keys для каждой позиции
                for pos in parsed_positions:
                    if include_field_keys:
                        key_map = {field: field for field in ['product', 'drawing_number', 'material', 'cost_price', 'cost_price_per_kg', 'quantity', 'weight', 'duty_percent']}
                        field_keys.append(key_map)
                return (parsed_positions, field_keys) if include_field_keys else parsed_positions
        except (json.JSONDecodeError, TypeError, ValueError):
            # Если не удалось распарсить, продолжаем обычную обработку
            pass

    position_fields = [
        'product',
        'drawing_number',
        'material',
        'cost_price',
        'cost_price_per_kg',
        'quantity',
        'weight',
        'duty_percent'
    ]

    position_numbers: set[int] = set()

    for key in form_data.keys():
        for field in position_fields:
            if key == field:
                position_numbers.add(1)
            elif key.startswith(field + '_') and key[len(field) + 1:].isdigit():
                position_numbers.add(int(key[len(field) + 1:]))

    if not position_numbers:
        position = {}
        key_map = {}
        for field in position_fields:
            if field in form_data and form_data[field]:
                position[field] = form_data[field]
            if include_field_keys:
                key_map[field] = field
        if position:
            positions.append(position)
            if include_field_keys:
                field_keys.append(key_map)

        return (positions, field_keys) if include_field_keys else positions

    for pos_num in sorted(position_numbers):
        position = {}
        key_map = {}
        for field in position_fields:
            key = field if pos_num == 1 else f"{field}_{pos_num}"
            key_map[field] = key
            value = form_data.get(key)
            if value:
                position[field] = value

        if position:
            positions.append(position)
            if include_field_keys:
                field_keys.append(key_map)

    if include_field_keys:
        return positions, field_keys
    return positions
