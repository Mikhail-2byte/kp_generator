import os
import re
import json
from typing import List, Dict, Any

def check_templates_exist():
    """Проверяет существование шаблонов документов перед генерацией"""
    excel_template_path = os.path.join('templates_docs', 'template.xlsx')
    word_template_path = os.path.join('templates_docs', 'template.docx')
    
    errors = []
    if not os.path.exists(excel_template_path):
        errors.append(f'Excel template not found: {excel_template_path}')
    if not os.path.exists(word_template_path):
        errors.append(f'Word template not found: {word_template_path}')
    
    return errors


def get_safe_filename(company_name):
    """Создает безопасное имя файла из названия компании"""
    safe_name = re.sub(r'[^\w\s-]', '', company_name).strip()
    safe_name = re.sub(r'[-\s]+', '_', safe_name)
    return safe_name[:50]


def extract_positions_from_form(form_data: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Извлекает множественные позиции из данных формы."""
    positions = []
    
    # Поля позиций
    position_fields = ['product', 'drawing_number', 'material', 'cost_price', 'cost_price_per_kg', 'quantity', 'weight', 'duty_percent']
    
    # Проверяем, есть ли данные для множественных позиций
    # Ищем поля с суффиксами _2, _3, etc.
    position_numbers = set()
    
    for key in form_data.keys():
        for field in position_fields:
            if key == field:
                position_numbers.add(1)  # Первая позиция без суффикса
            elif key.startswith(field + '_') and key[len(field)+1:].isdigit():
                position_numbers.add(int(key[len(field)+1:]))
    
    # Если нет множественных позиций, создаем одну из основных полей
    if not position_numbers:
        position = {}
        for field in position_fields:
            if field in form_data and form_data[field]:
                position[field] = form_data[field]
        if position:
            positions.append(position)
        return positions
    
    # Собираем данные для каждой позиции
    for pos_num in sorted(position_numbers):
        position = {}
        for field in position_fields:
            if pos_num == 1:
                key = field
            else:
                key = f"{field}_{pos_num}"
            
            if key in form_data and form_data[key]:
                position[field] = form_data[key]
        
        if position:  # Добавляем только если есть данные
            positions.append(position)
    
    return positions


def validate_form_data(form_data):
    """Проверяет корректность данных формы"""
    errors = []
    
    # Извлекаем позиции
    positions = extract_positions_from_form(form_data)
    
    if not positions:
        errors.append('Необходимо указать хотя бы одну позицию')
        return errors
    
    # Проверка обязательных полей для каждой позиции
    required_position_fields = ['product', 'quantity', 'cost_price', 'weight']
    
    for i, position in enumerate(positions, 1):
        for field in required_position_fields:
            if not position.get(field) or not str(position[field]).strip():
                errors.append(f'Позиция {i}: поле "{field}" является обязательным.')
    
    # Проверка общих обязательных полей
    general_required_fields = ['company', 'logistics', 'margin_percent', 'delivery_time']
    for field in general_required_fields:
        if not form_data.get(field) or not form_data[field].strip():
            errors.append(f'Поле "{field}" является обязательным.')
    
    # Проверка числовых значений для каждой позиции
    numeric_position_fields = ['cost_price', 'cost_price_per_kg', 'weight', 'duty_percent', 'quantity']
    
    for i, position in enumerate(positions, 1):
        for field in numeric_position_fields:
            if position.get(field) and str(position[field]).strip():
                try:
                    value = float(position[field])
                    if value < 0:
                        errors.append(f'Позиция {i}: поле "{field}" должно быть неотрицательным числом.')
                    if field in ['duty_percent'] and value > 100:
                        errors.append(f'Позиция {i}: поле "{field}" не может превышать 100%.')
                    if field == 'quantity' and value == 0:
                        errors.append(f'Позиция {i}: поле "{field}" не может быть нулевым.')
                except ValueError:
                    errors.append(f'Позиция {i}: поле "{field}" должно быть числом.')
    
    # Проверка общих числовых полей
    general_numeric_fields = ['logistics', 'delivery_time', 'margin_percent']
    for field in general_numeric_fields:
        if form_data.get(field) and form_data[field].strip():
            try:
                value = float(form_data[field])
                if value < 0:
                    errors.append(f'Поле "{field}" должно быть неотрицательным числом.')
                if field in ['margin_percent'] and value > 100:
                    errors.append(f'Поле "{field}" не может превышать 100%.')
                if field == 'delivery_time' and value < 1:
                    errors.append(f'Поле "Срок поставки" не может быть меньше 1 дня.')
            except ValueError:
                errors.append(f'Поле "{field}" должно быть числом.')
    
    return errors