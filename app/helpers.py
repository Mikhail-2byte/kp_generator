
import os
import re

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


def validate_form_data(form_data):
    """Проверяет корректность данных формы"""
    errors = []
    
    # Проверка обязательных полей
    required_fields = ['company', 'product', 'quantity', 'cost_price', 'weight', 'logistics', 'margin_percent', 'delivery_time']
    for field in required_fields:
        if not form_data.get(field) or not form_data[field].strip():
            errors.append(f'Поле "{field}" является обязательным.')
    
    # Проверка числовых значений
    numeric_fields = ['quantity', 'cost_price', 'weight', 'logistics', 'duty_percent', 'delivery_time', 'margin_percent']
    for field in numeric_fields:
        if form_data.get(field) and form_data[field].strip():
            try:
                value = float(form_data[field])
                if value < 0:
                    errors.append(f'Поле "{field}" должно быть неотрицательным числом.')
                if field in ['duty_percent', 'margin_percent'] and value > 100:
                    errors.append(f'Поле "{field}" не может превышать 100%.')
                if field == 'quantity' and value == 0:
                    errors.append(f'Поле "{field}" не может быть нулевым.')
                if field == 'delivery_time' and value < 1:
                    errors.append(f'Поле "Срок поставки" не может быть меньше 1 дня.')
            except ValueError:
                errors.append(f'Поле "{field}" должно быть числом.')
    
    return errors