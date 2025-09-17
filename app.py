from flask import Flask, render_template, request, send_file, flash, send_from_directory, jsonify
from openpyxl import load_workbook
from docx import Document
import os
from io import BytesIO
import logging
from logging.handlers import RotatingFileHandler
import zipfile
import re
from datetime import datetime
import json

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'your-secret-key-here')

# Настройка логирования
if not os.path.exists('logs'):
    os.makedirs('logs')
file_handler = RotatingFileHandler('logs/kp_generator.log', maxBytes=10240, backupCount=10)
file_handler.setFormatter(logging.Formatter(
    '%(asctime)s %(levelname)s: %(message)s [in %(pathname)s:%(lineno)d]'))
file_handler.setLevel(logging.INFO)
app.logger.addHandler(file_handler)
app.logger.setLevel(logging.INFO)
app.logger.info('KP Generator startup')

def validate_form_data(form_data):
    """Проверяет корректность данных формы"""
    errors = []
    
    # Проверка обязательных полей
    required_fields = ['company', 'product', 'quantity', 'cost_price', 'weight', 'logistics']
    for field in required_fields:
        if not form_data.get(field) or not form_data[field].strip():
            errors.append(f'Поле "{field}" является обязательным.')
    
    # Проверка числовых значений
    numeric_fields = ['quantity', 'cost_price', 'weight', 'logistics', 'duty_percent']
    for field in numeric_fields:
        if form_data.get(field) and form_data[field].strip():
            try:
                value = float(form_data[field])
                if value < 0:
                    errors.append(f'Поле "{field}" должно быть неотрицательным числом.')
                if field == 'duty_percent' and value > 100:
                    errors.append(f'Поле "{field}" не может превышать 100%.')
                if field == 'quantity' and value == 0:
                    errors.append(f'Поле "{field}" не может быть нулевым.')
            except ValueError:
                errors.append(f'Поле "{field}" должно быть числом.')
    
    return errors

def get_safe_filename(company_name):
    """Создает безопасное имя файла из названия компании"""
    safe_name = re.sub(r'[^\w\s-]', '', company_name).strip()
    safe_name = re.sub(r'[-\s]+', '_', safe_name)
    return safe_name[:50]

def calculate_prices(quantity, cost_price, logistics, duty_percent, margin_percent):
    """Выполняет все необходимые расчеты цен"""
    total_cost = quantity * cost_price
    duty_amount = total_cost * (duty_percent / 100)
    total_with_duty = total_cost + duty_amount
    cost_per_unit = (total_with_duty + logistics) / quantity
    final_price = cost_per_unit / (1 - margin_percent / 100)
    
    return {
        'total_cost': total_cost,
        'duty_amount': duty_amount,
        'total_with_duty': total_with_duty,
        'cost_per_unit': cost_per_unit,
        'final_price': final_price
    }

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/favicon.ico')
def favicon():
    return send_from_directory(os.path.join(app.root_path, 'static'),
                               'favicon.ico', mimetype='image/vnd.microsoft.icon')

@app.route('/generate', methods=['POST'])
def generate():
    form_data = request.form.to_dict()
    errors = validate_form_data(form_data)
    
    if errors:
        for error in errors:
            flash(error, 'danger')
        return render_template('index.html', form_data=form_data)
    
    try:
        # Извлечение и преобразование данных
        company = form_data['company'].strip()
        product = form_data['product'].strip()
        quantity = int(form_data['quantity'])
        cost_price = float(form_data['cost_price'])
        weight = float(form_data['weight'])
        logistics = float(form_data['logistics'])
        
        # Новые поля
        tender_number = form_data.get('tender_number', '').strip()
        drawing_number = form_data.get('drawing_number', '').strip()
        material = form_data.get('material', '').strip()
        delivery_address = form_data.get('delivery_address', '').strip()
        duty_percent = float(form_data.get('duty_percent', 0))
        
        # Настройки
        margin_percent = 20  # 20% наценка
        
        # Выполнение расчетов
        calculations = calculate_prices(quantity, cost_price, logistics, duty_percent, margin_percent)
        final_price = calculations['final_price']
        
        # Проверка существования шаблонов
        excel_template_path = os.path.join('templates_docs', 'template.xlsx')
        word_template_path = os.path.join('templates_docs', 'template.docx')
        
        if not os.path.exists(excel_template_path):
            flash('Шаблон Excel не найден. Обратитесь к администратору.', 'danger')
            app.logger.error(f'Excel template not found: {excel_template_path}')
            return render_template('index.html', form_data=form_data)
            
        if not os.path.exists(word_template_path):
            flash('Шаблон Word не найден. Обратитесь к администратору.', 'danger')
            app.logger.error(f'Word template not found: {word_template_path}')
            return render_template('index.html', form_data=form_data)
        
        # Работа с Excel
        try:
            wb = load_workbook(excel_template_path)
            ws = wb.active
            
            # Заполняем данные в Excel
            current_date = datetime.now().strftime('%d.%m.%Yг.')
            
            # Основные данные
            ws['D4'] = company
            ws['G10'] = quantity
            ws['M10'] = cost_price
            ws['P10'] = weight
            ws['U14'] = logistics
            ws['X10'] = duty_percent / 100  # Процент пошлины (доля)
            
            # Дополнительные поля
            ws['D2'] = current_date  # Дата формирования
            ws['D5'] = tender_number  # Номер тендера
            ws['D10'] = material  # Материал
            ws['E10'] = drawing_number  # Номер чертежа
            ws['P4'] = delivery_address  # Адрес доставки
            
            # Расчетные поля
            ws['H10'] = final_price  # Финальная цена
            
            excel_file = BytesIO()
            wb.save(excel_file)
            excel_file.seek(0)
        except Exception as e:
            flash('Ошибка при обработке Excel-шаблона.', 'danger')
            app.logger.error(f'Excel processing error: {str(e)}')
            return render_template('index.html', form_data=form_data)
        
        # Работа с Word
        try:
            doc = Document(word_template_path)
            
            # Форматируем текущую дату
            current_date = datetime.now().strftime('%d.%m.%Yг.')
            
            word_data = {
                '{{ company }}': company,
                '{{ product }}': product,
                '{{ quantity }}': str(quantity),
                '{{ cost_price }}': f"{cost_price:.2f}",
                '{{ weight }}': f"{weight:.2f}",
                '{{ logistics }}': f"{logistics:.2f}",
                '{{ final_price }}': f"{final_price:.2f}",
                '{{ tender_number }}': tender_number,
                '{{ drawing_number }}': drawing_number,
                '{{ material }}': material,
                '{{ delivery_address }}': delivery_address,
                '{{ date }}': current_date,
                '{{ duty_percent }}': f"{duty_percent:.1f}",
            }
            
            # Замена в параграфах
            for paragraph in doc.paragraphs:
                for key, value in word_data.items():
                    if key in paragraph.text:
                        paragraph.text = paragraph.text.replace(key, value)
            
            # Замена в таблицах
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for key, value in word_data.items():
                            if key in cell.text:
                                cell.text = cell.text.replace(key, value)
            
            word_file = BytesIO()
            doc.save(word_file)
            word_file.seek(0)
        except Exception as e:
            flash('Ошибка при обработке Word-шаблона.', 'danger')
            app.logger.error(f'Word processing error: {str(e)}')
            return render_template('index.html', form_data=form_data)
        
        # Создание ZIP-архива
        file_prefix = f"КП_{get_safe_filename(company)}_{datetime.now().strftime('%Y%m%d_%H%M')}"
        
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            # Добавляем Excel-файл
            zip_file.writestr(f"{file_prefix}.xlsx", excel_file.getvalue())
            # Добавляем Word-файл
            zip_file.writestr(f"{file_prefix}.docx", word_file.getvalue())
        
        zip_buffer.seek(0)
        
        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name=f"{file_prefix}.zip",
            mimetype='application/zip'
        )
    
    except Exception as e:
        flash('Произошла непредвиденная ошибка. Попробуйте еще раз.', 'danger')
        app.logger.error(f'Unexpected error: {str(e)}')
        return render_template('index.html', form_data=form_data)

@app.errorhandler(404)
def not_found_error(error):
    return render_template('404.html'), 404

@app.errorhandler(500)
def internal_error(error):
    return render_template('500.html'), 500

if __name__ == '__main__':
    # Создаем необходимые папки при запуске
    for folder in ['logs', 'templates_docs']:
        if not os.path.exists(folder):
            os.makedirs(folder)
    
    app.run(debug=True, host='0.0.0.0', port=5000)