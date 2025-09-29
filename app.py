# app.py
from flask import Flask, render_template, request, send_file, flash, jsonify, redirect, url_for, session
from flask_wtf.csrf import CSRFProtect
import os
import json

# Импорты из наших модулей
from app.helpers import check_templates_exist, validate_form_data
from app.calculate import calculate_selling_price
from app.database import init_db, get_generation_history, save_generation_history, get_generation_details, load_generation_data
from app.document_generator import generate_excel_document, generate_word_document, create_zip_archive
from app.config import load_config, setup_app_security, setup_logging


app = Flask(__name__)

# Настройка приложения
config = load_config(app)
setup_app_security(app, config)  # Передаем конфиг
setup_logging(app, config)       # Передаем конфиг

# CSRF защита
csrf = CSRFProtect()
csrf.init_app(app)


@app.route('/')
def index():
    form_data = session.pop('form_data', {})
    for field in ['id', 'timestamp', 'final_price']:
        if field in form_data:
            del form_data[field]
    return render_template('index.html', form_data=form_data)

@app.route('/history/details/<int:record_id>')
def history_details(record_id):
    try:
        record = get_generation_details(record_id)
        if record:
            return jsonify(record)
        else:
            return jsonify({'error': 'Record not found'}), 404
    except Exception as e:
        app.logger.error(f'Error getting history details: {str(e)}')
        return jsonify({'error': 'Internal server error'}), 500

@app.route('/history')
def history():
    history_data = get_generation_history(config)
    return render_template('history.html', history=history_data)

@app.route('/load_generation/<int:gen_id>')
def load_generation(gen_id):
    try:
        generation_dict = load_generation_data(gen_id)
        if generation_dict:
            session['form_data'] = generation_dict
            return redirect(url_for('index'))
        else:
            flash('Запись не найдена.', 'danger')
            return redirect(url_for('history'))
    except Exception as e:
        app.logger.error(f'Error loading generation: {str(e)}')
        flash('Произошла ошибка при загрузке данных.', 'danger')
        return redirect(url_for('history'))

@app.route('/generate', methods=['POST'])
def generate():
    form_data = request.form.to_dict()
    errors = validate_form_data(form_data)
    
    if errors:
        for error in errors:
            flash(error, 'danger')
        return render_template('index.html', form_data=form_data)
    
    try:
        # Извлечение данных
        company = form_data['company'].strip()
        quantity = int(form_data['quantity'])
        cost_price = float(form_data['cost_price'])
        weight = float(form_data['weight'])
        logistics = float(form_data['logistics'])
        margin_percent = float(form_data['margin_percent'])
        delivery_time = int(form_data['delivery_time'])
        duty_percent = float(form_data.get('duty_percent', 0))
        
        # Расчет цены
        final_price = calculate_selling_price(
            quantity=quantity, 
            purchase_cost=cost_price, 
            logistics_rub=logistics,
            duty_percent=duty_percent, 
            weight=weight, 
            delivery_time=delivery_time,
            margin_percent=margin_percent,
            config=config  # Добавляем передачу конфига
        )

        general_prise = final_price * quantity  # Общая цена за количество
        final_price_NDS = general_prise * 1.2   # Общая цена с НДС
        
        # Сохранение в историю
        if not save_generation_history(form_data, final_price, config):
            app.logger.warning('Failed to save generation history')
        
        # Проверка шаблонов
        template_errors = check_templates_exist()
        if template_errors:
            for error in template_errors:
                flash(f'{error}. Обратитесь к администратору.', 'danger')
                app.logger.error(error)
            return render_template('index.html', form_data=form_data)
        
        # Генерация документов
        excel_template_path = os.path.join('templates_docs', 'template.xlsx')
        word_template_path = os.path.join('templates_docs', 'template.docx')
        
        # Передаем general_prise в функции генерации
        excel_file = generate_excel_document(excel_template_path, form_data, final_price, general_prise)
        word_file = generate_word_document(word_template_path, form_data, final_price, general_prise, final_price_NDS)
        zip_buffer, file_prefix = create_zip_archive(excel_file, word_file, company)
        
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
    

@app.route('/logistics')
def logistics():
    """Страница расчета логистики"""
    try:
        # Используем правильный путь к проекту
        current_dir = os.path.dirname(os.path.abspath(__file__))
        project_root = current_dir  # Теперь project_root = kp_generator-main
        logistics_path = os.path.join(project_root, 'config', 'logistics_cities.json')
        
        app.logger.info(f"Looking for logistics data at: {logistics_path}")
        
        if not os.path.exists(logistics_path):
            app.logger.warning(f"Logistics file not found: {logistics_path}")
            # Создаем базовый файл, если не существует
            default_cities = {
                "cities": [
                    {"name": "Москва", "price": 1100000, "region": "Центральный"},
                    {"name": "Екатеринбург", "price": 1000000, "region": "Уральский"}
                ]
            }
            os.makedirs(os.path.dirname(logistics_path), exist_ok=True)
            with open(logistics_path, 'w', encoding='utf-8') as f:
                json.dump(default_cities, f, ensure_ascii=False, indent=2)
            app.logger.info("Created default logistics file")
            cities = default_cities["cities"]
        else:
            with open(logistics_path, 'r', encoding='utf-8') as f:
                logistics_data = json.load(f)
            cities = logistics_data.get('cities', [])
            app.logger.info(f"Loaded {len(cities)} cities from logistics data")
            
    except Exception as e:
        app.logger.error(f'Error loading logistics data: {e}')
        # Возвращаем пустой список в случае ошибки
        cities = []
    
    return render_template('logistics.html', cities=cities)

    
@app.errorhandler(404)
def not_found_error(error):
    return render_template('404.html'), 404

@app.errorhandler(500)
def internal_error(error):
    return render_template('500.html'), 500

if __name__ == '__main__':
    # Создаем необходимые папки
    for folder in ['logs', 'templates_docs']:
        if not os.path.exists(folder):
            os.makedirs(folder)
    
    # Инициализируем БД
    init_db()
    
    app.logger.info('KP Generator started successfully')
    app.run(debug=True, host='0.0.0.0', port=5000)