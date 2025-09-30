# app.py
from flask import Flask, render_template, request, send_file, flash, jsonify, redirect, url_for, session
from flask_wtf import FlaskForm
from flask_wtf.csrf import CSRFProtect
from flask_login import (
    LoginManager,
    login_user,
    logout_user,
    login_required,
    current_user,
    UserMixin
)
import os
import json
from datetime import datetime
from wtforms import StringField, PasswordField, SubmitField, BooleanField
from wtforms.validators import DataRequired, Length, EqualTo
from werkzeug.security import generate_password_hash, check_password_hash

# Импорты из наших модулей
from app.helpers import check_templates_exist, validate_form_data
from app.calculate import calculate_selling_price
from app.database import (
    init_db,
    get_generation_history,
    save_generation_history,
    get_generation_details,
    load_generation_data,
    create_user,
    get_user_by_username,
    get_user_by_id,
    update_last_login,
    get_user_statistics
)
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

login_manager = LoginManager()
login_manager.login_view = 'profile'
login_manager.login_message_category = 'info'
login_manager.init_app(app)

FEEDBACK_FILE_PATH = os.path.join('data', 'feedback_entries.jsonl')
GB_ANALOGS = []
DUTY_RATES = []


def load_gb_materials():
    project_root = os.path.dirname(os.path.abspath(__file__))
    materials_path = os.path.join(project_root, 'config', 'gb_materials.json')
    try:
        with open(materials_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        materials = data.get('materials', [])

        for material in materials:
            composition = material.get('composition', [])
            material['composition_search'] = ' '.join(
                f"{item.get('element', '')} {item.get('content', '')}" for item in composition
            )

        return materials
    except FileNotFoundError:
        app.logger.error('GB materials file not found at %s', materials_path)
        return []
    except json.JSONDecodeError as exc:
        app.logger.error('Failed to parse GB materials file: %s', exc)
        return []


GB_ANALOGS = load_gb_materials()


def load_duty_rates():
    project_root = os.path.dirname(os.path.abspath(__file__))
    duty_path = os.path.join(project_root, 'config', 'duty_rates.json')
    try:
        with open(duty_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        items = data.get('items', [])

        for item in items:
            item['product_search'] = str(item.get('product', '')).lower()
            item['category_search'] = str(item.get('category', '')).lower()
            item['duty_search'] = str(item.get('duty_percent', '')).lower()

        return items
    except FileNotFoundError:
        app.logger.error('Duty rates file not found at %s', duty_path)
        return []
    except json.JSONDecodeError as exc:
        app.logger.error('Failed to parse duty rates file: %s', exc)
        return []


DUTY_RATES = load_duty_rates()


class User(UserMixin):
    def __init__(self, user_id, username, password_hash, created_at=None, last_login=None):
        self.id = str(user_id)
        self.username = username
        self.password_hash = password_hash
        self.created_at = created_at
        self.last_login = last_login

    @classmethod
    def from_row(cls, row):
        if not row:
            return None
        return cls(user_id=row[0], username=row[1], password_hash=row[2], created_at=row[3], last_login=row[4])


class LoginForm(FlaskForm):
    username = StringField('Логин', validators=[DataRequired(), Length(min=3, max=50)])
    password = PasswordField('Пароль', validators=[DataRequired(), Length(min=6, max=128)])
    remember_me = BooleanField('Запомнить меня')
    submit_login = SubmitField('Войти')


class RegistrationForm(FlaskForm):
    username = StringField('Логин', validators=[DataRequired(), Length(min=3, max=50)])
    password = PasswordField('Пароль', validators=[DataRequired(), Length(min=6, max=128)])
    confirm_password = PasswordField('Подтвердите пароль', validators=[DataRequired(), EqualTo('password', message='Пароли должны совпадать')])
    submit_register = SubmitField('Зарегистрироваться')


@login_manager.user_loader
def load_user(user_id):
    row = get_user_by_id(user_id)
    return User.from_row(row) if row else None


def save_feedback_entry(entry):
    try:
        feedback_dir = os.path.dirname(FEEDBACK_FILE_PATH)
        if feedback_dir and not os.path.exists(feedback_dir):
            os.makedirs(feedback_dir)
        with open(FEEDBACK_FILE_PATH, 'a', encoding='utf-8') as feedback_file:
            json.dump(entry, feedback_file, ensure_ascii=False)
            feedback_file.write('\n')
        return True
    except Exception as err:
        app.logger.error(f'Error saving feedback entry: {err}')
        return False


@app.route('/')
def index():
    form_data = session.pop('form_data', {})
    for field in ['id', 'timestamp', 'final_price', 'user_id']:
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


@app.route('/feedback', methods=['GET', 'POST'])
def feedback():
    form_data = {}

    if request.method == 'POST':
        name = request.form.get('name', '').strip()
        contact = request.form.get('contact', '').strip()
        feedback_text = request.form.get('feedback_text', '').strip()
        improvement_text = request.form.get('improvement_text', '').strip()

        form_data = {
            'name': name,
            'contact': contact,
            'feedback_text': feedback_text,
            'improvement_text': improvement_text
        }

        if not feedback_text and not improvement_text:
            flash('Пожалуйста, поделитесь отзывом или предложением.', 'danger')
            return render_template('feedback.html', form_data=form_data)

        entry = {
            'timestamp': datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ'),
            'name': name or 'Аноним',
            'contact': contact,
            'feedback': feedback_text,
            'improvement': improvement_text
        }

        if save_feedback_entry(entry):
            flash('Спасибо! Ваш отзыв успешно отправлен.', 'success')
            return redirect(url_for('feedback'))
        else:
            flash('Не удалось сохранить отзыв. Попробуйте позже.', 'danger')
            return render_template('feedback.html', form_data=form_data)

    return render_template('feedback.html', form_data=form_data)

@app.route('/gb-analogs')
def gb_analogs():
    query = request.args.get('q', '').strip()
    normalized_query = query.lower()
    filtered_materials = GB_ANALOGS

    if normalized_query:
        filtered_materials = []
        for material in GB_ANALOGS:
            composition_values = (
                material.get('composition_search', '')
                or ' '.join(
                    f"{comp.get('element', '')} {comp.get('content', '')}"
                    for comp in material.get('composition', [])
                )
            ).lower()

            if (
                normalized_query in material['russian'].lower()
                or normalized_query in material['gb'].lower()
                or normalized_query in material.get('notes', '').lower()
                or normalized_query in composition_values
            ):
                filtered_materials.append(material)

    return render_template('gb_analogs.html', materials=filtered_materials, query=query)


@app.route('/duty')
def duty():
    query = request.args.get('q', '').strip()
    normalized_query = query.lower()
    filtered_items = DUTY_RATES

    if normalized_query:
        filtered_items = [
            item for item in DUTY_RATES
            if normalized_query in item.get('product_search', '')
            or normalized_query in item.get('category_search', '')
            or normalized_query in item.get('duty_search', '')
        ]

    return render_template('duty.html', items=filtered_items, query=query)

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
    form_data['comment'] = form_data.get('comment', '').strip()
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
        user_id = int(current_user.id) if current_user.is_authenticated else None
        if not save_generation_history(form_data, final_price, config, user_id):
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
    

@app.route('/profile', methods=['GET', 'POST'])
def profile():
    login_form = LoginForm(prefix='login')
    register_form = RegistrationForm(prefix='register')

    if request.method == 'POST':
        if login_form.submit_login.data and login_form.validate():
            username = login_form.username.data.strip()
            login_form.username.data = username
            user_row = get_user_by_username(username)
            if user_row and check_password_hash(user_row[2], login_form.password.data):
                user = User.from_row(user_row)
                login_user(user, remember=login_form.remember_me.data)
                update_last_login(int(user.id))
                user.last_login = datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')
                flash('Вы успешно вошли в систему.', 'success')
                return redirect(url_for('profile'))
            else:
                flash('Неверный логин или пароль.', 'danger')
        elif register_form.submit_register.data and register_form.validate():
            username = register_form.username.data.strip()
            register_form.username.data = username
            if get_user_by_username(username):
                register_form.username.errors.append('Пользователь с таким логином уже существует.')
            else:
                password_hash = generate_password_hash(register_form.password.data.strip())
                new_user_id = create_user(username, password_hash)
                if new_user_id:
                    user = User.from_row(get_user_by_id(new_user_id))
                    login_user(user)
                    update_last_login(int(user.id))
                    user.last_login = datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')
                    flash('Аккаунт успешно создан.', 'success')
                    return redirect(url_for('profile'))
                else:
                    flash('Не удалось создать пользователя. Попробуйте позже.', 'danger')

    stats = None
    if current_user.is_authenticated:
        stats = get_user_statistics(int(current_user.id))
        last_gen = stats.get('last_generation_at')
        if last_gen:
            try:
                stats['last_generation_at_formatted'] = datetime.strptime(last_gen, '%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y %H:%M')
            except ValueError:
                stats['last_generation_at_formatted'] = last_gen
        else:
            stats['last_generation_at_formatted'] = None

        formatted_recent = []
        for record in stats.get('recent_generations', []):
            record_id, company, product, margin_percent, final_price, timestamp_value = record
            try:
                formatted_ts = datetime.strptime(timestamp_value, '%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y %H:%M')
            except ValueError:
                formatted_ts = timestamp_value
            formatted_recent.append({
                'id': record_id,
                'company': company,
                'product': product,
                'margin_percent': margin_percent,
                'final_price': final_price,
                'timestamp': formatted_ts
            })
        stats['recent_generations'] = formatted_recent

    return render_template('profile.html', login_form=login_form, register_form=register_form, stats=stats)


@app.route('/logout')
@login_required
def logout():
    logout_user()
    flash('Вы вышли из системы.', 'info')
    return redirect(url_for('profile'))


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