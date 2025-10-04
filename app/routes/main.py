from datetime import datetime

from flask import (
    Blueprint,
    current_app,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for
)
from flask_login import current_user

from app.calculate import calculate_selling_price
from app.document_generator import create_zip_archive, generate_excel_document, generate_word_document
from app.helpers import check_templates_exist, validate_form_data
from app.services import datasets
from app.services.repositories import generation_repository
from app.services.feedback import save_feedback_entry
from app.ui import build_context


main_bp = Blueprint('main', __name__)  # Основные страницы и бизнес-логика генератора КП


@main_bp.route('/')
def index():
    """Показывает форму расчёта коммерческого предложения и заполняет справочники."""
    form_data = session.pop('form_data', {})
    for field in ['id', 'timestamp', 'final_price', 'user_id']:
        form_data.pop(field, None)

    try:
        cities = datasets.load_logistics_cities()
    except Exception as exc:  # pragma: no cover - defensive logging
        current_app.logger.error('Error loading logistics data for index: %s', exc)
        cities = []

    return render_template(
        'index.html',
        **build_context('index', 'Создание коммерческого предложения', form_data=form_data, cities=cities)
    )


@main_bp.route('/history/details/<int:record_id>')
def history_details(record_id):
    """Возвращает детальную информацию о сохранённой генерации в формате JSON."""
    try:
        record = generation_repository.get_details(record_id)
        if record:
            return jsonify(record)
        return jsonify({'error': 'Record not found'}), 404
    except Exception as exc:  # pragma: no cover - log unexpected failures
        current_app.logger.error('Error getting history details: %s', exc)
        return jsonify({'error': 'Internal server error'}), 500


@main_bp.route('/history')
def history():
    """Отображает список последних генераций пользователя."""
    app_config = current_app.config['APP_SETTINGS']
    history_data = generation_repository.get_history(app_config)
    return render_template(
        'history.html',
        **build_context('history', 'История генераций КП', history=history_data)
    )


@main_bp.route('/feedback', methods=['GET', 'POST'])
def feedback():
    """Принимает обратную связь от пользователей и сохраняет её локально."""
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
            return render_template(
                'feedback.html',
                **build_context('feedback', 'Обратная связь', form_data=form_data)
            )

        entry = {
            'timestamp': datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ'),
            'name': name or 'Аноним',
            'contact': contact,
            'feedback': feedback_text,
            'improvement': improvement_text
        }

        if save_feedback_entry(entry):
            flash('Спасибо! Ваш отзыв успешно отправлен.', 'success')
            return redirect(url_for('main.feedback'))
        flash('Не удалось сохранить отзыв. Попробуйте позже.', 'danger')
        return render_template(
            'feedback.html',
            **build_context('feedback', 'Обратная связь', form_data=form_data)
        )

    return render_template(
        'feedback.html',
        **build_context('feedback', 'Обратная связь', form_data=form_data)
    )


@main_bp.route('/gb-analogs')
def gb_analogs():
    """Показывает таблицу аналогов материалов по китайскому стандарту GB."""
    query = request.args.get('q', '').strip()
    normalized_query = query.lower()
    filtered_materials = datasets.get_gb_materials()

    if normalized_query:
        filtered = []
        for material in datasets.get_gb_materials():
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
                filtered.append(material)
        filtered_materials = filtered

    return render_template(
        'gb_analogs.html',
        **build_context('gb', 'Аналоги по стандарту GB', materials=filtered_materials, query=query)
    )


@main_bp.route('/orders')
def orders_page():
    """Отображает раздел с распоряжениями и внутренними документами."""
    return render_template(
        'orders.html',
        **build_context('orders', 'Распоряжения')
    )


@main_bp.route('/templates-library')
def templates_page():
    """Выводит список шаблонов документов."""
    return render_template(
        'templates_page.html',
        **build_context('templates', 'Шаблоны')
    )


@main_bp.route('/instructions')
def instructions_page():
    """Содержит краткие инструкции по бизнес-процессам."""
    return render_template(
        'instructions.html',
        **build_context('instructions', 'Инструкция')
    )


@main_bp.route('/duty')
def duty():
    """Предоставляет поиск по ставкам пошлин и категориям товаров."""
    query = request.args.get('q', '').strip()
    normalized_query = query.lower()
    filtered_items = datasets.get_duty_rates()

    if normalized_query:
        filtered_items = [
            item for item in datasets.get_duty_rates()
            if normalized_query in item.get('product_search', '')
            or normalized_query in item.get('category_search', '')
            or normalized_query in item.get('duty_search', '')
        ]

    return render_template(
        'duty.html',
        **build_context('duty', 'Ставки пошлин', items=filtered_items, query=query)
    )


@main_bp.route('/load_generation/<int:gen_id>')
def load_generation(gen_id):
    """Загружает ранее рассчитанную генерацию в форму для повторного использования."""
    try:
        generation_dict = generation_repository.load_generation(gen_id)
        if generation_dict:
            session['form_data'] = generation_dict
            return redirect(url_for('main.index'))
        flash('Запись не найдена.', 'danger')
        return redirect(url_for('main.history'))
    except Exception as exc:  # pragma: no cover - defensive logging
        current_app.logger.error('Error loading generation: %s', exc)
        flash('Произошла ошибка при загрузке данных.', 'danger')
        return redirect(url_for('main.history'))


@main_bp.route('/generate', methods=['POST'])
def generate():
    """Выполняет расчёт КП, сохраняет историю и формирует пакет документов."""
    form_data = request.form.to_dict()
    form_data['comment'] = form_data.get('comment', '').strip()
    errors = validate_form_data(form_data)

    if errors:
        for error in errors:
            flash(error, 'danger')
        return render_template(
            'index.html',
            **build_context('index', 'Создание коммерческого предложения', form_data=form_data)
        )

    try:
        company = form_data['company'].strip()
        quantity = int(form_data['quantity'])
        cost_price = float(form_data['cost_price'])
        weight = float(form_data['weight'])
        logistics_rub = float(form_data['logistics'])
        margin_percent = float(form_data['margin_percent'])
        delivery_time = int(form_data['delivery_time'])
        duty_percent = float(form_data.get('duty_percent', 0))

        app_config = current_app.config['APP_SETTINGS']

        final_price = calculate_selling_price(
            quantity=quantity,
            purchase_cost=cost_price,
            logistics_rub=logistics_rub,
            duty_percent=duty_percent,
            weight=weight,
            delivery_time=delivery_time,
            margin_percent=margin_percent,
            config=app_config
        )

        general_price = final_price * quantity
        final_price_nds = general_price * 1.2

        user_id = int(current_user.id) if current_user.is_authenticated else None
        if not generation_repository.save_history(form_data, final_price, app_config, user_id):
            current_app.logger.warning('Failed to save generation history')

        template_errors = check_templates_exist()
        if template_errors:
            for error in template_errors:
                flash(f'{error}. Обратитесь к администратору.', 'danger')
                current_app.logger.error(error)
            return render_template(
                'index.html',
                **build_context('index', 'Создание коммерческого предложения', form_data=form_data)
            )

        excel_template_path = 'templates_docs/template.xlsx'
        word_template_path = 'templates_docs/template.docx'

        excel_file = generate_excel_document(excel_template_path, form_data, final_price, general_price)
        word_file = generate_word_document(word_template_path, form_data, final_price, general_price, final_price_nds)
        zip_buffer, file_prefix = create_zip_archive(excel_file, word_file, company)

        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name=f'{file_prefix}.zip',
            mimetype='application/zip'
        )

    except Exception as exc:  # pragma: no cover - defensive logging
        flash('Произошла непредвиденная ошибка. Попробуйте еще раз.', 'danger')
        current_app.logger.error('Unexpected error: %s', exc)
        return render_template(
            'index.html',
            **build_context('index', 'Создание коммерческого предложения', form_data=form_data)
        )


@main_bp.route('/logistics')
def logistics():
    """Отображает справочную информацию по логистическим тарифам."""
    try:
        cities = datasets.load_logistics_cities()
        if not cities:
            current_app.logger.info('Logistics data is empty or missing')
    except Exception as exc:  # pragma: no cover - defensive logging
        current_app.logger.error('Error loading logistics data: %s', exc)
        cities = []

    return render_template(
        'logistics.html',
        **build_context('logistics', 'Просчёт логистики', cities=cities)
    )
