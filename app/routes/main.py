import json
import re
from collections import OrderedDict
from datetime import datetime, timezone

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

from app.business.price_calculator import calculate_selling_price
from app.business.document_generator import create_zip_archive, generate_excel_document, generate_word_document
from app.presentation.helpers import check_templates_exist, extract_positions_from_form
from app.presentation.validators import validate_form_data
from app.services.multi_position_calculator import MultiPositionCalculator
from app.services import (
    AnalyticsProcessingError,
    analyze_excel,
    datasets
)
from app.services.repositories import generation_repository
from app.services.feedback import save_feedback_entry
from app.services.excel_importer import ExcelImportError, parse_positions_from_excel
from app.presentation.ui import build_context
from app.core.extensions import csrf


main_bp = Blueprint('main', __name__)  # Основные страницы и бизнес-логика генератора КП


@main_bp.route('/')
def index():
    """Показывает форму расчёта коммерческого предложения и заполняет справочники."""
    form_data = session.pop('form_data', {})
    for field in ['id', 'timestamp', 'final_price', 'user_id']:
        form_data.pop(field, None)

    imported_positions = session.pop('imported_positions', None)
    if imported_positions:
        form_data.setdefault('positions_payload', json.dumps(imported_positions, ensure_ascii=False))
        first_position = imported_positions[0]
        for key in ['product', 'drawing_number', 'material', 'cost_price', 'quantity', 'weight', 'duty_percent']:
            if first_position.get(key):
                form_data.setdefault(key, first_position[key])

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


@main_bp.route('/history/drawing')
def history_by_drawing():
    """Возвращает генерации с указанным номером чертежа."""
    drawing_number = request.args.get('number', '').strip()
    if not drawing_number:
        return jsonify({'matches': []})

    try:
        matches = generation_repository.get_by_drawing(drawing_number)
        return jsonify({'matches': matches})
    except Exception as exc:  # pragma: no cover - defensive logging
        current_app.logger.error('Error fetching drawing matches: %s', exc)
        return jsonify({'matches': []}), 500


@main_bp.route('/history/tender/<string:tender_number>')
def history_by_tender(tender_number):
    """Возвращает все версии генераций по номеру тендера."""
    if not tender_number:
        return jsonify({'records': []})
    records = generation_repository.get_by_tender(tender_number)
    return jsonify({'records': records})


def _build_tender_groups(items):
    groups = []
    lookup = OrderedDict()

    for item in items:
        tender_number = item.get('tender_number')
        if tender_number:
            key = tender_number.lower()
        else:
            key = f'record-{item["id"]}'

        group = lookup.get(key)
        if group is None:
            identifier_source = tender_number or f'{item["id"]}'
            identifier = re.sub(r'[^a-zA-Z0-9_-]+', '-', identifier_source).strip('-').lower() or f'group-{item["id"]}'
            group = {
                'identifier': identifier,
                'tender_number': tender_number,
                'display_tender': tender_number or 'Без номера тендера',
                'latest': item,
                'previous': [],
                'version_count': item.get('version_count', 1),
                'has_remote': bool(tender_number),
            }
            lookup[key] = group
            groups.append(group)
        else:
            group['previous'].append(item)

    for group in groups:
        group['previous'] = sorted(group['previous'], key=lambda x: x.get('timestamp', ''), reverse=True)
        if not group.get('version_count'):
            group['version_count'] = 1

    return groups


@main_bp.route('/history')
def history():
    """Отображает список последних генераций пользователя."""
    app_config = current_app.config['APP_SETTINGS']
    try:
        page = int(request.args.get('page', '1'))
    except ValueError:
        page = 1
    history_data = generation_repository.get_history(app_config, page=page)
    history_groups = _build_tender_groups(history_data['items'])
    return render_template(
        'history.html',
        **build_context(
            'history',
            'История генераций КП',
            history_groups=history_groups,
            pagination=history_data['pagination']
        )
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
            'timestamp': datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'),
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
    orders = datasets.get_orders_documents()
    return render_template(
        'orders.html',
        **build_context('orders', 'Распоряжения', orders=orders)
    )


@main_bp.route('/templates-library')
def templates_page():
    """Выводит список шаблонов документов."""
    templates_list = datasets.get_task_templates()
    return render_template(
        'templates_page.html',
        **build_context('templates', 'Шаблоны', templates=templates_list)
    )


@main_bp.route('/instructions')
def instructions_page():
    """Содержит краткие инструкции по бизнес-процессам."""
    instructions_list = datasets.get_task_instructions()
    return render_template(
        'instructions.html',
        **build_context('instructions', 'Инструкции', instructions=instructions_list)
    )


@main_bp.route('/analytics', methods=['GET', 'POST'])
def analytics_page():
    """Отображает раздел аналитики и обрабатывает загрузку файлов."""
    analysis_result = None
    error_message = None

    if request.method == 'POST':
        uploaded_file = request.files.get('analytics_file')
        if not uploaded_file or not uploaded_file.filename:
            error_message = 'Выберите файл Excel для анализа.'
        else:
            try:
                analysis_result = analyze_excel(uploaded_file)
            except AnalyticsProcessingError as exc:
                error_message = str(exc)
            except Exception as exc:  # pragma: no cover - логирование неожиданных ошибок
                current_app.logger.exception('Ошибка обработки аналитики: ')  # noqa: TRY401
                error_message = 'Не удалось обработать файл. Попробуйте позже.'

    return render_template(
        'analytics.html',
        **build_context('analytics', 'Аналитика', analysis=analysis_result, error_message=error_message)
    )


@main_bp.route('/duty')
def duty():
    """Предоставляет поиск по ставкам пошлин и категориям товаров."""
    query = request.args.get('q', '').strip()
    normalized_query = query.lower()
    catalog = datasets.get_duty_catalog()

    if normalized_query:
        filtered_items = [
            item for item in catalog
            if normalized_query in item.get('search_blob', '')
        ]
    else:
        filtered_items = catalog

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

    if not form_data.get('cost_price', '').strip():
        raw_per_kg = (form_data.get('cost_price_per_kg') or '').replace(',', '.').strip()
        raw_weight = (form_data.get('weight') or '').replace(',', '.').strip()
        try:
            per_kg_value = float(raw_per_kg) if raw_per_kg else None
            weight_value = float(raw_weight) if raw_weight else None
        except ValueError:
            per_kg_value = None
            weight_value = None

        if per_kg_value is not None and weight_value is not None and weight_value > 0:
            form_data['cost_price'] = str(per_kg_value * weight_value)

    validation = validate_form_data(form_data)
    form_data = validation.cleaned_data

    if validation.errors:
        for error in validation.errors:
            flash(error, 'danger')

        if validation.invalid_fields:
            form_data['_invalid_fields'] = validation.invalid_fields

        try:
            cities = datasets.load_logistics_cities()
        except Exception as exc:  # pragma: no cover - defensive logging
            current_app.logger.error('Error loading logistics data for validation errors: %s', exc)
            cities = []

        return render_template(
            'index.html',
            **build_context(
                'index',
                'Создание коммерческого предложения',
                form_data=form_data,
                cities=cities
            )
        )

    try:
        company = form_data['company'].strip()
        logistics_rub = float(form_data['logistics'])
        margin_percent = float(form_data['margin_percent'])
        delivery_time = int(form_data['delivery_time'])
        
        positions = validation.positions or extract_positions_from_form(form_data)
        form_data.pop('_invalid_fields', None)
        
        app_config = current_app.config['APP_SETTINGS']
        
        # Создаем калькулятор для множественных позиций
        calculator = MultiPositionCalculator(app_config)
        
        # Рассчитываем цены с единой итоговой маржой
        if len(positions) == 1:
            # Для одной позиции используем старый метод
            result = calculator.calculate_legacy_single_position(
                positions[0], logistics_rub, delivery_time, margin_percent
            )
            position_prices = [result]
            total_general_price = result['general_price']
        else:
            # Для множественных позиций используем новый метод с единой маржой
            calculation_result = calculator.calculate_multi_position_prices(
                positions, logistics_rub, delivery_time, margin_percent
            )
            position_prices = calculation_result['positions']
            total_general_price = calculation_result['total_revenue']
        
        # Проверяем, что есть хотя бы одна позиция
        if not position_prices:
            flash('Не удалось рассчитать цены для позиций.', 'danger')
            try:
                cities = datasets.load_logistics_cities()
            except Exception as exc:
                current_app.logger.error('Error loading logistics data for error page: %s', exc)
                cities = []
            return render_template(
                'index.html',
                **build_context('index', 'Создание коммерческого предложения', form_data=form_data, cities=cities)
            )
        
        # Для совместимости с существующим кодом используем первую позицию
        first_position = position_prices[0]
        final_price = first_position['final_price']
        general_price = first_position['general_price']
        final_price_nds = total_general_price * 1.2

        user_id = int(current_user.id) if current_user.is_authenticated else None
        if not generation_repository.save_history(form_data, final_price, app_config, user_id):
            current_app.logger.warning('Failed to save generation history')

        template_errors = check_templates_exist()
        if template_errors:
            for error in template_errors:
                flash(f'{error}. Обратитесь к администратору.', 'danger')
                current_app.logger.error(error)
            try:
                cities = datasets.load_logistics_cities()
            except Exception as exc:
                current_app.logger.error('Error loading logistics data for template error page: %s', exc)
                cities = []
            return render_template(
                'index.html',
                **build_context('index', 'Создание коммерческого предложения', form_data=form_data, cities=cities)
            )

        # Формируем ФИО менеджера для заполнения в шаблоне
        manager_fio = None
        contact_info = None
        if current_user.is_authenticated:
            last_name = current_user.last_name or ''
            first_name = current_user.first_name or ''
            if last_name or first_name:
                manager_fio = f"{last_name} {first_name}".strip()
            # Получаем контактную информацию из профиля пользователя
            contact_info = current_user.contact_info or None

        excel_template_path = 'templates_docs/template.xlsx'
        word_template_path = 'templates_docs/template.docx'

        excel_file = generate_excel_document(
            excel_template_path,
            form_data,
            final_price,
            total_general_price,
            position_prices=position_prices,
            manager_fio=manager_fio,
        )
        word_file = generate_word_document(
            word_template_path,
            form_data,
            final_price,
            total_general_price,
            final_price_nds,
            positions=positions,
            position_prices=position_prices,
            contact_info=contact_info,
        )
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
        try:
            cities = datasets.load_logistics_cities()
        except Exception as exc2:
            current_app.logger.error('Error loading logistics data for exception page: %s', exc2)
            cities = []
        return render_template(
            'index.html',
            **build_context('index', 'Создание коммерческого предложения', form_data=form_data, cities=cities)
        )


@main_bp.route('/import-positions', methods=['POST'])
def import_positions():
    """Импортирует позиции из Excel шаблона и возвращает их в формате JSON."""

    uploaded_file = request.files.get('positions_file')
    if not uploaded_file or not uploaded_file.filename:
        return jsonify({'error': 'Выберите файл шаблона в формате Excel.'}), 400

    filename = uploaded_file.filename.lower()
    if not filename.endswith(('.xlsx', '.xlsm')):
        return jsonify({'error': 'Поддерживается только формат .xlsx.'}), 400

    try:
        positions = parse_positions_from_excel(uploaded_file.stream)
    except ExcelImportError as exc:
        return jsonify({'error': str(exc)}), 400
    except Exception as exc:  # pragma: no cover - defensive logging
        current_app.logger.error('Ошибка импорта Excel: %s', exc)
        return jsonify({'error': 'Не удалось импортировать файл. Попробуйте позже.'}), 500

    session['imported_positions'] = positions
    return jsonify({'positions': positions})


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


@main_bp.route('/api/logistics/calculate', methods=['POST'])
@csrf.exempt
def calculate_logistics_api():
    """API endpoint для расчета логистики с учетом габаритов."""
    from app.services.logistics_calculator import calculate_logistics
    
    try:
        data = request.get_json()
        if data is None:
            current_app.logger.error('No JSON data received in logistics calculation request')
            return jsonify({'error': 'Не получены данные запроса'}), 400
        
        current_app.logger.debug('Logistics calculation request data: %s', data)
        
        # Получаем и валидируем обязательные поля
        weight_kg_raw = data.get('weight_kg')
        city_price_raw = data.get('city_price')
        
        if weight_kg_raw is None or city_price_raw is None:
            return jsonify({
                'error': 'Не указаны обязательные параметры: weight_kg и city_price'
            }), 400
        
        try:
            weight_kg = float(weight_kg_raw)
            city_price = float(city_price_raw)
        except (ValueError, TypeError) as e:
            current_app.logger.error('Invalid numeric values: weight_kg=%s, city_price=%s', weight_kg_raw, city_price_raw)
            return jsonify({
                'error': 'Некорректные числовые значения для веса или стоимости города'
            }), 400
        
        transport_type = data.get('transport_type', 'truck')
        length_mm = data.get('length_mm')
        width_mm = data.get('width_mm')
        height_mm = data.get('height_mm')
        
        if weight_kg <= 0:
            return jsonify({'error': 'Вес должен быть положительным числом'}), 400
        
        if city_price <= 0:
            return jsonify({'error': 'Стоимость города должна быть положительным числом'}), 400
        
        # Преобразуем габариты в float, если они указаны
        # Обрабатываем случаи: None, пустая строка, 0, или валидное число
        length_mm_float = None
        width_mm_float = None
        height_mm_float = None
        
        if length_mm is not None and length_mm != '':
            try:
                length_mm_float = float(length_mm)
                if length_mm_float <= 0:
                    return jsonify({'error': 'Длина должна быть положительным числом'}), 400
            except (ValueError, TypeError):
                return jsonify({'error': 'Некорректное значение длины'}), 400
        
        if width_mm is not None and width_mm != '':
            try:
                width_mm_float = float(width_mm)
                if width_mm_float <= 0:
                    return jsonify({'error': 'Ширина должна быть положительным числом'}), 400
            except (ValueError, TypeError):
                return jsonify({'error': 'Некорректное значение ширины'}), 400
        
        if height_mm is not None and height_mm != '':
            try:
                height_mm_float = float(height_mm)
                if height_mm_float <= 0:
                    return jsonify({'error': 'Высота должна быть положительным числом'}), 400
            except (ValueError, TypeError):
                return jsonify({'error': 'Некорректное значение высоты'}), 400
        
        # Если указаны не все три габарита, игнорируем их все
        if length_mm_float is None or width_mm_float is None or height_mm_float is None:
            length_mm_float = None
            width_mm_float = None
            height_mm_float = None
        
        current_app.logger.debug(
            'Calculating logistics: weight=%.2f, city_price=%.2f, transport=%s, dims=%sx%sx%s',
            weight_kg, city_price, transport_type, length_mm_float, width_mm_float, height_mm_float
        )
        
        result = calculate_logistics(
            weight_kg=weight_kg,
            city_price=city_price,
            transport_type=transport_type,
            length_mm=length_mm_float,
            width_mm=width_mm_float,
            height_mm=height_mm_float,
        )
        
        current_app.logger.debug('Logistics calculation result: %s', result)
        return jsonify(result)
    except Exception as exc:
        current_app.logger.exception('Error calculating logistics: %s', exc)
        return jsonify({'error': f'Ошибка при расчете логистики: {str(exc)}'}), 500
