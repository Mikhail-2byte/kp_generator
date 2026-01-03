import json
import re
from collections import OrderedDict
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, List, Optional, Tuple, Union

from flask import (
    Blueprint,
    Response,
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
from flask_login import current_user, login_required

from app.services.export_service import (
    export_generation_history_to_excel,
    export_generation_history_to_csv,
    create_excel_response,
    create_csv_response,
)
from app.services.generation_orchestrator import GenerationOrchestrator
from app.services import (
    AnalyticsProcessingError,
    analyze_excel,
    datasets
)
from app.services.analytics_enhancements import (
    generate_exchange_rate_analysis,
    generate_interactive_report,
    generate_margin_analysis,
)
from app.services.repositories import generation_repository
from app.services.feedback import save_feedback_entry
from app.services.excel_importer import ExcelImportError, parse_positions_from_excel
from app.presentation.ui import build_context
from app.core.extensions import csrf
from app.core.exceptions import CalculationError, DocumentGenerationError, ValidationError


main_bp = Blueprint('main', __name__)  # Основные страницы и бизнес-логика генератора КП
FORM_TTL_HOURS = 24  # время жизни сохранённой формы в сессии


def _load_logistics_cities_safe() -> List[Dict[str, Any]]:
    """Безопасно загружает список городов логистики с обработкой ошибок."""
    try:
        return datasets.load_logistics_cities()
    except Exception as exc:  # pragma: no cover - defensive logging
        current_app.logger.error('Error loading logistics data: %s', exc)
        return []


def _clear_form_session() -> None:
    """Очищает сохранённые данные формы и импортированные позиции в сессии."""
    session.pop('form_data', None)
    session.pop('form_saved_at', None)
    session.pop('imported_positions', None)


def _save_form_session(
    form_data: Dict[str, Any], imported_positions: Optional[List[Dict[str, Any]]] = None
) -> None:
    """Сохраняет состояние формы и (опционально) импортированные позиции в сессии с отметкой времени."""
    session['form_data'] = form_data or {}
    if imported_positions is not None:
        session['imported_positions'] = imported_positions
    session['form_saved_at'] = datetime.now(timezone.utc).isoformat()


def _load_form_session() -> Tuple[Dict[str, Any], Optional[List[Dict[str, Any]]]]:
    """Возвращает сохранённое состояние формы, если оно не устарело."""
    saved_at_raw = session.get('form_saved_at')
    form_data = session.get('form_data') or {}
    imported_positions = session.get('imported_positions')

    if not saved_at_raw:
        return form_data, imported_positions

    try:
        saved_at = datetime.fromisoformat(saved_at_raw)
    except ValueError:
        _clear_form_session()
        return {}, None

    if saved_at < datetime.now(timezone.utc) - timedelta(hours=FORM_TTL_HOURS):
        _clear_form_session()
        return {}, None

    return form_data, imported_positions


@main_bp.route('/')
def index() -> str:
    """Показывает форму расчёта коммерческого предложения и заполняет справочники."""
    # Забираем сохранённое состояние формы, но не удаляем — чтобы после навигации данные оставались
    form_data, imported_positions = _load_form_session()
    for field in ['id', 'timestamp', 'final_price', 'user_id']:
        form_data.pop(field, None)

    if imported_positions:
        form_data.setdefault('positions_payload', json.dumps(imported_positions, ensure_ascii=False))
        first_position = imported_positions[0]
        for key in ['product', 'drawing_number', 'material', 'cost_price', 'quantity', 'weight', 'duty_percent']:
            if first_position.get(key):
                form_data.setdefault(key, first_position[key])

    cities = _load_logistics_cities_safe()

    return render_template(
        'index.html',
        **build_context('index', 'Создание коммерческого предложения', form_data=form_data, cities=cities)
    )


@main_bp.route('/history/details/<int:record_id>')
def history_details(record_id: int) -> Response:
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
def history_by_drawing() -> Response:
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
def history_by_tender(tender_number: str) -> Response:
    """Возвращает все версии генераций по номеру тендера."""
    if not tender_number:
        return jsonify({'records': []})
    records = generation_repository.get_by_tender(tender_number)
    return jsonify({'records': records})


@main_bp.route('/history/companies')
def history_companies():
    """Возвращает список уникальных компаний."""
    try:
        companies = generation_repository.get_unique_companies()
        return jsonify({'companies': companies})
    except Exception as exc:
        current_app.logger.error('Error getting companies: %s', exc)
        return jsonify({'companies': []}), 500


def _build_tender_groups(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
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
def history() -> str:
    """Отображает список последних генераций пользователя."""
    app_config = current_app.config['APP_SETTINGS']
    try:
        page = int(request.args.get('page', '1'))
    except ValueError:
        page = 1
    
    # Получаем параметры фильтрации
    date_from = request.args.get('date_from')
    date_to = request.args.get('date_to')
    price_from_raw = request.args.get('price_from')
    price_to_raw = request.args.get('price_to')
    margin_from_raw = request.args.get('margin_from')
    margin_to_raw = request.args.get('margin_to')
    companies_raw = request.args.get('companies')
    search = request.args.get('search') or request.args.get('q')

    def _safe_float(value: str | None) -> Optional[float]:
        if value is None or value == '':
            return None
        try:
            return float(value)
        except (TypeError, ValueError):
            return None

    price_from = _safe_float(price_from_raw)
    price_to = _safe_float(price_to_raw)
    margin_from = _safe_float(margin_from_raw)
    margin_to = _safe_float(margin_to_raw)

    companies: Optional[list[str]] = None
    if companies_raw:
        companies = [c.strip() for c in companies_raw.split(',') if c.strip()]
    
    # Параметры сортировки
    sort_by = request.args.get('sort_by')
    sort_order = request.args.get('sort_order')
    
    # Валидация sort_order
    if sort_order and sort_order.lower() not in ('asc', 'desc'):
        sort_order = None
    
    history_data = generation_repository.get_history(
        app_config, 
        page=page,
        date_from=date_from,
        date_to=date_to,
        price_from=price_from,
        price_to=price_to,
        margin_from=margin_from,
        margin_to=margin_to,
        companies=companies,
        search=search,
        sort_by=sort_by,
        sort_order=sort_order,
    )
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


@main_bp.route('/history/export')
@login_required
def export_history() -> Response:
    """Экспортирует историю генераций в Excel или CSV с применением текущих фильтров."""
    app_config = current_app.config['APP_SETTINGS']
    export_format = request.args.get('format', 'excel').lower()
    
    if export_format not in ('excel', 'csv'):
        flash('Неподдерживаемый формат экспорта. Используйте excel или csv.', 'danger')
        return redirect(url_for('main.history'))
    
    try:
        # Получаем те же параметры фильтрации, что и на странице
        date_from = request.args.get('date_from')
        date_to = request.args.get('date_to')
        price_from_raw = request.args.get('price_from')
        price_to_raw = request.args.get('price_to')
        margin_from_raw = request.args.get('margin_from')
        margin_to_raw = request.args.get('margin_to')
        companies_raw = request.args.get('companies')
        search = request.args.get('search') or request.args.get('q')
        sort_by = request.args.get('sort_by')
        sort_order = request.args.get('sort_order')

        def _safe_float(value: str | None) -> Optional[float]:
            if value is None or value == '':
                return None
            try:
                return float(value)
            except (TypeError, ValueError):
                return None

        price_from = _safe_float(price_from_raw)
        price_to = _safe_float(price_to_raw)
        margin_from = _safe_float(margin_from_raw)
        margin_to = _safe_float(margin_to_raw)

        companies: Optional[list[str]] = None
        if companies_raw:
            companies = [c.strip() for c in companies_raw.split(',') if c.strip()]

        # Валидация sort_order
        if sort_order and sort_order.lower() not in ('asc', 'desc'):
            sort_order = None

        # Получаем ВСЕ записи без пагинации (используем большой per_page)
        history_data = generation_repository.get_history(
            app_config,
            page=1,
            per_page=10000,  # Большое значение для получения всех записей
            date_from=date_from,
            date_to=date_to,
            price_from=price_from,
            price_to=price_to,
            margin_from=margin_from,
            margin_to=margin_to,
            companies=companies,
            search=search,
            sort_by=sort_by,
            sort_order=sort_order,
        )

        # Генерируем имя файла с датой и временем
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f'generation_history_{timestamp}'

        if export_format == 'excel':
            buffer = export_generation_history_to_excel(history_data)
            return create_excel_response(buffer, f'{filename}.xlsx')
        else:  # csv
            buffer = export_generation_history_to_csv(history_data)
            return create_csv_response(buffer, f'{filename}.csv')

    except Exception as exc:
        current_app.logger.error('Error exporting history: %s', exc)
        flash('Произошла ошибка при экспорте данных.', 'danger')
        return redirect(url_for('main.history'))


@main_bp.route('/ai-agent')
def ai_agent() -> str:
    """Страница AI агента-консультанта. Доступна без авторизации."""
    return render_template(
        'ai_agent.html',
        **build_context('ai_agent', 'AI Консультант')
    )


@main_bp.route('/api/ai-agent/chat', methods=['POST'])
@csrf.exempt
def ai_agent_chat() -> Response:
    """API endpoint для общения с AI агентом. Доступен без авторизации."""
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'Не получены данные запроса', 'response': None}), 400
        
        message = data.get('message', '').strip()
        if not message:
            return jsonify({'error': 'Сообщение не может быть пустым', 'response': None}), 400
        
        # Инициализируем AI агента
        # Добавляем корень проекта в sys.path для корректного импорта
        import sys
        from pathlib import Path
        project_root = Path(__file__).resolve().parents[1]
        if str(project_root) not in sys.path:
            sys.path.insert(0, str(project_root))
        
        try:
            from ai_agent.agent import AIAgent
        except ImportError as import_err:
            current_app.logger.error('Failed to import AIAgent: %s', import_err, exc_info=True)
            return jsonify({
                'response': None,
                'error': f'Ошибка импорта AI агента: {str(import_err)}. Проверьте, что модуль ai_agent доступен.'
            }), 500
        
        try:
            current_app.logger.info('Initializing AIAgent...')
            agent = AIAgent()
            current_app.logger.info('AI agent initialized successfully')
        except Exception as init_err:
            current_app.logger.error('Failed to initialize AIAgent: %s', init_err, exc_info=True)
            import traceback
            current_app.logger.error('Traceback: %s', traceback.format_exc())
            return jsonify({
                'response': None,
                'error': f'Ошибка инициализации AI агента: {str(init_err)}. Проверьте логи сервера для деталей.'
            }), 500
        
        # Восстанавливаем историю из сессии, если есть
        if 'ai_agent_history' in session:
            try:
                history_data = session['ai_agent_history']
                # Проверяем, что это список
                if isinstance(history_data, list):
                    agent.conversation_history = history_data
                else:
                    current_app.logger.warning('Invalid history format in session, resetting')
                    agent.conversation_history = []
            except Exception as e:
                current_app.logger.warning('Failed to restore history from session: %s', e)
                # Если не удалось восстановить, начинаем с чистого листа
                agent.conversation_history = []
        
        # Отправляем сообщение агенту
        current_app.logger.info('Sending message to AI agent: %s', message[:100])
        response = agent.chat(message)
        current_app.logger.info('Received response from AI agent, length: %d', len(response) if response else 0)
        
        # Проверяем, что получили ответ
        if not response:
            current_app.logger.warning('Empty response from AI agent')
            return jsonify({
                'response': None,
                'error': 'Агент не вернул ответ. Попробуйте еще раз.'
            }), 500
        
        # Сохраняем историю в сессии (ограничиваем размер)
        try:
            # Сохраняем только последние 20 сообщений, чтобы не перегружать сессию
            history_to_save = agent.conversation_history[-20:] if len(agent.conversation_history) > 20 else agent.conversation_history
            
            # Очищаем history от объектов, которые не могут быть сериализованы (например, reasoning_details с сложными объектами)
            serializable_history = []
            for msg in history_to_save:
                clean_msg = {
                    'role': msg.get('role'),
                    'content': msg.get('content', '')
                }
                # reasoning_details может содержать сложные объекты, пропускаем его для сессии
                # или сериализуем только если это простой dict
                if 'reasoning_details' in msg:
                    try:
                        # Пытаемся сохранить только если это простой dict
                        if isinstance(msg['reasoning_details'], dict):
                            clean_msg['reasoning_details'] = msg['reasoning_details']
                    except Exception:
                        pass  # Пропускаем reasoning_details если не можем сериализовать
                serializable_history.append(clean_msg)
            
            session['ai_agent_history'] = serializable_history
        except Exception as e:
            current_app.logger.warning('Failed to save AI agent history to session: %s', e)
        
        return jsonify({
            'response': response,
            'error': None
        })
        
    except ImportError as e:
        current_app.logger.error('Import error in AI agent chat: %s', e, exc_info=True)
        return jsonify({
            'response': None,
            'error': f'Ошибка импорта модуля AI агента: {str(e)}'
        }), 500
    except Exception as e:
        current_app.logger.error('Error in AI agent chat: %s', e, exc_info=True)
        return jsonify({
            'response': None,
            'error': f'Ошибка при обработке запроса: {str(e)}'
        }), 500


@main_bp.route('/feedback', methods=['GET', 'POST'])
def feedback() -> Union[str, Response]:
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
def gb_analogs() -> str:
    """Показывает таблицу аналогов материалов по китайскому стандарту GB."""
    query = request.args.get('q', '').strip()
    search_field = request.args.get('field', 'all').strip()
    
    normalized_query = query.lower()
    filtered_materials = datasets.get_gb_materials()

    # Применяем поиск по выбранному полю
    if normalized_query:
        filtered = []
        for material in filtered_materials:
            match = False
            
            if search_field == 'all' or not search_field:
                # Поиск по всем полям
                match = (
                    normalized_query in material['russian'].lower()
                    or normalized_query in material['gb'].lower()
                    or normalized_query in material.get('gost', '').lower()
                    or normalized_query in material.get('price', '').lower()
                    or normalized_query in material.get('workpiece_type', '').lower()
                )
            elif search_field == 'russian':
                # Поиск по российскому материалу
                match = normalized_query in material['russian'].lower()
            elif search_field == 'gb':
                # Поиск по GB аналогу
                match = normalized_query in material['gb'].lower()
            elif search_field == 'gost':
                # Поиск по ГОСТ
                match = normalized_query in material.get('gost', '').lower()
            elif search_field == 'price':
                # Поиск по цене
                match = normalized_query in material.get('price', '').lower()
            elif search_field == 'workpiece_type':
                # Поиск по виду заготовки
                match = normalized_query in material.get('workpiece_type', '').lower()
            
            if match:
                filtered.append(material)
        
        filtered_materials = filtered

    return render_template(
        'gb_analogs.html',
        **build_context(
            'gb', 
            'Аналоги по стандарту GB', 
            materials=filtered_materials, 
            query=query,
            search_field=search_field
        )
    )


@main_bp.route('/orders')
def orders_page() -> str:
    """Отображает раздел с распоряжениями и внутренними документами."""
    orders = datasets.get_orders_documents()
    if not orders:
        flash(
            'Справочник распоряжений пуст или недоступен. '
            'Добавьте записи в разделе «Администрирование → Распоряжения».',
            'warning',
        )
    return render_template(
        'orders.html',
        **build_context('orders', 'Распоряжения', orders=orders)
    )


@main_bp.route('/templates-library')
def templates_page() -> str:
    """Выводит список шаблонов документов."""
    templates_list = datasets.get_task_templates()
    if not templates_list:
        flash(
            'Справочник шаблонов пуст или недоступен. '
            'Добавьте шаблоны в разделе «Администрирование → Шаблоны».',
            'warning',
        )
    return render_template(
        'templates_page.html',
        **build_context('templates', 'Шаблоны', templates=templates_list)
    )


def _parse_instruction_content(content_text: str) -> Dict[str, Any]:
    """
    Парсит текст инструкции и выделяет разделы для визуального отображения.
    
    Returns:
        dict с ключами: title, goal, steps, result, errors, checklist
    """
    if not content_text:
        return {}
    
    lines = content_text.split('\n')
    parsed = {
        'title': '',
        'subtitle': '',
        'goal': '',
        'steps': [],
        'result': '',
        'errors': '',
        'checklist': ''
    }
    
    current_section = None
    current_content = []
    current_step = None
    separator_pattern = re.compile(r'^_{10,}$')
    
    for i, line in enumerate(lines):
        line = line.strip()
        
        # Определяем разделы
        if line == 'Цель':
            if current_step:
                parsed['steps'].append(current_step)
                current_step = None
            current_section = 'goal'
            current_content = []
            continue
        elif line == 'Порядок действий':
            if current_step:
                parsed['steps'].append(current_step)
                current_step = None
            current_section = 'steps'
            current_content = []
            continue
        elif line == 'Результат':
            if current_step:
                parsed['steps'].append(current_step)
                current_step = None
            current_section = 'result'
            current_content = []
            continue
        elif line == 'Типовые ошибки':
            if current_step:
                parsed['steps'].append(current_step)
                current_step = None
            current_section = 'errors'
            current_content = []
            continue
        elif line.startswith('Краткий чек-лист') or line.startswith('Чек-лист'):
            if current_step:
                parsed['steps'].append(current_step)
                current_step = None
            current_section = 'checklist'
            current_content = []
            continue
        elif separator_pattern.match(line):
            # Разделитель - сохраняем текущий контент
            if current_section == 'goal' and current_content:
                parsed['goal'] = '\n'.join(current_content).strip()
                current_content = []
            elif current_section == 'result' and current_content:
                parsed['result'] = '\n'.join(current_content).strip()
                current_content = []
            elif current_section == 'errors' and current_content:
                parsed['errors'] = '\n'.join(current_content).strip()
                current_content = []
            elif current_section == 'checklist' and current_content:
                parsed['checklist'] = '\n'.join(current_content).strip()
                current_content = []
            elif current_section == 'steps' and current_step:
                # Сохраняем текущий шаг при разделителе
                parsed['steps'].append(current_step)
                current_step = None
            continue
        
        if not line:
            # Пустая строка - добавляем в контент, если есть текущий раздел
            if current_section and current_section != 'steps':
                current_content.append('')
            elif current_step and current_step.get('items'):
                # Пустая строка в списке пунктов шага
                pass
            continue
            
        if current_section == 'steps':
            # Обработка нумерованных шагов
            step_match = re.match(r'^(\d+)\.\s*(.+)$', line)
            if step_match:
                # Новый шаг
                if current_step:
                    parsed['steps'].append(current_step)
                current_step = {
                    'number': step_match.group(1),
                    'title': step_match.group(2).strip(),
                    'items': []
                }
            elif line.startswith('•') or line.startswith('-') or (line and line[0] in '•-'):
                # Пункт списка
                item_text = re.sub(r'^[•\-]\s*', '', line).strip()
                if current_step:
                    current_step['items'].append(item_text)
                else:
                    current_content.append(line)
            elif line.startswith('Определить:') or line.startswith('Скачать') or line.startswith('Кратко') or line.startswith('Понять:') or line.startswith('Оценить:') or line.startswith('Действовать'):
                # Описательный текст шага
                if current_step:
                    if not current_step['items']:
                        current_step['title'] += ' ' + line
                    else:
                        # Добавляем как дополнительный пункт
                        current_step['items'].append(line)
                else:
                    current_content.append(line)
            elif current_step:
                # Продолжение текста шага
                if current_step['items']:
                    # Добавляем к последнему пункту или как новый пункт
                    if current_step['items']:
                        current_step['items'][-1] += ' ' + line
                    else:
                        current_step['items'].append(line)
                else:
                    current_step['title'] += ' ' + line
            else:
                current_content.append(line)
        elif current_section:
            # Обычный контент для других разделов
            current_content.append(line)
        elif i < 3:
            # Первые строки - заголовок и подзаголовок
            if not parsed['title']:
                parsed['title'] = line
            elif not parsed['subtitle'] and line != parsed['title']:
                parsed['subtitle'] = line
    
    # Сохраняем последний шаг
    if current_step:
        parsed['steps'].append(current_step)
    
    # Сохраняем последний контент
    if current_section == 'goal' and current_content:
        parsed['goal'] = '\n'.join(current_content).strip()
    elif current_section == 'result' and current_content:
        parsed['result'] = '\n'.join(current_content).strip()
    elif current_section == 'errors' and current_content:
        parsed['errors'] = '\n'.join(current_content).strip()
    elif current_section == 'checklist' and current_content:
        parsed['checklist'] = '\n'.join(current_content).strip()
    
    return parsed


@main_bp.route('/instructions')
def instructions_page() -> str:
    """Содержит краткие инструкции по бизнес-процессам."""
    instructions_list = datasets.get_task_instructions()
    if not instructions_list:
        flash(
            'Справочник инструкций пуст или недоступен. '
            'Добавьте инструкции в разделе «Администрирование → Инструкции».',
            'warning',
        )
    
    # Парсим содержимое каждой инструкции
    for instruction in instructions_list:
        if isinstance(instruction, dict) and instruction.get('content_text'):
            instruction['parsed'] = _parse_instruction_content(instruction['content_text'])
    
    return render_template(
        'instructions.html',
        **build_context('instructions', 'Инструкции', instructions=instructions_list)
    )


@main_bp.route('/analytics', methods=['GET', 'POST'])
def analytics_page() -> str:
    """Отображает раздел аналитики и обрабатывает загрузку файлов."""
    analysis_result = None
    error_message = None
    margin_analysis = None
    exchange_analysis = None
    interactive_report = None

    # Получаем параметры для расширенной аналитики
    show_margin = request.args.get('margin', 'false').lower() == 'true'
    show_exchange = request.args.get('exchange', 'false').lower() == 'true'
    show_report = request.args.get('report', 'false').lower() == 'true'
    
    user_id = int(current_user.id) if current_user.is_authenticated else None
    days = int(request.args.get('days', 30))
    date_from = request.args.get('date_from')
    date_to = request.args.get('date_to')

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
    
    # Генерируем расширенную аналитику при запросе
    if show_margin:
        try:
            margin_analysis = generate_margin_analysis(user_id=user_id, days=days)
        except Exception as exc:
            current_app.logger.error(f'Ошибка генерации анализа маржи: {exc}')
    
    if show_exchange:
        try:
            exchange_analysis = generate_exchange_rate_analysis(days=days)
        except Exception as exc:
            current_app.logger.error(f'Ошибка генерации анализа курсов: {exc}')
    
    if show_report:
        try:
            interactive_report = generate_interactive_report(
                user_id=user_id,
                date_from=date_from,
                date_to=date_to
            )
        except Exception as exc:
            current_app.logger.error(f'Ошибка генерации отчета: {exc}')

    return render_template(
        'analytics.html',
        **build_context(
            'analytics',
            'Аналитика',
            analysis=analysis_result,
            error_message=error_message,
            margin_analysis=margin_analysis,
            exchange_analysis=exchange_analysis,
            interactive_report=interactive_report,
            show_margin=show_margin,
            show_exchange=show_exchange,
            show_report=show_report,
            days=days,
            date_from=date_from,
            date_to=date_to
        )
    )


@main_bp.route('/duty')
def duty() -> str:
    """Предоставляет поиск по ставкам пошлин и категориям товаров."""
    query = request.args.get('q', '').strip()
    normalized_query = query.lower()
    catalog = datasets.get_duty_catalog()

    if not catalog:
        flash(
            'Справочник ставок пошлин пуст или недоступен. '
            'Обратитесь к администратору для настройки раздела «Ставки пошлин».',
            'warning',
        )

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
def load_generation(gen_id: int) -> Response:
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
def generate() -> Union[str, Response]:
    """Выполняет расчёт КП, сохраняет историю и формирует пакет документов."""
    form_data = request.form.to_dict()
    form_data['comment'] = form_data.get('comment', '').strip()

    app_config = current_app.config['APP_SETTINGS']
    orchestrator = GenerationOrchestrator(app_config)

    try:
        # Подготовка данных пользователя
        user_id = int(current_user.id) if current_user.is_authenticated else None
        manager_fio = None
        contact_info = None
        if current_user.is_authenticated:
            last_name = current_user.last_name or ''
            first_name = current_user.first_name or ''
            if last_name or first_name:
                manager_fio = f"{last_name} {first_name}".strip()
            contact_info = current_user.contact_info or None

        # Используем оркестратор для генерации
        zip_buffer, file_prefix = orchestrator.orchestrate(
            form_data, user_id, manager_fio, contact_info
        )

        # Успешная генерация — очищаем сохранённую форму/импорт
        _clear_form_session()

        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name=f'{file_prefix}.zip',
            mimetype='application/zip'
        )

    except ValidationError as exc:
        # Обработка ошибок валидации с сохранением формы
        for error in exc.details.get('errors', [exc.message]):
            flash(error if isinstance(error, str) else exc.message, 'danger')
        
        cleaned_data = exc.details.get('invalid_fields', {})
        _save_form_session(cleaned_data, session.get('imported_positions'))
        if exc.details.get('invalid_fields'):
            cleaned_data['_invalid_fields'] = exc.details['invalid_fields']
        
        cities = _load_logistics_cities_safe()
        return render_template(
            'index.html',
            **build_context(
                'index',
                'Создание коммерческого предложения',
                form_data=cleaned_data,
                cities=cities
            )
        )

    except (CalculationError, DocumentGenerationError) as exc:
        flash(f'Ошибка: {exc.message}', 'danger')
        cities = _load_logistics_cities_safe()
        _save_form_session(form_data, session.get('imported_positions'))
        return render_template(
            'index.html',
            **build_context('index', 'Создание коммерческого предложения', form_data=form_data, cities=cities)
        )

    except Exception as exc:  # pragma: no cover - defensive logging
        flash('Произошла непредвиденная ошибка. Попробуйте еще раз.', 'danger')
        current_app.logger.error('Unexpected error: %s', exc, exc_info=True)
        cities = _load_logistics_cities_safe()
        _save_form_session(form_data, session.get('imported_positions'))
        return render_template(
            'index.html',
            **build_context('index', 'Создание коммерческого предложения', form_data=form_data, cities=cities)
        )


@main_bp.route('/import-positions', methods=['POST'])
def import_positions() -> Response:
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
    # Сохраняем текущую форму, если она была отправлена вместе с импортом
    if request.form:
        _save_form_session({**request.form.to_dict()}, positions)
    return jsonify({'positions': positions})


@main_bp.route('/form/reset', methods=['POST'])
@csrf.exempt
def reset_form_state() -> Response:
    """Очищает сохранённый черновик формы и импортированные позиции в сессии."""
    _clear_form_session()
    return jsonify({'status': 'ok'})


@main_bp.route('/form/save', methods=['POST'])
@csrf.exempt
def save_form_state() -> Response:
    """
    Сохраняет черновик формы из JSON-пейлоада.
    Ожидает тело вида:
    {
        "form_data": {...},
        "positions": [...]
    }
    """
    try:
        payload = request.get_json(silent=True) or {}
        form_data = payload.get('form_data') or {}
        positions = payload.get('positions')
        if positions is not None:
            try:
                # Сериализуем positions в строку JSON, если нужно сохранить в форму для совместимости
                form_data['positions_payload'] = json.dumps(positions, ensure_ascii=False)
            except (TypeError, ValueError):
                pass
        _save_form_session(form_data, positions)
        return jsonify({'status': 'ok'})
    except Exception as exc:  # pragma: no cover - defensive
        current_app.logger.error('Failed to save draft: %s', exc)
        return jsonify({'status': 'error', 'message': 'Не удалось сохранить черновик'}), 500


@main_bp.route('/logistics')
def logistics() -> str:
    """Отображает справочную информацию по логистическим тарифам."""
    cities = _load_logistics_cities_safe()
    if not cities:
        current_app.logger.info('Logistics data is empty or missing')
        flash(
            'Справочник логистики пуст или недоступен. '
            'Проверьте настройки справочников логистики в административном разделе.',
            'warning',
        )

    return render_template(
        'logistics.html',
        **build_context('logistics', 'Просчёт логистики', cities=cities)
    )


@main_bp.route('/api/logistics/calculate', methods=['POST'])
@csrf.exempt
def calculate_logistics_api() -> Response:
    """API endpoint для расчета логистики с учетом габаритов и новой логики."""
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
        
        # Новые параметры
        transport_type = data.get('transport_type', 'truck')
        city_name = data.get('city_name')
        is_main_route = data.get('is_main_route', True)
        distance_from_ekb_km_raw = data.get('distance_from_ekb_km')
        length_mm = data.get('length_mm')
        width_mm = data.get('width_mm')
        height_mm = data.get('height_mm')
        
        if weight_kg <= 0:
            return jsonify({'error': 'Вес должен быть положительным числом'}), 400
        
        # Автоматическое определение типа города из справочников
        city_in_ekb_rf_catalog = False
        if city_name:
            from app.services.datasets import is_city_in_ekb_rf_catalog, get_ekb_rf_city_distance
            city_in_ekb_rf_catalog = is_city_in_ekb_rf_catalog(city_name)
        
        # Для городов ЕКБ+РФ city_price может быть 0 (расчет по алгоритму)
        # Для остальных городов проверяем, что цена положительная
        if not city_in_ekb_rf_catalog and city_price <= 0:
            return jsonify({'error': 'Стоимость города должна быть положительным числом'}), 400
            
            # Если город в справочнике ЕКБ+РФ, автоматически устанавливаем is_main_route=False
            if city_in_ekb_rf_catalog:
                is_main_route = False
                # Получаем расстояние из справочника, если не указано в запросе
                if distance_from_ekb_km_raw is None or distance_from_ekb_km_raw == '':
                    catalog_distance = get_ekb_rf_city_distance(city_name)
                    if catalog_distance is not None:
                        distance_from_ekb_km_raw = catalog_distance
        
        # Обрабатываем расстояние от ЕКБ
        distance_from_ekb_km = None
        if distance_from_ekb_km_raw is not None and distance_from_ekb_km_raw != '':
            try:
                distance_from_ekb_km = int(distance_from_ekb_km_raw)
                if distance_from_ekb_km < 0:
                    return jsonify({'error': 'Расстояние не может быть отрицательным'}), 400
            except (ValueError, TypeError):
                return jsonify({'error': 'Некорректное значение расстояния'}), 400
        
        # Преобразуем габариты в float, если они указаны
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
            'Calculating logistics: weight=%.2f, city_price=%.2f, transport=%s, city=%s, is_main_route=%s, distance_ekb=%s, dims=%sx%sx%s',
            weight_kg, city_price, transport_type, city_name, is_main_route, distance_from_ekb_km, 
            length_mm_float, width_mm_float, height_mm_float
        )
        
        result = calculate_logistics(
            weight_kg=weight_kg,
            city_price=city_price,
            transport_type=transport_type,
            city_name=city_name,
            is_main_route=is_main_route,
            distance_from_ekb_km=distance_from_ekb_km,
            length_mm=length_mm_float,
            width_mm=width_mm_float,
            height_mm=height_mm_float,
            city_in_ekb_rf_catalog=city_in_ekb_rf_catalog if city_name else None,
        )
        
        current_app.logger.debug('Logistics calculation result: %s', result)
        return jsonify(result)
    except Exception as exc:
        current_app.logger.exception('Error calculating logistics: %s', exc)
        return jsonify({'error': f'Ошибка при расчете логистики: {str(exc)}'}), 500


@main_bp.route('/api/logistics/save-distance', methods=['POST'])
@csrf.exempt
def save_distance_from_ekb() -> Response:
    """
    Сохраняет введенное пользователем расстояние от ЕКБ до города.
    Обновляет logistics_cities.json.
    """
    from app.services.datasets import update_city_distance
    
    try:
        data = request.get_json()
        if data is None:
            return jsonify({'error': 'Не получены данные запроса'}), 400
        
        city_name = data.get('city_name')
        # Поддерживаем оба варианта названия параметра для обратной совместимости
        distance_km_raw = data.get('distance_from_ekb_km') or data.get('distance_km')
        
        if not city_name:
            return jsonify({'error': 'Не указано название города'}), 400
        
        if distance_km_raw is None:
            return jsonify({'error': 'Не указано расстояние'}), 400
        
        try:
            distance_km = int(distance_km_raw)
            if distance_km < 0:
                return jsonify({'error': 'Расстояние не может быть отрицательным'}), 400
        except (ValueError, TypeError):
            return jsonify({'error': 'Некорректное значение расстояния'}), 400
        
        # Получаем информацию о пользователе для логирования
        actor = current_user.username if current_user.is_authenticated else 'anonymous'
        
        # Обновляем расстояние
        success = update_city_distance(city_name, distance_km, actor=actor)
        
        if success:
            current_app.logger.info(
                'Distance updated for city %s: %d km (by %s)', 
                city_name, distance_km, actor
            )
            return jsonify({
                'success': True,
                'message': f'Расстояние для города {city_name} успешно сохранено: {distance_km} км'
            })
        else:
            return jsonify({
                'error': f'Город {city_name} не найден в базе данных'
            }), 404
    
    except Exception as exc:
        current_app.logger.exception('Error saving distance: %s', exc)
        return jsonify({'error': f'Ошибка при сохранении расстояния: {str(exc)}'}), 500
