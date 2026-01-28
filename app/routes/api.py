"""REST API endpoints для приложения."""

from typing import Any, Dict, List, Optional

from flask import Blueprint, jsonify, request
from flask_login import login_required

from app.core.exceptions import NotFoundError, ValidationError
from app.core.extensions import csrf
from app.services.repositories import (
    customer_contact_repository,
    generation_repository,
    user_repository,
)
from app.services.multi_position_calculator import MultiPositionCalculator
from app.services.logistics_calculator import calculate_logistics
from app.services.datasets import get_duty_catalog, get_gb_materials, load_logistics_cities
from app.services.exchange_rate_service import get_exchange_rate_info, clear_cache


api_bp = Blueprint('api', __name__, url_prefix='/api/v1')


@api_bp.route('/health', methods=['GET'])
def health() -> Dict[str, Any]:
    """
    Проверка работоспособности API.
    
    Returns:
        Словарь со статусом API
    """
    return jsonify({
        'status': 'ok',
        'version': '1.0',
        'service': 'kp_generator'
    })


@api_bp.route('/generations', methods=['GET'])
@csrf.exempt
@login_required
def list_generations() -> Dict[str, Any]:
    """
    Получить список генераций с пагинацией.
    Требует аутентификации.
    
    Query Parameters:
        page: Номер страницы (по умолчанию 1)
        per_page: Количество записей на странице (по умолчанию 25)
        date_from: Начальная дата фильтрации (формат: YYYY-MM-DD)
        date_to: Конечная дата фильтрации (формат: YYYY-MM-DD)
    
    Returns:
        Словарь со списком генераций и информацией о пагинации
    """
    from flask import current_app

    def _safe_int(value: Optional[str], default: int) -> int:
        try:
            return int(value) if value is not None else default
        except (TypeError, ValueError):
            return default

    def _safe_float(value: Optional[str]) -> Optional[float]:
        if value is None or value == '':
            return None
        try:
            return float(value)
        except (TypeError, ValueError):
            return None

    page = _safe_int(request.args.get('page'), 1)
    per_page = _safe_int(request.args.get('per_page'), 25)
    date_from = request.args.get('date_from')
    date_to = request.args.get('date_to')
    price_from = _safe_float(request.args.get('price_from'))
    price_to = _safe_float(request.args.get('price_to'))
    margin_from = _safe_float(request.args.get('margin_from'))
    margin_to = _safe_float(request.args.get('margin_to'))
    companies_raw = request.args.get('companies')
    search = request.args.get('search') or request.args.get('q')

    companies: Optional[List[str]] = None
    if companies_raw:
        companies = [c.strip() for c in companies_raw.split(',') if c.strip()]

    history = generation_repository.get_history(
        current_app.config['APP_SETTINGS'],
        page=page,
        per_page=per_page,
        date_from=date_from,
        date_to=date_to,
        price_from=price_from,
        price_to=price_to,
        margin_from=margin_from,
        margin_to=margin_to,
        companies=companies,
        search=search,
    )

    return jsonify(history)


@api_bp.route('/generations/<int:generation_id>', methods=['GET'])
@csrf.exempt
@login_required
def get_generation(generation_id: int) -> Dict[str, Any]:
    """
    Получить детальную информацию о генерации.
    
    Args:
        generation_id: ID генерации
    
    Returns:
        Словарь с детальной информацией о генерации
    
    Raises:
        404: Если генерация не найдена
    """
    details = generation_repository.get_details(generation_id)
    
    if details is None:
        raise NotFoundError(
            f'Генерация с ID {generation_id} не найдена',
            resource_type='generation',
            resource_id=str(generation_id)
        )
    
    return jsonify(details)


@api_bp.route('/calculate', methods=['POST'])
@csrf.exempt
@login_required
def calculate_prices() -> Dict[str, Any]:
    """
    Рассчитать цены для позиций.
    
    Request Body:
        positions: Список позиций для расчета
        logistics_rub: Стоимость логистики в рублях
        delivery_time: Время доставки в днях
        margin_percent: Целевая маржа в процентах
        config: Опциональная конфигурация расчета
    
    Returns:
        Словарь с результатами расчета
    """
    from flask import current_app
    
    data = request.get_json() or {}
    
    positions = data.get('positions', [])
    if not positions:
        raise ValidationError('Не указаны позиции для расчета', field='positions')
    
    logistics_rub = float(data.get('logistics_rub', 0))
    delivery_time = int(data.get('delivery_time', 0))
    margin_percent = float(data.get('margin_percent', 30))
    
    app_config = current_app.config.get('APP_SETTINGS', {})
    if 'config' in data:
        app_config = {**app_config, **data['config']}
    
    calculator = MultiPositionCalculator(app_config)
    
    result = calculator.calculate_multi_position_prices(
        positions, logistics_rub, delivery_time, margin_percent
    )
    
    return jsonify(result)


@api_bp.route('/logistics/calculate', methods=['POST'])
def calculate_logistics_api() -> Dict[str, Any]:
    """
    Рассчитать стоимость логистики.
    
    Request Body:
        weight_kg: Вес груза в кг
        city_price: Стоимость доставки до города
        transport_type: Тип транспорта ('truck' или 'trail')
        city_name: Название города
        is_main_route: Город на основном маршруте
        distance_from_ekb_km: Расстояние от Екатеринбурга
        length_mm: Длина груза в мм (опционально)
        width_mm: Ширина груза в мм (опционально)
        height_mm: Высота груза в мм (опционально)
    
    Returns:
        Словарь с результатами расчета логистики
    """
    data = request.get_json() or {}
    
    weight_kg = float(data.get('weight_kg', 0))
    city_price = float(data.get('city_price', 0))
    transport_type = data.get('transport_type', 'truck')
    city_name = data.get('city_name')
    is_main_route = data.get('is_main_route', True)
    distance_from_ekb_km = data.get('distance_from_ekb_km')
    
    length_mm = data.get('length_mm')
    width_mm = data.get('width_mm')
    height_mm = data.get('height_mm')
    
    result = calculate_logistics(
        weight_kg=weight_kg,
        city_price=city_price,
        transport_type=transport_type,
        city_name=city_name,
        is_main_route=is_main_route,
        distance_from_ekb_km=distance_from_ekb_km,
        length_mm=length_mm,
        width_mm=width_mm,
        height_mm=height_mm,
    )
    
    return jsonify(result)


@api_bp.route('/duty/search', methods=['GET'])
def search_duty() -> Dict[str, Any]:
    """
    Поиск ставок пошлин.
    
    Query Parameters:
        query: Поисковый запрос
        limit: Максимальное количество результатов (по умолчанию 50)
    
    Returns:
        Словарь со списком найденных ставок пошлин
    """
    query = request.args.get('query', '').strip()
    limit = int(request.args.get('limit', 50))
    
    catalog = get_duty_catalog()
    
    if query:
        query_lower = query.lower()
        catalog = [
            item for item in catalog
            if query_lower in item.get('search_blob', '').lower()
            or query_lower in item.get('title', '').lower()
        ]
    
    catalog = catalog[:limit]
    
    return jsonify({
        'items': catalog,
        'total': len(catalog),
        'query': query
    })


@api_bp.route('/materials/gb', methods=['GET'])
def get_gb_materials_api() -> Dict[str, Any]:
    """
    Получить список аналогов материалов GB.
    
    Query Parameters:
        search: Поисковый запрос (опционально)
    
    Returns:
        Словарь со списком материалов
    """
    search = request.args.get('search', '').strip()
    
    materials = get_gb_materials()
    
    if search:
        search_lower = search.lower()
        materials = [
            m for m in materials
            if search_lower in m.get('russian', '').lower()
            or search_lower in m.get('gb', '').lower()
            or search_lower in m.get('gost', '').lower()
        ]
    
    return jsonify({
        'items': materials,
        'total': len(materials)
    })


@api_bp.route('/logistics/cities', methods=['GET'])
def get_logistics_cities() -> Dict[str, Any]:
    """
    Получить список городов логистики.
    
    Returns:
        Словарь со списком городов
    """
    cities = load_logistics_cities()
    
    return jsonify({
        'items': cities,
        'total': len(cities)
    })


@api_bp.route('/customers', methods=['GET'])
@csrf.exempt
@login_required
def list_customers() -> Dict[str, Any]:
    """
    Получить список контактов заказчиков.
    
    Query Parameters:
        search: Поисковый запрос (опционально)
    
    Returns:
        Словарь со списком контактов
    """
    search = request.args.get('search', '').strip() or None
    
    contacts = customer_contact_repository.get_all(search=search)
    
    return jsonify({
        'items': contacts,
        'total': len(contacts)
    })


@api_bp.route('/customers', methods=['POST'])
@csrf.exempt
@login_required
def create_customer() -> Dict[str, Any]:
    """
    Создать новый контакт заказчика.
    
    Request Body:
        company_name: Название компании (обязательно)
        contact_person: Контактное лицо (опционально)
        phone: Телефон (опционально)
        email: Email (опционально)
        address: Адрес (опционально)
        notes: Заметки (опционально)
    
    Returns:
        Словарь с данными созданного контакта
    """
    data = request.get_json() or {}
    
    company_name = data.get('company_name', '').strip()
    if not company_name:
        raise ValidationError('Название компании обязательно', field='company_name')
    
    contact_id = customer_contact_repository.create(
        company_name=company_name,
        contact_person=data.get('contact_person'),
        phone=data.get('phone'),
        email=data.get('email'),
        address=data.get('address'),
        notes=data.get('notes')
    )
    
    if contact_id is None:
        raise ValidationError('Не удалось создать контакт')
    
    contact = customer_contact_repository.get_by_id(contact_id)
    return jsonify(contact), 201


@api_bp.route('/customers/<int:customer_id>', methods=['GET'])
@csrf.exempt
@login_required
def get_customer(customer_id: int) -> Dict[str, Any]:
    """
    Получить контакт заказчика по ID.
    
    Args:
        customer_id: ID контакта
    
    Returns:
        Словарь с данными контакта
    
    Raises:
        404: Если контакт не найден
    """
    contact = customer_contact_repository.get_by_id(customer_id)
    
    if contact is None:
        raise NotFoundError(
            f'Контакт с ID {customer_id} не найден',
            resource_type='customer_contact',
            resource_id=str(customer_id)
        )
    
    return jsonify(contact)


@api_bp.route('/customers/<int:customer_id>', methods=['PUT'])
@csrf.exempt
@login_required
def update_customer(customer_id: int) -> Dict[str, Any]:
    """
    Обновить контакт заказчика.
    
    Args:
        customer_id: ID контакта
    
    Request Body:
        company_name: Название компании (опционально)
        contact_person: Контактное лицо (опционально)
        phone: Телефон (опционально)
        email: Email (опционально)
        address: Адрес (опционально)
        notes: Заметки (опционально)
    
    Returns:
        Словарь с обновленными данными контакта
    
    Raises:
        404: Если контакт не найден
    """
    data = request.get_json() or {}
    
    success = customer_contact_repository.update(
        customer_id,
        company_name=data.get('company_name'),
        contact_person=data.get('contact_person'),
        phone=data.get('phone'),
        email=data.get('email'),
        address=data.get('address'),
        notes=data.get('notes')
    )
    
    if not success:
        raise NotFoundError(
            f'Контакт с ID {customer_id} не найден',
            resource_type='customer_contact',
            resource_id=str(customer_id)
        )
    
    contact = customer_contact_repository.get_by_id(customer_id)
    return jsonify(contact)


@api_bp.route('/customers/<int:customer_id>', methods=['DELETE'])
@csrf.exempt
@login_required
def delete_customer(customer_id: int) -> Dict[str, Any]:
    """
    Удалить контакт заказчика.
    
    Args:
        customer_id: ID контакта
    
    Returns:
        Словарь с подтверждением удаления
    
    Raises:
        404: Если контакт не найден
    """
    success = customer_contact_repository.delete(customer_id)
    
    if not success:
        raise NotFoundError(
            f'Контакт с ID {customer_id} не найден',
            resource_type='customer_contact',
            resource_id=str(customer_id)
        )
    
    return jsonify({'message': 'Контакт успешно удален'}), 200


@api_bp.route('/exchange-rate', methods=['GET'])
@csrf.exempt
def get_exchange_rate() -> Dict[str, Any]:
    """
    Получить актуальный курс китайского юаня (CNY) к российскому рублю (RUB).
    
    Query Parameters:
        force_refresh: Если true, очищает кэш и принудительно обновляет курс
    
    Returns:
        Словарь с информацией о курсе валют:
        {
            "rate": 14.5,
            "date": "2026-01-28",
            "source": "api",
            "cached": false
        }
    """
    from flask import current_app
    
    # Проверяем, нужно ли принудительно обновить курс
    force_refresh = request.args.get('force_refresh', 'false').lower() == 'true'
    if force_refresh:
        clear_cache()
        current_app.logger.info("Кэш курса валют очищен по запросу force_refresh")
    
    rate_info = get_exchange_rate_info()
    current_app.logger.info("Запрос курса валют: rate=%.4f, source=%s", rate_info.get('rate'), rate_info.get('source'))
    return jsonify(rate_info)

