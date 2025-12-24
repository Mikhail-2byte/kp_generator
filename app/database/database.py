# Database operations
import logging
import math
import sqlite3
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional, Sequence

from alembic import command
from alembic.config import Config
from alembic.runtime.migration import MigrationContext
from alembic.script import ScriptDirectory
from sqlalchemy import create_engine, func, String, cast, or_, and_
from sqlalchemy.exc import IntegrityError, SQLAlchemyError
from sqlalchemy.orm import joinedload

from app.core.exceptions import DatabaseError, NotFoundError
from app.core.extensions import SessionLocal, get_database_url
from app.models.models import GenerationHistoryRecord, UserRecord


PROJECT_ROOT = Path(__file__).resolve().parents[2]


def _resolve_sqlite_path(database_url: str) -> str:
    """Определяет фактический путь к файлу SQLite по строке подключения."""
    if database_url == 'sqlite:///:memory:':
        return ':memory:'

    if database_url.startswith('sqlite:////'):
        return database_url.replace('sqlite:////', '/', 1)

    if database_url.startswith('sqlite:///'):
        relative_path = database_url.replace('sqlite:///', '', 1)
        resolved = PROJECT_ROOT / relative_path
        return str(resolved)

    if database_url.startswith('sqlite://'):
        return database_url.replace('sqlite://', '', 1)

    raise ValueError(f'Unsupported database URL: {database_url}')


def _alembic_ini_path() -> Path:
    return PROJECT_ROOT / 'alembic.ini'


def get_alembic_config(strict: bool = False) -> Optional[Config]:
    """Возвращает конфигурацию Alembic с подставленным URL базы данных."""
    alembic_cfg = _alembic_ini_path()
    if not alembic_cfg.exists():
        message = f'Alembic configuration not found at {alembic_cfg}'
        if strict:
            raise FileNotFoundError(message)
        logging.warning(message)
        return None

    config = Config(str(alembic_cfg))
    config.set_main_option('sqlalchemy.url', get_database_url())
    return config


def get_migration_status(config: Optional[Config] = None) -> Dict[str, object]:
    """Возвращает сведения о состоянии миграций Alembic."""
    config = config or get_alembic_config()
    database_url = get_database_url()

    if config is None:
        return {
            'configured': False,
            'database_url': database_url,
            'current_revision': None,
            'head_revision': None,
            'head_revisions': [],
            'pending': False,
            'has_database': False,
        }

    script = ScriptDirectory.from_config(config)
    head_revisions = script.get_heads()
    head_revision = script.get_current_head() if head_revisions else None

    current_revision = None
    has_database = False

    engine = create_engine(database_url, future=True)
    try:
        with engine.connect() as connection:
            context = MigrationContext.configure(connection)
            current_revision = context.get_current_revision()
            has_database = True
    finally:
        engine.dispose()

    pending = bool(head_revisions) and current_revision not in head_revisions

    return {
        'configured': True,
        'database_url': database_url,
        'current_revision': current_revision,
        'head_revision': head_revision,
        'head_revisions': head_revisions,
        'pending': pending,
        'has_database': has_database,
    }


def apply_migrations(target_revision: str = 'head', *, raise_on_error: bool = True) -> bool:
    """Применяет миграции Alembic до указанной ревизии."""
    config = get_alembic_config()
    if config is None:
        message = 'Alembic configuration is missing; cannot apply migrations.'
        if raise_on_error:
            raise RuntimeError(message)
        logging.error(message)
        return False

    try:
        command.upgrade(config, target_revision)
        return True
    except Exception as exc:  # pragma: no cover - проксируем исключение вызывающему коду
        logging.exception('Failed to apply alembic migrations: %s', exc)
        if raise_on_error:
            raise
        return False


def downgrade_migrations(target_revision: str = '-1', *, raise_on_error: bool = True) -> bool:
    """Откатывает миграции Alembic до указанной ревизии."""
    config = get_alembic_config()
    if config is None:
        message = 'Alembic configuration is missing; cannot downgrade migrations.'
        if raise_on_error:
            raise RuntimeError(message)
        logging.error(message)
        return False

    try:
        command.downgrade(config, target_revision)
        return True
    except Exception as exc:  # pragma: no cover - проксируем исключение вызывающему коду
        logging.exception('Failed to downgrade alembic migrations: %s', exc)
        if raise_on_error:
            raise
        return False


def connect_db():
    """Устанавливает соединение с базой данных через sqlite3 (наследие)."""
    database_url = get_database_url()
    path = _resolve_sqlite_path(database_url)
    return sqlite3.connect(path, detect_types=sqlite3.PARSE_DECLTYPES)


def init_db():
    """Запускает миграции Alembic и проверяет актуальность схемы."""
    config = get_alembic_config()
    if config is None:
        return

    status = get_migration_status(config=config)
    if status.get('pending'):
        logging.info(
            'Pending migrations detected (current=%s, head=%s). Applying...',
            status.get('current_revision') or 'base',
            status.get('head_revision') or '—',
        )

    try:
        command.upgrade(config, 'head')
    except Exception as exc:
        logging.exception('Failed to apply database migrations: %s', exc)
        raise

    final_status = get_migration_status(config=config)
    if final_status.get('pending'):
        raise RuntimeError('Database schema is out of date after migrations execution.')


@contextmanager
def _session_scope():
    """Гарантирует корректное создание, коммит и закрытие ORM-сессии."""
    session = SessionLocal()
    try:
        yield session
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        SessionLocal.remove()


def _format_timestamp(value: Optional[datetime]) -> Optional[str]:
    """Конвертирует отметку времени в строку для отображения."""
    if value is None:
        return None
    if isinstance(value, str):
        return value
    return value.strftime('%Y-%m-%d %H:%M:%S')


def _ensure_datetime_utc(value: Optional[datetime]) -> Optional[datetime]:
    """Гарантирует, что дата содержит информацию о часовом поясе (UTC)."""
    if value is None:
        return None
    if isinstance(value, str):
        try:
            parsed = datetime.fromisoformat(value)
        except ValueError:
            return value
        value = parsed
    if isinstance(value, datetime) and value.tzinfo is None:
        return value.replace(tzinfo=timezone.utc)
    return value


def _format_history_timestamp(value: Optional[datetime]) -> str:
    """Преобразует дату истории в пользовательский формат dd.mm.yyyy HH:MM."""
    if value is None:
        return ''
    if isinstance(value, str):
        try:
            parsed = datetime.strptime(value, '%Y-%m-%d %H:%M:%S')
            return parsed.strftime('%d.%m.%Y %H:%M')
        except ValueError:
            return value
    return value.strftime('%d.%m.%Y %H:%M')


def _build_generation_detail(record: Optional[GenerationHistoryRecord]):
    """Формирует подробный словарь по записи генерации."""
    if record is None:
        return None

    import json
    
    # Базовые данные
    data = {
        'id': record.id,
        'timestamp': _format_timestamp(record.timestamp),
        'tender_number': record.tender_number,
        'company': record.company,
        'logistics': record.logistics,
        'margin_percent': record.margin_percent,
        'delivery_address': record.delivery_address,
        'delivery_time': record.delivery_time,
        'payment_terms': record.payment_terms or '',
        'proposal_validity': record.proposal_validity or '',
        'warranty_period': record.warranty_period or '',
        'comment': record.comment,
        'user_id': record.user_id,
        'username': record.user.username if record.user else None,
        'last_name': record.user.last_name if record.user else None,
        'first_name': record.user.first_name if record.user else None,
    }
    
    # Если есть данные множественных позиций, добавляем их
    if record.positions_data and record.positions_count and record.positions_count > 1:
        try:
            positions = json.loads(record.positions_data)
            data['positions'] = positions
            data['positions_count'] = record.positions_count
            data['total_general_price'] = record.total_general_price
            
            # Для совместимости используем данные первой позиции
            if positions:
                first_position = positions[0]
                data.update({
                    'product': first_position.get('product', record.product),
                    'quantity': first_position.get('quantity', record.quantity),
                    'cost_price': first_position.get('cost_price', record.cost_price),
                    'weight': first_position.get('weight', record.weight),
                    'drawing_number': first_position.get('drawing_number', record.drawing_number),
                    'material': first_position.get('material', record.material),
                    'duty_percent': first_position.get('duty_percent', record.duty_percent),
                    'final_price': record.final_price,
                })
        except (json.JSONDecodeError, TypeError) as exc:
            # Если не удалось распарсить JSON, используем старые поля
            logging.warning('Failed to parse positions_data JSON for record %d: %s', record.id, exc)
            data.update({
                'product': record.product,
                'quantity': record.quantity,
                'cost_price': record.cost_price,
                'weight': record.weight,
                'drawing_number': record.drawing_number,
                'material': record.material,
                'duty_percent': record.duty_percent,
                'final_price': record.final_price,
            })
    else:
        # Старые данные (одна позиция) - создаем массив с одной позицией для совместимости
        data['positions'] = [{
            'product': record.product,
            'quantity': record.quantity,
            'cost_price': record.cost_price,
            'weight': record.weight,
            'drawing_number': record.drawing_number,
            'material': record.material,
            'duty_percent': record.duty_percent,
            'final_price': record.final_price,
        }]
        data['positions_count'] = 1
        data['total_general_price'] = record.final_price * record.quantity
    
    return data


def _user_to_tuple(user: Optional[UserRecord]):
    """Поддерживает совместимость с устаревшим интерфейсом кортежей пользователя."""
    if user is None:
        return None
    created_at = _ensure_datetime_utc(user.created_at)
    last_login = _ensure_datetime_utc(user.last_login)
    return (
        user.id,
        user.username,
        user.password_hash,
        created_at,
        last_login,
        user.last_name,
        user.first_name,
        user.contact_info,
        user.role,
    )


def _normalize_pagination(config: Dict[str, object], page: int, per_page: Optional[int]) -> tuple[int, int]:
    default_page_size = int(config.get('history_page_size') or config.get('max_history_items', 50))
    page = max(int(page or 1), 1)
    page_size = per_page or default_page_size
    page_size = max(1, min(int(page_size), 200))
    return page, page_size


def get_generation_history(
    config,
    *,
    page: int = 1,
    per_page: Optional[int] = None,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
    price_from: Optional[float] = None,
    price_to: Optional[float] = None,
    margin_from: Optional[float] = None,
    margin_to: Optional[float] = None,
    companies: Optional[List[str]] = None,
    search: Optional[str] = None,
    sort_by: Optional[str] = None,
    sort_order: Optional[str] = None,
) -> Dict[str, object]:
    """Возвращает историю генераций с пагинацией по тендерам (последняя версия).

    Поддерживает серверную фильтрацию по датам, цене продажи, марже,
    компаниям и текстовому поиску.
    """
    page, limit = _normalize_pagination(config, page, per_page)
    offset = (page - 1) * limit
    try:
        from datetime import datetime
        with _session_scope() as session:
            tender_key_expr = func.coalesce(
                func.lower(GenerationHistoryRecord.tender_number),
                cast(GenerationHistoryRecord.id, String)
            )

            # Базовый запрос для фильтрации
            base_filter = session.query(GenerationHistoryRecord)

            # Применяем фильтры по датам
            if date_from:
                try:
                    date_from_obj = datetime.strptime(date_from, '%Y-%m-%d')
                    base_filter = base_filter.filter(GenerationHistoryRecord.timestamp >= date_from_obj)
                except ValueError:
                    pass  # Игнорируем неверный формат даты

            if date_to:
                try:
                    date_to_obj = datetime.strptime(date_to, '%Y-%m-%d')
                    # Добавляем время до конца дня
                    date_to_obj = date_to_obj.replace(hour=23, minute=59, second=59)
                    base_filter = base_filter.filter(GenerationHistoryRecord.timestamp <= date_to_obj)
                except ValueError:
                    pass  # Игнорируем неверный формат даты

            # Фильтр по цене продажи (используем общую цену, если есть)
            sale_price_expr = func.coalesce(
                GenerationHistoryRecord.total_general_price,
                GenerationHistoryRecord.final_price * func.coalesce(GenerationHistoryRecord.quantity, 1),
            )

            if price_from is not None:
                base_filter = base_filter.filter(sale_price_expr >= float(price_from))

            if price_to is not None:
                base_filter = base_filter.filter(sale_price_expr <= float(price_to))

            # Фильтр по марже
            if margin_from is not None:
                base_filter = base_filter.filter(GenerationHistoryRecord.margin_percent >= float(margin_from))

            if margin_to is not None:
                base_filter = base_filter.filter(GenerationHistoryRecord.margin_percent <= float(margin_to))

            # Фильтр по компаниям
            if companies:
                # Нормализуем список компаний, отбрасывая пустые значения
                normalized_companies = [c.strip() for c in companies if c and c.strip()]
                if normalized_companies:
                    base_filter = base_filter.filter(GenerationHistoryRecord.company.in_(normalized_companies))

            # Текстовый поиск по нескольким полям
            if search:
                search_pattern = f"%{search.lower()}%"
                search_conditions = [
                    func.lower(GenerationHistoryRecord.tender_number).like(search_pattern),
                    func.lower(GenerationHistoryRecord.product).like(search_pattern),
                    func.lower(GenerationHistoryRecord.drawing_number).like(search_pattern),
                    func.lower(GenerationHistoryRecord.company).like(search_pattern),
                ]
                
                # Поиск по наименованиям позиций в JSON (positions_data)
                # Ищем в поле "product" внутри массива позиций
                # JSON формат: [{"product":"Ротор",...}, {"product":"Колесо",...}]
                # json.dumps() создает компактный формат без пробелов: "product":"значение"
                # Важно: не используем func.lower() на всем JSON, так как это может нарушить структуру
                # Вместо этого ищем с учетом различных вариантов регистра в самом паттерне
                
                search_lower = search.lower()
                
                # Используем более надежный подход: ищем слово в JSON с учетом регистра
                # но используем паттерны, которые покрывают основные варианты регистра
                # JSON формат компактный: [{"product":"Ротор",...}]
                
                # Генерируем варианты поискового запроса с разным регистром
                # для покрытия случаев, когда в JSON сохранено с разным регистром
                search_variants = [
                    search_lower,           # "ротор"
                    search.upper(),         # "РОТОР"
                    search.capitalize(),    # "Ротор"
                    search.title(),         # "Ротор" (если несколько слов)
                ]
                # Убираем дубликаты
                search_variants = list(dict.fromkeys(search_variants))
                
                # Создаем паттерны для каждого варианта регистра
                json_patterns = []
                for variant in search_variants:
                    # Компактный формат JSON (без пробелов после :) - основной формат json.dumps()
                    json_patterns.extend([
                        f'%"product":"%{variant}%"%',    # "product":"...ротор..." (слово внутри строки)
                        f'%"product":"{variant}"%',      # "product":"ротор" (точное совпадение)
                        f'%"product":"{variant}",%',     # "product":"ротор", (в конце объекта)
                        f'%,"product":"{variant}"%',     # ,"product":"ротор" (не первая позиция)
                    ])
                    
                    # Формат с пробелом после : (на случай если формат изменится в будущем)
                    json_patterns.extend([
                        f'%"product": "%{variant}%"%',   # "product": "...ротор..."
                        f'%"product": "{variant}"%',     # "product": "ротор"
                    ])
                
                # Добавляем условие поиска в positions_data для каждого паттерна
                # Используем AND с проверкой на NULL, чтобы избежать ошибок
                # Используем OR для всех паттернов, чтобы найти любое совпадение
                for json_pattern in json_patterns:
                    search_conditions.append(
                        and_(
                            GenerationHistoryRecord.positions_data.isnot(None),
                            GenerationHistoryRecord.positions_data.like(json_pattern)
                        )
                    )
                
                base_filter = base_filter.filter(or_(*search_conditions))

            # Создаем подзапрос с фильтрацией для window функции
            filtered_subquery = base_filter.with_entities(GenerationHistoryRecord.id).subquery()

            # Общее количество уникальных тендеров с учетом фильтров
            # Считаем только среди отфильтрованных записей
            total = session.query(
                func.count(func.distinct(
                    func.coalesce(
                        func.lower(GenerationHistoryRecord.tender_number),
                        cast(GenerationHistoryRecord.id, String)
                    )
                ))
            ).filter(GenerationHistoryRecord.id.in_(session.query(filtered_subquery.c.id))).scalar() or 0

            window_subquery = (
                session.query(
                    GenerationHistoryRecord.id.label('id'),
                    tender_key_expr.label('tender_key'),
                    func.row_number()
                    .over(
                        partition_by=tender_key_expr,
                        order_by=(
                            GenerationHistoryRecord.timestamp.desc(),
                            GenerationHistoryRecord.id.desc(),
                        ),
                    )
                    .label('row_number'),
                    func.count()
                    .over(partition_by=tender_key_expr)
                    .label('version_count')
                )
                .filter(GenerationHistoryRecord.id.in_(session.query(filtered_subquery.c.id)))
            ).subquery()

            # Определяем сортировку
            order_by_clauses = []
            
            # Валидация и нормализация параметров сортировки
            valid_sort_fields = {
                'timestamp': GenerationHistoryRecord.timestamp,
                'tender_number': GenerationHistoryRecord.tender_number,
                'company': GenerationHistoryRecord.company,
                'margin_percent': GenerationHistoryRecord.margin_percent,
                'weight': GenerationHistoryRecord.weight,
                'duty_percent': GenerationHistoryRecord.duty_percent,
            }
            
            # Для цены продажи используем вычисляемое выражение
            sale_price_expr = func.coalesce(
                GenerationHistoryRecord.total_general_price,
                GenerationHistoryRecord.final_price * func.coalesce(GenerationHistoryRecord.quantity, 1),
            )
            
            # Для цены закупки нужно вычислять из позиций, но для простоты используем cost_price
            purchase_price_expr = GenerationHistoryRecord.cost_price * func.coalesce(GenerationHistoryRecord.quantity, 1)
            
            if sort_by:
                sort_by_lower = sort_by.lower()
                sort_order_lower = (sort_order or 'desc').lower() if sort_order else 'desc'
                
                if sort_order_lower not in ('asc', 'desc'):
                    sort_order_lower = 'desc'
                
                if sort_by_lower == 'price_sale' or sort_by_lower == 'total_general_price':
                    order_by_clauses.append(sale_price_expr.desc() if sort_order_lower == 'desc' else sale_price_expr.asc())
                elif sort_by_lower == 'price_purchase' or sort_by_lower == 'total_purchase_price':
                    order_by_clauses.append(purchase_price_expr.desc() if sort_order_lower == 'desc' else purchase_price_expr.asc())
                elif sort_by_lower in valid_sort_fields:
                    field = valid_sort_fields[sort_by_lower]
                    order_by_clauses.append(field.desc() if sort_order_lower == 'desc' else field.asc())
            
            # Если сортировка не указана или невалидна, используем сортировку по умолчанию
            if not order_by_clauses:
                order_by_clauses = [
                    GenerationHistoryRecord.timestamp.desc(),
                    GenerationHistoryRecord.id.desc(),
                ]
            else:
                # Добавляем сортировку по id для детерминированного порядка
                order_by_clauses.append(GenerationHistoryRecord.id.desc())
            
            records_query = (
                session.query(
                    GenerationHistoryRecord,
                    window_subquery.c.tender_key,
                    window_subquery.c.version_count,
                )
                .options(joinedload(GenerationHistoryRecord.user))
                .join(window_subquery, GenerationHistoryRecord.id == window_subquery.c.id)
                .filter(window_subquery.c.row_number == 1)
                .order_by(*order_by_clauses)
                .offset(offset)
                .limit(limit)
            )
            records = records_query.all()

            import json

            result: List[Dict[str, object]] = []
            for record, tender_key, version_count in records:
                positions: List[Dict[str, object]] = []

                if record.positions_data:
                    try:
                        positions = json.loads(record.positions_data) or []
                    except (json.JSONDecodeError, TypeError, ValueError) as exc:  # pragma: no cover - защитное ветвление
                        logging.warning('Failed to parse positions_data JSON for record %d in history: %s', record.id, exc)
                        positions = []

                if not positions:
                    quantity = record.quantity or 0
                    cost_price = record.cost_price or 0
                    duty_percent = record.duty_percent or 0
                    weight_value = 0.0
                    if quantity > 0:
                        weight_value = (record.weight or 0) / quantity

                    positions = [{
                        'quantity': quantity,
                        'cost_price': cost_price,
                        'weight': weight_value,
                        'duty_percent': duty_percent,
                    }]

                total_purchase_price = 0.0
                total_weight = float(record.weight or 0)
                duty_weight_sum = 0.0
                duty_value_sum = 0.0

                for position in positions:
                    quantity = float(position.get('quantity') or 0)
                    cost_price = float(position.get('cost_price') or 0)
                    duty_percent = float(position.get('duty_percent') or 0)
                    weight_per_unit = position.get('weight')

                    total_purchase_price += cost_price * quantity

                    if total_weight == 0 and weight_per_unit is not None:
                        total_weight += float(weight_per_unit or 0) * quantity

                    if quantity > 0:
                        duty_weight_sum += quantity
                        duty_value_sum += duty_percent * quantity
                    else:
                        duty_value_sum += duty_percent

                avg_duty_percent = duty_value_sum / duty_weight_sum if duty_weight_sum else float(record.duty_percent or 0)

                total_sale_price = record.total_general_price
                if total_sale_price is None:
                    total_sale_price = record.final_price * (record.quantity or 1)

                positions_count = record.positions_count if record.positions_count is not None else (len(positions) if positions else 1)

                result.append({
                    'id': record.id,
                    'timestamp': _format_history_timestamp(record.timestamp),
                    'tender_number': record.tender_number or 'Не указан',
                    'company': record.company,
                    'product': record.product,
                    'quantity': record.quantity,
                    'cost_price': record.cost_price,
                    'margin_percent': record.margin_percent,
                    'final_price': record.final_price,
                    'total_general_price': total_sale_price,
                    'total_purchase_price': total_purchase_price,
                    'total_weight': total_weight,
                    'avg_duty_percent': avg_duty_percent,
                    'drawing_number': record.drawing_number or 'Не указан',
                    'duty_percent': record.duty_percent,
                    'weight': record.weight,
                    'username': record.user.username if record.user else None,
                    'last_name': record.user.last_name if record.user else None,
                    'first_name': record.user.first_name if record.user else None,
                    'positions_count': positions_count,
                    'tender_key': tender_key,
                    'version_count': int(version_count or 1),
                })

            total_pages = math.ceil(total / limit) if limit else 1

            return {
                'items': result,
                'pagination': {
                    'page': page,
                    'per_page': limit,
                    'total': total,
                    'pages': max(total_pages, 1),
                    'has_prev': page > 1,
                    'has_next': page < max(total_pages, 1),
                }
            }
    except SQLAlchemyError as exc:
        logging.error('Database error getting generation history: %s', exc)
        raise DatabaseError(
            'Ошибка при получении истории генераций',
            operation='get_generation_history',
            details={'page': page, 'per_page': limit}
        ) from exc
    except Exception as exc:
        logging.error('Unexpected error getting generation history: %s', exc)
        raise DatabaseError(
            'Неожиданная ошибка при получении истории генераций',
            operation='get_generation_history'
        ) from exc


def save_generation_history(form_data, final_price, config, user_id=None, total_general_price=None) -> bool:
    """Сохраняет расчёт генерации вместе с расчётными параметрами пользователя."""
    try:
        import json
        from app.presentation.helpers import extract_positions_from_form
        
        # Извлекаем позиции из формы (используем готовые позиции из form_data если есть)
        if 'positions' in form_data and isinstance(form_data['positions'], list):
            positions = form_data['positions']
        else:
            positions = extract_positions_from_form(form_data)
        
        # Основные данные (для совместимости)
        quantity = int(form_data.get('quantity', 0))
        cost_price = float(form_data.get('cost_price', 0))
        weight = float(form_data.get('weight', 0))
        logistics = float(form_data.get('logistics', 0))
        margin_percent = float(form_data.get('margin_percent', config.get('margin_percent', 30)))
        duty_percent = float(form_data.get('duty_percent', config.get('default_duty_percent', 0)))
        delivery_time = int(form_data.get('delivery_time', 0))
        
        # Если есть множественные позиции, используем данные первой позиции для основных полей
        if len(positions) > 1:
            first_position = positions[0]
            quantity = int(first_position.get('quantity', 0))
            cost_price = float(first_position.get('cost_price', 0))
            weight = float(first_position.get('weight', 0))
            duty_percent = float(first_position.get('duty_percent', 0))
        
        # Рассчитываем общий вес всех позиций
        total_weight = sum(float(p.get('weight', 0)) * int(p.get('quantity', 0)) for p in positions)
        
        with _session_scope() as session:
            record = GenerationHistoryRecord(
                tender_number=form_data.get('tender_number', ''),
                company=form_data.get('company', ''),
                product=form_data.get('product', ''),
                quantity=quantity,
                cost_price=cost_price,
                weight=total_weight,  # Общий вес всех позиций
                logistics=logistics,
                margin_percent=margin_percent,
                final_price=final_price,
                drawing_number=form_data.get('drawing_number', ''),
                material=form_data.get('material', ''),
                delivery_address=form_data.get('delivery_address', ''),
                duty_percent=duty_percent,
                delivery_time=delivery_time,
                payment_terms=form_data.get('payment_terms', ''),
                proposal_validity=form_data.get('proposal_validity', ''),
                warranty_period=form_data.get('warranty_period', ''),
                comment=form_data.get('comment', ''),
                user_id=user_id,
                # Новые поля для множественных позиций
                positions_data=json.dumps(positions, ensure_ascii=False),
                total_general_price=total_general_price if total_general_price is not None else (final_price * quantity),
                positions_count=len(positions),
            )
            session.add(record)
        return True
    except SQLAlchemyError as exc:
        logging.error('Database error saving generation history: %s', exc)
        raise DatabaseError(
            'Ошибка при сохранении истории генерации',
            operation='save_generation_history'
        ) from exc
    except Exception as exc:
        logging.error('Unexpected error saving generation history: %s', exc)
        raise DatabaseError(
            'Неожиданная ошибка при сохранении истории генерации',
            operation='save_generation_history'
        ) from exc


def get_generation_details(record_id: int) -> Optional[Dict[str, object]]:
    """Возвращает детальную информацию о генерации по идентификатору."""
    try:
        with _session_scope() as session:
            record = (
                session.query(GenerationHistoryRecord)
                .options(joinedload(GenerationHistoryRecord.user))
                .filter(GenerationHistoryRecord.id == record_id)
                .one_or_none()
            )

            if record is None:
                return None

            return _build_generation_detail(record)
    except SQLAlchemyError as exc:
        logging.error('Database error getting generation details: %s', exc)
        raise DatabaseError(
            'Ошибка при получении деталей генерации',
            operation='get_generation_details',
            details={'record_id': record_id}
        ) from exc
    except Exception as exc:
        logging.error('Unexpected error getting generation details: %s', exc)
        raise DatabaseError(
            'Неожиданная ошибка при получении деталей генерации',
            operation='get_generation_details'
        ) from exc


def load_generation_data(gen_id: int) -> Optional[Dict[str, object]]:
    """Загружает сохранённую генерацию для повторного редактирования."""
    try:
        import json
        
        with _session_scope() as session:
            # Используем joinedload для оптимизации загрузки связанных данных
            record = (
                session.query(GenerationHistoryRecord)
                .options(joinedload(GenerationHistoryRecord.user))
                .filter(GenerationHistoryRecord.id == gen_id)
                .one_or_none()
            )
            if record is None:
                raise NotFoundError(
                    f'Генерация с ID {gen_id} не найдена',
                    resource_type='generation',
                    resource_id=str(gen_id)
                )

            # Базовые данные
            data = {
                'id': record.id,
                'timestamp': _format_timestamp(record.timestamp),
                'tender_number': record.tender_number,
                'company': record.company,
                'logistics': record.logistics,
                'margin_percent': record.margin_percent,
                'delivery_address': record.delivery_address,
                'delivery_time': record.delivery_time,
                'payment_terms': record.payment_terms or '',
                'proposal_validity': record.proposal_validity or '',
                'warranty_period': record.warranty_period or '',
                'comment': record.comment,
                'user_id': record.user_id,
            }
            
            # Если есть данные позиций, используем их
            if record.positions_data:
                try:
                    positions = json.loads(record.positions_data)
                    # Загружаем все позиции в форму
                    for i, position in enumerate(positions):
                        if i == 0:
                            # Первая позиция в основные поля
                            for field in ['product', 'quantity', 'cost_price', 'weight', 'drawing_number', 'material', 'duty_percent']:
                                if field in position:
                                    data[field] = position[field]
                        else:
                            # Остальные позиции с суффиксами
                            for field in ['product', 'quantity', 'cost_price', 'weight', 'drawing_number', 'material', 'duty_percent']:
                                if field in position:
                                    data[f"{field}_{i+1}"] = position[field]
                except (json.JSONDecodeError, TypeError) as exc:
                    # Если не удалось распарсить JSON, используем старые поля
                    logging.warning('Failed to parse positions_data JSON for record %d in load_generation_data: %s', record.id, exc)
                    data.update({
                        'product': record.product,
                        'quantity': record.quantity,
                        'cost_price': record.cost_price,
                        'weight': record.weight,
                        'drawing_number': record.drawing_number,
                        'material': record.material,
                        'duty_percent': record.duty_percent,
                        'final_price': record.final_price,
                    })
            else:
                # Старые данные (одна позиция)
                data.update({
                    'product': record.product,
                    'quantity': record.quantity,
                    'cost_price': record.cost_price,
                    'weight': record.weight,
                    'drawing_number': record.drawing_number,
                    'material': record.material,
                    'duty_percent': record.duty_percent,
                    'final_price': record.final_price,
                })
            
            return data
    except NotFoundError:
        raise
    except SQLAlchemyError as exc:
        logging.error('Database error loading generation data: %s', exc)
        raise DatabaseError(
            'Ошибка при загрузке данных генерации',
            operation='load_generation_data',
            details={'gen_id': gen_id}
        ) from exc
    except Exception as exc:
        logging.error('Unexpected error loading generation data: %s', exc)
        raise DatabaseError(
            'Неожиданная ошибка при загрузке данных генерации',
            operation='load_generation_data'
        ) from exc


def create_user(
    username,
    password_hash,
    last_name='',
    first_name='',
    role='user',
    contact_info=None,
) -> Optional[int]:
    """Создаёт нового пользователя и возвращает его идентификатор."""
    try:
        with _session_scope() as session:
            user = UserRecord(
                username=username,
                password_hash=password_hash,
                last_name=last_name or None,
                first_name=first_name or None,
                contact_info=contact_info or None,
                role=(role or 'user').lower(),
            )
            session.add(user)
            session.flush()
            return user.id
    except IntegrityError:
        # Пользователь с таким логином уже существует.
        # Для идемпотентности возвращаем его идентификатор, чтобы верхний уровень
        # мог решить, считать это ошибкой или использовать существующего пользователя.
        logging.warning('User with username %s already exists, returning existing user_id', username)
        existing_user = get_user_by_username(username)
        if existing_user:
            # get_user_by_username возвращает кортеж, где [0] — ID пользователя
            return existing_user[0]
        return None
    except Exception as exc:
        logging.error('Error creating user: %s', exc)
        return None


def get_user_by_username(username):
    """Возвращает пользователя по логину для авторизации."""
    try:
        user = None
        with _session_scope() as session:
            user = session.query(UserRecord).filter(UserRecord.username == username).one_or_none()
            if user is not None:
                session.expunge(user)
        return _user_to_tuple(user)
    except Exception as exc:
        logging.error('Error fetching user by username: %s', exc)
        return None


def get_user_by_id(user_id):
    """Находит пользователя по идентификатору (используется Flask-Login)."""
    try:
        user = None
        with _session_scope() as session:
            user = session.query(UserRecord).filter(UserRecord.id == user_id).one_or_none()
            if user is not None:
                session.expunge(user)
        return _user_to_tuple(user)
    except Exception as exc:
        logging.error('Error fetching user by id: %s', exc)
        return None


def update_last_login(user_id, last_login=None) -> bool:
    """Обновляет отметку времени последнего входа пользователя."""
    try:
        with _session_scope() as session:
            user = session.query(UserRecord).filter(UserRecord.id == user_id).one_or_none()
            if user is None:
                return False
            timestamp = last_login or datetime.now(timezone.utc)
            if timestamp.tzinfo is None:
                timestamp = timestamp.replace(tzinfo=timezone.utc)
            user.last_login = timestamp
        return True
    except Exception as exc:
        logging.error('Error updating last login: %s', exc)
        return False


def get_user_statistics(user_id) -> Dict[str, object]:
    """Собирает агрегированную статистику активности пользователя."""
    try:
        with _session_scope() as session:
            from datetime import datetime, timedelta
            
            # Базовая статистика с использованием индекса по user_id
            total, last_timestamp = (
                session.query(
                    func.count(GenerationHistoryRecord.id),
                    func.max(GenerationHistoryRecord.timestamp),
                )
                .filter(GenerationHistoryRecord.user_id == user_id)
                .one()
            )

            # Статистика по ценам и марже
            # Используем total_general_price если есть, иначе вычисляем из final_price * quantity
            price_stats = (
                session.query(
                    func.avg(GenerationHistoryRecord.margin_percent).label('avg_margin'),
                    func.sum(
                        func.coalesce(
                            GenerationHistoryRecord.total_general_price,
                            GenerationHistoryRecord.final_price * func.coalesce(GenerationHistoryRecord.quantity, 1)
                        )
                    ).label('total_sum'),
                    func.avg(
                        func.coalesce(
                            GenerationHistoryRecord.total_general_price,
                            GenerationHistoryRecord.final_price * func.coalesce(GenerationHistoryRecord.quantity, 1)
                        )
                    ).label('avg_price'),
                    func.max(
                        func.coalesce(
                            GenerationHistoryRecord.total_general_price,
                            GenerationHistoryRecord.final_price * func.coalesce(GenerationHistoryRecord.quantity, 1)
                        )
                    ).label('max_price'),
                )
                .filter(GenerationHistoryRecord.user_id == user_id)
                .one()
            )

            # Генерации за текущий месяц
            now = datetime.now()
            month_start = datetime(now.year, now.month, 1)
            month_generations = (
                session.query(func.count(GenerationHistoryRecord.id))
                .filter(GenerationHistoryRecord.user_id == user_id)
                .filter(GenerationHistoryRecord.timestamp >= month_start)
                .scalar() or 0
            )

            # Генерации за последние 7 дней
            week_ago = now - timedelta(days=7)
            week_generations = (
                session.query(func.count(GenerationHistoryRecord.id))
                .filter(GenerationHistoryRecord.user_id == user_id)
                .filter(GenerationHistoryRecord.timestamp >= week_ago)
                .scalar() or 0
            )

            recent_records = (
                session.query(
                    GenerationHistoryRecord.id,
                    GenerationHistoryRecord.company,
                    GenerationHistoryRecord.product,
                    GenerationHistoryRecord.margin_percent,
                    GenerationHistoryRecord.final_price,
                    GenerationHistoryRecord.timestamp,
                )
                .filter(GenerationHistoryRecord.user_id == user_id)
                .order_by(GenerationHistoryRecord.timestamp.desc())
                .limit(5)
                .all()
            )

            formatted_recent = [
                (
                    rec.id,
                    rec.company,
                    rec.product,
                    rec.margin_percent,
                    rec.final_price,
                    _format_timestamp(rec.timestamp),
                )
                for rec in recent_records
            ]

            return {
                'total_generations': total or 0,
                'last_generation_at': _format_timestamp(last_timestamp) if last_timestamp else None,
                'recent_generations': formatted_recent,
                'avg_margin': float(price_stats.avg_margin) if price_stats.avg_margin else 0.0,
                'total_sum': float(price_stats.total_sum) if price_stats.total_sum else 0.0,
                'avg_price': float(price_stats.avg_price) if price_stats.avg_price else 0.0,
                'max_price': float(price_stats.max_price) if price_stats.max_price else 0.0,
                'month_generations': month_generations,
                'week_generations': week_generations,
            }
    except Exception as exc:
        logging.error('Error getting user statistics: %s', exc)
        return {
            'total_generations': 0,
            'last_generation_at': None,
            'recent_generations': [],
            'avg_margin': 0.0,
            'total_sum': 0.0,
            'avg_price': 0.0,
            'max_price': 0.0,
            'month_generations': 0,
            'week_generations': 0,
        }


def get_admin_user_activity(limit: int = 10) -> Dict[str, object]:
    """Формирует агрегаты активности пользователей для административной панели."""
    try:
        with _session_scope() as session:
            total_generations = session.query(func.count(GenerationHistoryRecord.id)).scalar() or 0
            users_count = session.query(func.count(UserRecord.id)).scalar() or 0

            last_logins = (
                session.query(
                    UserRecord.username,
                    UserRecord.last_login,
                    func.coalesce(UserRecord.last_name, ''),
                    func.coalesce(UserRecord.first_name, '')
                )
                .order_by(UserRecord.last_login.desc().nullslast())
                .limit(limit)
                .all()
            )

            users_activity = (
                session.query(
                    UserRecord.username,
                    func.coalesce(UserRecord.last_name, ''),
                    func.coalesce(UserRecord.first_name, ''),
                    func.count(GenerationHistoryRecord.id).label('generations')
                )
                .outerjoin(GenerationHistoryRecord, GenerationHistoryRecord.user_id == UserRecord.id)
                .group_by(UserRecord.id)
                .order_by(func.count(GenerationHistoryRecord.id).desc())
                .limit(limit)
                .all()
            )

            popular_materials = (
                session.query(
                    GenerationHistoryRecord.material,
                    func.count(GenerationHistoryRecord.id).label('total')
                )
                .filter(GenerationHistoryRecord.material.isnot(None))
                .filter(func.trim(GenerationHistoryRecord.material) != '')
                .group_by(GenerationHistoryRecord.material)
                .order_by(func.count(GenerationHistoryRecord.id).desc())
                .limit(limit)
                .all()
            )

            def _display_name(username, last_name, first_name):
                full_name = ' '.join(value for value in [last_name, first_name] if value).strip()
                return full_name or username

            formatted_logins = [
                {
                    'username': _display_name(username, last_name, first_name),
                    'last_login': _format_timestamp(last_login) if last_login else '—'
                }
                for username, last_login, last_name, first_name in last_logins
            ]

            formatted_activity = [
                {
                    'username': _display_name(username, last_name, first_name),
                    'generations': generations
                }
                for username, last_name, first_name, generations in users_activity
            ]

            formatted_materials = [
                {'material': material, 'count': total}
                for material, total in popular_materials
            ]

            return {
                'total_generations': total_generations,
                'users_count': users_count,
                'last_logins': formatted_logins,
                'users_activity': formatted_activity,
                'popular_materials': formatted_materials,
            }
    except Exception as exc:  # pragma: no cover - логирование ошибок
        logging.error('Failed to build admin user activity stats: %s', exc)
        return {
            'total_generations': 0,
            'users_count': 0,
            'last_logins': [],
            'users_activity': [],
            'popular_materials': [],
        }


def update_user_profile(
    user_id,
    username,
    last_name,
    first_name,
    contact_info,
    password_hash=None
) -> bool:
    """Обновляет данные профиля и при необходимости пароль пользователя."""
    try:
        with _session_scope() as session:
            user = session.query(UserRecord).filter(UserRecord.id == user_id).one_or_none()
            if user is None:
                return False

            user.username = username
            user.last_name = last_name or None
            user.first_name = first_name or None
            user.contact_info = contact_info or None
            if password_hash:
                user.password_hash = password_hash
        return True
    except Exception as exc:
        logging.error('Error updating user profile: %s', exc)
        return False


def delete_user(user_id) -> bool:
    """Удаляет пользователя и отвязывает его историю генераций."""
    try:
        with _session_scope() as session:
            histories = (
                session.query(GenerationHistoryRecord)
                .filter(GenerationHistoryRecord.user_id == user_id)
                .all()
            )
            for record in histories:
                record.user_id = None

            user = session.query(UserRecord).filter(UserRecord.id == user_id).one_or_none()
            if user is None:
                return False
            session.delete(user)
        return True
    except Exception as exc:
        logging.error('Error deleting user: %s', exc)
        return False


def get_generations_by_drawing(drawing_number: str, limit: int = 5) -> List[Dict[str, object]]:
    """Возвращает последние генерации, созданные с указанным номером чертежа.

    Поиск ведётся как по основному полю ``drawing_number``, так и по всем позициям
    в JSON-поле ``positions_data`` (для множественных позиций).
    """
    if not drawing_number:
        return []

    try:
        normalized = drawing_number.strip().lower()
        if not normalized:
            return []

        with _session_scope() as session:
            base_query = (
                session.query(GenerationHistoryRecord)
                .options(joinedload(GenerationHistoryRecord.user))
            )

            # 1. Сначала ищем по основному полю drawing_number (как раньше)
            primary_records: Sequence[GenerationHistoryRecord] = (
                base_query
                .filter(func.lower(GenerationHistoryRecord.drawing_number) == normalized)
                .order_by(GenerationHistoryRecord.timestamp.desc())
                .all()
            )

            records_by_id = {record.id: record for record in primary_records}

            # 2. Дополнительно ищем по JSON с позициями.
            #
            # Для производительности сначала отфильтруем по LIKE,
            # а затем точно проверим номер чертежа на уровне Python, распарсив JSON.
            import json

            candidates: Sequence[GenerationHistoryRecord] = (
                base_query
                .filter(GenerationHistoryRecord.positions_data.isnot(None))
                .filter(GenerationHistoryRecord.positions_data != '')
                .filter(
                    func.lower(GenerationHistoryRecord.positions_data).like(
                        f'%{normalized}%'
                    )
                )
                .order_by(GenerationHistoryRecord.timestamp.desc())
                .all()
            )

            for record in candidates:
                if record.id in records_by_id:
                    # Уже есть из основного запроса
                    continue

                try:
                    positions = json.loads(record.positions_data or '[]') or []
                except (TypeError, json.JSONDecodeError):
                    continue

                for pos in positions:
                    value = (pos.get('drawing_number') or '').strip().lower()
                    if value == normalized:
                        records_by_id[record.id] = record
                        break

            # 3. Сортируем по дате и ограничиваем количеством
            all_records = sorted(
                records_by_id.values(),
                key=lambda r: r.timestamp or datetime.min,
                reverse=True,
            )[:limit]

            return [detail for record in all_records if (detail := _build_generation_detail(record))]
    except Exception as exc:
        logging.error('Error fetching generations by drawing: %s', exc)
        return []


def get_generations_by_tender(tender_number: str) -> List[Dict[str, object]]:
    """Возвращает все генерации по указанному номеру тендера."""
    if not tender_number:
        return []

    try:
        with _session_scope() as session:
            records: Sequence[GenerationHistoryRecord] = (
                session.query(GenerationHistoryRecord)
                .options(joinedload(GenerationHistoryRecord.user))
                .filter(func.lower(GenerationHistoryRecord.tender_number) == tender_number.lower())
                .order_by(GenerationHistoryRecord.timestamp.desc())
                .all()
            )

            return [
                {
                    'id': record.id,
                    'timestamp': _format_history_timestamp(record.timestamp),
                    'company': record.company,
                    'product': record.product,
                    'margin_percent': record.margin_percent,
                    'final_price': record.final_price,
                    'total_general_price': record.total_general_price if record.total_general_price is not None else (record.final_price * (record.quantity or 1)),
                    'positions_count': record.positions_count or 1,
                }
                for record in records
            ]
    except Exception as exc:
        logging.error('Error fetching generations by tender: %s', exc)
        return []


def get_unique_companies() -> List[str]:
    """Возвращает список уникальных компаний из истории генераций."""
    try:
        with _session_scope() as session:
            companies = (
                session.query(GenerationHistoryRecord.company)
                .distinct()
                .filter(GenerationHistoryRecord.company.isnot(None))
                .filter(GenerationHistoryRecord.company != '')
                .order_by(GenerationHistoryRecord.company)
                .all()
            )
            return [company[0] for company in companies if company[0]]
    except Exception as exc:
        logging.error('Error fetching unique companies: %s', exc)
        return []


def get_users_list(
    *,
    page: int = 1,
    per_page: int = 25,
    search: Optional[str] = None,
    role_filter: Optional[str] = None,
) -> Dict[str, object]:
    """Возвращает список пользователей с пагинацией, поиском и фильтрацией."""
    try:
        page = max(int(page or 1), 1)
        per_page = max(1, min(int(per_page or 25), 100))
        offset = (page - 1) * per_page

        with _session_scope() as session:
            query = session.query(UserRecord)

            # Поиск по имени пользователя, фамилии, имени
            if search:
                search_term = f'%{search.strip().lower()}%'
                query = query.filter(
                    func.lower(UserRecord.username).like(search_term)
                    | func.lower(func.coalesce(UserRecord.last_name, '')).like(search_term)
                    | func.lower(func.coalesce(UserRecord.first_name, '')).like(search_term)
                )

            # Фильтр по роли
            if role_filter and role_filter.lower() in ('admin', 'user'):
                query = query.filter(func.lower(UserRecord.role) == role_filter.lower())

            # Общее количество
            total = query.count()

            # Получаем пользователей с пагинацией
            users = query.order_by(UserRecord.username).offset(offset).limit(per_page).all()

            result = []
            for user in users:
                full_name = ' '.join(
                    value for value in [user.last_name, user.first_name] if value
                ).strip() or user.username

                result.append({
                    'id': user.id,
                    'username': user.username,
                    'role': user.role,
                    'last_name': user.last_name,
                    'first_name': user.first_name,
                    'full_name': full_name,
                    'contact_info': user.contact_info,
                    'created_at': _format_timestamp(user.created_at),
                    'last_login': _format_timestamp(user.last_login),
                })

            total_pages = math.ceil(total / per_page) if per_page else 1

            return {
                'items': result,
                'pagination': {
                    'page': page,
                    'per_page': per_page,
                    'total': total,
                    'pages': max(total_pages, 1),
                    'has_prev': page > 1,
                    'has_next': page < max(total_pages, 1),
                }
            }
    except Exception as exc:
        logging.error('Error getting users list: %s', exc)
        return {
            'items': [],
            'pagination': {
                'page': page,
                'per_page': per_page,
                'total': 0,
                'pages': 1,
                'has_prev': False,
                'has_next': False,
            }
        }
