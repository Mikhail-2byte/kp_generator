# Database operations
import logging
import sqlite3
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional

from alembic import command
from alembic.config import Config
from alembic.runtime.migration import MigrationContext
from alembic.script import ScriptDirectory
from sqlalchemy import create_engine, func
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import joinedload

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
        except (json.JSONDecodeError, TypeError):
            # Если не удалось распарсить JSON, используем старые поля
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
    return (
        user.id,
        user.username,
        user.password_hash,
        user.created_at,
        user.last_login,
        user.last_name,
        user.first_name,
        user.contact_info,
        user.role,
    )


def get_generation_history(config) -> List[Dict[str, object]]:
    """Возвращает историю генераций с учётом ограничений конфигурации."""
    limit = config.get('max_history_items', 50)
    try:
        with _session_scope() as session:
            records = (
                session.query(GenerationHistoryRecord)
                .options(joinedload(GenerationHistoryRecord.user))
                .order_by(GenerationHistoryRecord.timestamp.desc())
                .limit(limit)
                .all()
            )

            import json

            result: List[Dict[str, object]] = []
            for record in records:
                positions: List[Dict[str, object]] = []

                if record.positions_data:
                    try:
                        positions = json.loads(record.positions_data) or []
                    except (json.JSONDecodeError, TypeError, ValueError):  # pragma: no cover - защитное ветвление
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

                positions_count = record.positions_count or len(positions) or 1

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
                })

            return result
    except Exception as exc:
        logging.error('Error getting generation history: %s', exc)
        return []


def save_generation_history(form_data, final_price, config, user_id=None, total_general_price=None) -> bool:
    """Сохраняет расчёт генерации вместе с расчётными параметрами пользователя."""
    try:
        import json
        from app.presentation.helpers import extract_positions_from_form
        
        # Извлекаем позиции из формы
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
                comment=form_data.get('comment', ''),
                user_id=user_id,
                # Новые поля для множественных позиций
                positions_data=json.dumps(positions, ensure_ascii=False),
                total_general_price=total_general_price or (final_price * quantity),
                positions_count=len(positions),
            )
            session.add(record)
        return True
    except Exception as exc:
        logging.error('Error saving generation history: %s', exc)
        return False


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
    except Exception as exc:
        logging.error('Error getting generation details: %s', exc)
        return None


def load_generation_data(gen_id: int) -> Optional[Dict[str, object]]:
    """Загружает сохранённую генерацию для повторного редактирования."""
    try:
        import json
        
        with _session_scope() as session:
            record = session.query(GenerationHistoryRecord).filter(GenerationHistoryRecord.id == gen_id).one_or_none()
            if record is None:
                return None

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
                except (json.JSONDecodeError, TypeError):
                    # Если не удалось распарсить JSON, используем старые поля
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
    except Exception as exc:
        logging.error('Error loading generation data: %s', exc)
        return None


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
        logging.error('User with username %s already exists', username)
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
            if last_login is None:
                user.last_login = datetime.now(timezone.utc)
            else:
                user.last_login = last_login
        return True
    except Exception as exc:
        logging.error('Error updating last login: %s', exc)
        return False


def get_user_statistics(user_id) -> Dict[str, object]:
    """Собирает агрегированную статистику активности пользователя."""
    try:
        with _session_scope() as session:
            total, last_timestamp = (
                session.query(
                    func.count(GenerationHistoryRecord.id),
                    func.max(GenerationHistoryRecord.timestamp),
                )
                .filter(GenerationHistoryRecord.user_id == user_id)
                .one()
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
            }
    except Exception as exc:
        logging.error('Error getting user statistics: %s', exc)
        return {
            'total_generations': 0,
            'last_generation_at': None,
            'recent_generations': [],
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
    """Возвращает последние генерации, созданные с указанным номером чертежа."""
    if not drawing_number:
        return []

    try:
        with _session_scope() as session:
            records = (
                session.query(GenerationHistoryRecord)
                .options(joinedload(GenerationHistoryRecord.user))
                .filter(func.lower(GenerationHistoryRecord.drawing_number) == drawing_number.lower())
                .order_by(GenerationHistoryRecord.timestamp.desc())
                .limit(limit)
                .all()
            )

            return [detail for record in records if (detail := _build_generation_detail(record))]
    except Exception as exc:
        logging.error('Error fetching generations by drawing: %s', exc)
        return []
