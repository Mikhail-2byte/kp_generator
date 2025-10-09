# app/database.py
import logging
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from alembic import command
from alembic.config import Config
from sqlalchemy import func
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import joinedload

from app.extensions import SessionLocal, get_database_url
from app.models import GenerationHistoryRecord, UserRecord


PROJECT_ROOT = Path(__file__).resolve().parents[1]


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


def connect_db():
    """Устанавливает соединение с базой данных через sqlite3 (наследие)."""
    database_url = get_database_url()
    path = _resolve_sqlite_path(database_url)
    return sqlite3.connect(path, detect_types=sqlite3.PARSE_DECLTYPES)


def init_db():
    """Запускает миграции Alembic и приводит схему к актуальному состоянию."""
    alembic_cfg = PROJECT_ROOT / 'alembic.ini'
    if not alembic_cfg.exists():
        logging.warning('Alembic configuration not found at %s', alembic_cfg)
        return

    config = Config(str(alembic_cfg))
    config.set_main_option('sqlalchemy.url', get_database_url())

    try:
        command.upgrade(config, 'head')
    except Exception as exc:
        logging.exception('Failed to apply database migrations: %s', exc)
        raise


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

            result: List[Dict[str, object]] = []
            for record in records:
                result.append({
                    'id': record.id,
                    'timestamp': _format_history_timestamp(record.timestamp),
                    'tender_number': record.tender_number or 'Не указан',
                    'company': record.company,
                    'product': record.product,
                    'margin_percent': record.margin_percent,
                    'final_price': record.final_price,
                    'drawing_number': record.drawing_number or 'Не указан',
                    'duty_percent': record.duty_percent,
                    'username': record.user.username if record.user else None,
                    'last_name': record.user.last_name if record.user else None,
                    'first_name': record.user.first_name if record.user else None,
                })
            return result
    except Exception as exc:
        logging.error('Error getting generation history: %s', exc)
        return []


def save_generation_history(form_data, final_price, config, user_id=None) -> bool:
    """Сохраняет расчёт генерации вместе с расчётными параметрами пользователя."""
    try:
        quantity = int(form_data.get('quantity', 0))
        cost_price = float(form_data.get('cost_price', 0))
        weight = float(form_data.get('weight', 0))
        logistics = float(form_data.get('logistics', 0))
        margin_percent = float(form_data.get('margin_percent', config.get('margin_percent', 30)))
        duty_percent = float(form_data.get('duty_percent', config.get('default_duty_percent', 0)))
        delivery_time = int(form_data.get('delivery_time', 0))

        with _session_scope() as session:
            record = GenerationHistoryRecord(
                tender_number=form_data.get('tender_number', ''),
                company=form_data.get('company', ''),
                product=form_data.get('product', ''),
                quantity=quantity,
                cost_price=cost_price,
                weight=weight,
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

            return {
                'id': record.id,
                'timestamp': _format_timestamp(record.timestamp),
                'tender_number': record.tender_number,
                'company': record.company,
                'product': record.product,
                'quantity': record.quantity,
                'cost_price': record.cost_price,
                'weight': record.weight,
                'logistics': record.logistics,
                'margin_percent': record.margin_percent,
                'final_price': record.final_price,
                'drawing_number': record.drawing_number,
                'material': record.material,
                'delivery_address': record.delivery_address,
                'duty_percent': record.duty_percent,
                'delivery_time': record.delivery_time,
                'comment': record.comment,
                'user_id': record.user_id,
                'username': record.user.username if record.user else None,
                'last_name': record.user.last_name if record.user else None,
                'first_name': record.user.first_name if record.user else None,
            }
    except Exception as exc:
        logging.error('Error getting generation details: %s', exc)
        return None


def load_generation_data(gen_id: int) -> Optional[Dict[str, object]]:
    """Загружает сохранённую генерацию для повторного редактирования."""
    try:
        with _session_scope() as session:
            record = session.query(GenerationHistoryRecord).filter(GenerationHistoryRecord.id == gen_id).one_or_none()
            if record is None:
                return None

            return {
                'id': record.id,
                'timestamp': _format_timestamp(record.timestamp),
                'tender_number': record.tender_number,
                'company': record.company,
                'product': record.product,
                'quantity': record.quantity,
                'cost_price': record.cost_price,
                'weight': record.weight,
                'logistics': record.logistics,
                'margin_percent': record.margin_percent,
                'final_price': record.final_price,
                'drawing_number': record.drawing_number,
                'material': record.material,
                'delivery_address': record.delivery_address,
                'duty_percent': record.duty_percent,
                'delivery_time': record.delivery_time,
                'comment': record.comment,
                'user_id': record.user_id,
            }
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
                user.last_login = datetime.utcnow()
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