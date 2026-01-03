from datetime import datetime, timezone
from typing import Any, Mapping, Sequence

from flask_login import UserMixin
from sqlalchemy import Boolean, Column, DateTime, Float, ForeignKey, Integer, String, Text, func
from sqlalchemy.orm import declarative_base, relationship

Base = declarative_base()


class UserRecord(Base):
    """ORM-модель пользователя, хранящего учётные данные и профиль."""
    __tablename__ = 'users'

    id = Column(Integer, primary_key=True, autoincrement=True)
    username = Column(Text, nullable=False, unique=True)
    password_hash = Column(Text, nullable=False)
    created_at = Column(DateTime, server_default=func.current_timestamp())
    last_login = Column(DateTime)
    last_name = Column(Text)
    first_name = Column(Text)
    contact_info = Column(Text)
    role = Column(Text, nullable=False, server_default='user')

    generations = relationship('GenerationHistoryRecord', back_populates='user')


class GenerationHistoryRecord(Base):
    """ORM-модель сохранённой генерации коммерческого предложения."""
    __tablename__ = 'generation_history'

    id = Column(Integer, primary_key=True, autoincrement=True)
    timestamp = Column(DateTime, server_default=func.current_timestamp())
    tender_number = Column(Text)
    company = Column(Text, nullable=False)
    product = Column(Text, nullable=False)  # Основной товар (для совместимости)
    quantity = Column(Integer, nullable=False)  # Общее количество (для совместимости)
    cost_price = Column(Float, nullable=False)  # Основная цена закупа (для совместимости)
    weight = Column(Float, nullable=False)  # Общий вес (для совместимости)
    logistics = Column(Float, nullable=False)
    margin_percent = Column(Float, nullable=False)
    final_price = Column(Float, nullable=False)  # Цена за единицу основной позиции (для совместимости)
    drawing_number = Column(Text)  # Основной номер чертежа (для совместимости)
    material = Column(Text)  # Основной материал (для совместимости)
    delivery_address = Column(Text)
    duty_percent = Column(Float, server_default='0')  # Основная пошлина (для совместимости)
    delivery_time = Column(Integer, server_default='0')
    payment_terms = Column(Text)
    proposal_validity = Column(Text)
    warranty_period = Column(Text)
    comment = Column(Text)
    user_id = Column(Integer, ForeignKey('users.id'))
    
    # Новые поля для множественных позиций
    positions_data = Column(Text)  # JSON с данными всех позиций
    total_general_price = Column(Float)  # Общая цена всех позиций
    positions_count = Column(Integer, server_default='1')  # Количество позиций

    user = relationship('UserRecord', back_populates='generations')


class AuditLogRecord(Base):
    """ORM-модель лога действий пользователей для аудита и мониторинга."""
    __tablename__ = 'audit_logs'

    id = Column(Integer, primary_key=True, autoincrement=True)
    user_id = Column(Integer, ForeignKey('users.id'), nullable=True)
    username = Column(Text, nullable=False)
    action_type = Column(Text, nullable=False)  # login, logout, create, update, delete, view, etc.
    resource_type = Column(Text, nullable=True)  # user, generation, duty, material, logistics, etc.
    resource_id = Column(Text, nullable=True)  # ID ресурса (может быть строка для разных типов)
    description = Column(Text, nullable=False)  # Описание действия
    ip_address = Column(Text, nullable=True)
    user_agent = Column(Text, nullable=True)
    changes_before = Column(Text, nullable=True)  # JSON с данными до изменения
    changes_after = Column(Text, nullable=True)  # JSON с данными после изменения
    created_at = Column(DateTime, server_default=func.current_timestamp(), nullable=False)

    user = relationship('UserRecord', foreign_keys=[user_id])


class CustomerContactRecord(Base):
    """ORM-модель контактной информации заказчиков."""
    __tablename__ = 'customer_contacts'

    id = Column(Integer, primary_key=True, autoincrement=True)
    company_name = Column(Text, nullable=False)  # Название компании
    contact_person = Column(Text)  # Контактное лицо
    phone = Column(Text)  # Телефон
    email = Column(Text)  # Email
    address = Column(Text)  # Адрес
    notes = Column(Text)  # Дополнительные заметки
    created_at = Column(DateTime, server_default=func.current_timestamp())
    updated_at = Column(DateTime, server_default=func.current_timestamp(), onupdate=func.current_timestamp())


class AIAgentUsageRecord(Base):
    """ORM-модель для мониторинга использования AI агента."""
    __tablename__ = 'ai_agent_usage'

    id = Column(Integer, primary_key=True, autoincrement=True)
    timestamp = Column(DateTime, server_default=func.current_timestamp(), nullable=False)
    user_id = Column(Integer, ForeignKey('users.id', ondelete='SET NULL'), nullable=True)
    request_tokens = Column(Integer, nullable=True)
    response_tokens = Column(Integer, nullable=True)
    total_cost = Column(Float, nullable=True)
    response_time_ms = Column(Integer, nullable=True)
    model_name = Column(String(100), nullable=True)
    error_type = Column(String(50), nullable=True)
    user_message = Column(Text, nullable=True)
    cache_hit = Column(Boolean, default=False, nullable=False)

    user = relationship('UserRecord', foreign_keys=[user_id])


class User(UserMixin):
    """Адаптер пользователей к требованиям Flask-Login."""
    def __init__(
        self,
        user_id,
        username,
        password_hash,
        created_at=None,
        last_login=None,
        last_name=None,
        first_name=None,
        role='user',
        contact_info=None
    ):
        """Инициализирует объект пользователя для хранения в сессии."""
        self.id = str(user_id)
        self.username = username
        self.password_hash = password_hash
        self.created_at = created_at
        self.last_login = last_login
        self.last_name = last_name
        self.first_name = first_name
        self.contact_info = contact_info
        self.role = (role or 'user').lower()

    @classmethod
    def from_row(cls, row: Any):
        """Создаёт пользователя из ORM-объекта, Row или кортежа старого формата."""
        if not row:
            return None

        if hasattr(row, 'username'):
            return cls(
                user_id=row.id,
                username=row.username,
                password_hash=row.password_hash,
                created_at=getattr(row, 'created_at', None),
                last_login=getattr(row, 'last_login', None),
                last_name=getattr(row, 'last_name', None),
                first_name=getattr(row, 'first_name', None),
                role=getattr(row, 'role', 'user'),
                contact_info=getattr(row, 'contact_info', None)
            )

        if hasattr(row, '_mapping'):
            mapping: Mapping[str, Any] = row._mapping
            return cls(
                user_id=mapping.get('id'),
                username=mapping.get('username'),
                password_hash=mapping.get('password_hash'),
                created_at=mapping.get('created_at'),
                last_login=mapping.get('last_login'),
                last_name=mapping.get('last_name'),
                first_name=mapping.get('first_name'),
                role=mapping.get('role', 'user'),
                contact_info=mapping.get('contact_info')
            )

        if isinstance(row, Sequence):
            last_name = row[5] if len(row) > 5 else None
            first_name = row[6] if len(row) > 6 else None
            contact_info = row[7] if len(row) > 7 else None
            role = row[8] if len(row) > 8 else 'user'
            return cls(
                user_id=row[0],
                username=row[1],
                password_hash=row[2],
                created_at=row[3],
                last_login=row[4],
                last_name=last_name,
                first_name=first_name,
                contact_info=contact_info,
                role=role
            )

        raise TypeError(f'Unsupported row type: {type(row)}')

    @property
    def is_admin(self) -> bool:
        """Проверяет, имеет ли пользователь административные привилегии."""
        return self.role == 'admin'

    def set_last_login_now(self):
        """Проставляет текущее время последнего входа (используется в UI)."""
        self.last_login = datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S')
