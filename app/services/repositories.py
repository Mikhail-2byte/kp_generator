from __future__ import annotations

import json
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional

from sqlalchemy import func, and_, or_

from app.database import DatabaseService, database_service
from app.database.database import _session_scope
from app.models.models import AuditLogRecord, CustomerContactRecord, User


class UserRepository:
    """Инкапсулирует операции с пользователями поверх слоя базы данных."""

    def __init__(self, db: DatabaseService = database_service) -> None:
        self._db = db

    def create_user(
        self,
        username: str,
        password_hash: str,
        last_name: Optional[str] = None,
        first_name: Optional[str] = None,
        role: str = 'user',
        contact_info: Optional[str] = None
    ) -> Optional[User]:
        """Создаёт пользователя и возвращает готовый объект для работы в приложении."""
        user_id = self._db.create_user(username, password_hash, last_name, first_name, role, contact_info)
        return self.get_by_id(user_id) if user_id else None

    def get_by_username(self, username: str) -> Optional[User]:
        """Находит пользователя по логину (используется при входе)."""
        row = self._db.get_user_by_username(username)
        return User.from_row(row) if row else None

    def get_by_id(self, user_id: Any) -> Optional[User]:
        """Возвращает пользователя по идентификатору с учётом пустых значений."""
        if user_id is None:
            return None
        row = self._db.get_user_by_id(int(user_id))
        return User.from_row(row) if row else None

    def record_login(self, user_id: Any) -> bool:
        """Фиксирует время последнего входа пользователя."""
        if user_id is None:
            return False
        return self._db.update_last_login(int(user_id))

    def update_profile(
        self,
        user_id: Any,
        username: str,
        last_name: Optional[str] = None,
        first_name: Optional[str] = None,
        contact_info: Optional[str] = None,
        password_hash: Optional[str] = None
    ) -> bool:
        """Обновляет логин, имя и пароль пользователя."""
        if user_id is None:
            return False
        return self._db.update_user_profile(
            int(user_id),
            username,
            last_name,
            first_name,
            contact_info,
            password_hash
        )

    def delete(self, user_id: Any) -> bool:
        """Удаляет пользователя и связанные записи истории."""
        if user_id is None:
            return False
        return self._db.delete_user(int(user_id))

    def get_statistics(self, user_id: Any) -> Dict[str, Any]:
        """Возвращает превью активности пользователя для личного кабинета."""
        if user_id is None:
            return {}
        return self._db.get_user_statistics(int(user_id)) or {}

    def get_users_list(
        self,
        *,
        page: int = 1,
        per_page: int = 25,
        search: Optional[str] = None,
        role_filter: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Возвращает список пользователей с пагинацией, поиском и фильтрацией."""
        return self._db.get_users_list(page=page, per_page=per_page, search=search, role_filter=role_filter)


class GenerationRepository:
    """Сервис-обёртка для чтения и сохранения историй генераций."""

    def __init__(self, db: DatabaseService = database_service) -> None:
        self._db = db

    def save_history(self, payload: Dict[str, Any], final_price: float, config: Dict[str, Any], user_id: Any, total_general_price: Optional[float] = None) -> bool:
        """Сохраняет расчёт генерации с привязкой к пользователю."""
        return self._db.save_generation_history(payload, final_price, config, int(user_id) if user_id is not None else None, total_general_price)

    def get_history(
        self,
        config: Dict[str, Any],
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
    ):
        """Возвращает историю генераций с пагинацией, серверной фильтрацией и сортировкой."""
        return self._db.get_generation_history(
            config,
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
            sort_by=sort_by,
            sort_order=sort_order,
        )

    def get_details(self, record_id: int):
        """Загружает подробности конкретной генерации по идентификатору."""
        return self._db.get_generation_details(record_id)

    def load_generation(self, record_id: int):
        """Возвращает данные генерации для повторного использования в форме."""
        return self._db.load_generation_data(record_id)

    def get_by_drawing(self, drawing_number: str):
        """Возвращает список генераций, созданных с указанным номером чертежа."""
        return self._db.get_generations_by_drawing(drawing_number)

    def get_by_tender(self, tender_number: str):
        """Возвращает список генераций по номеру тендера."""
        return self._db.get_generations_by_tender(tender_number)

    def get_unique_companies(self):
        """Возвращает список уникальных компаний."""
        return self._db.get_unique_companies()


class AdminStatsRepository:
    """Предоставляет агрегированные данные для административной панели."""

    def __init__(self, db: DatabaseService = database_service) -> None:
        self._db = db

    def get_user_activity(self, limit: int = 10):
        """Возвращает статистику пользовательской активности."""
        return self._db.get_admin_user_activity(limit)


class AuditLogRepository:
    """Репозиторий для работы с логами аудита действий пользователей."""

    def create_log(
        self,
        user_id: Optional[int],
        username: str,
        action_type: str,
        description: str,
        resource_type: Optional[str] = None,
        resource_id: Optional[str] = None,
        ip_address: Optional[str] = None,
        user_agent: Optional[str] = None,
        changes_before: Optional[Dict[str, Any]] = None,
        changes_after: Optional[Dict[str, Any]] = None,
    ) -> bool:
        """Создает запись в логе аудита."""
        try:
            with _session_scope() as session:
                changes_before_json = json.dumps(changes_before, ensure_ascii=False) if changes_before else None
                changes_after_json = json.dumps(changes_after, ensure_ascii=False) if changes_after else None

                log_record = AuditLogRecord(
                    user_id=user_id,
                    username=username,
                    action_type=action_type,
                    resource_type=resource_type,
                    resource_id=resource_id,
                    description=description,
                    ip_address=ip_address,
                    user_agent=user_agent,
                    changes_before=changes_before_json,
                    changes_after=changes_after_json,
                )
                session.add(log_record)
                return True
        except Exception:
            return False

    def get_logs(
        self,
        *,
        page: int = 1,
        per_page: int = 50,
        user_id: Optional[int] = None,
        username: Optional[str] = None,
        action_type: Optional[str] = None,
        resource_type: Optional[str] = None,
        date_from: Optional[datetime] = None,
        date_to: Optional[datetime] = None,
        search: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Возвращает список логов с пагинацией и фильтрами."""
        with _session_scope() as session:
            query = session.query(AuditLogRecord)

            # Применяем фильтры
            if user_id is not None:
                query = query.filter(AuditLogRecord.user_id == user_id)
            if username:
                query = query.filter(AuditLogRecord.username.ilike(f'%{username}%'))
            if action_type:
                query = query.filter(AuditLogRecord.action_type == action_type)
            if resource_type:
                query = query.filter(AuditLogRecord.resource_type == resource_type)
            if date_from:
                query = query.filter(AuditLogRecord.created_at >= date_from)
            if date_to:
                query = query.filter(AuditLogRecord.created_at <= date_to)
            if search:
                search_pattern = f'%{search}%'
                query = query.filter(
                    or_(
                        AuditLogRecord.description.ilike(search_pattern),
                        AuditLogRecord.username.ilike(search_pattern),
                    )
                )

            # Подсчет общего количества
            total = query.count()

            # Пагинация
            offset = (page - 1) * per_page
            logs = query.order_by(AuditLogRecord.created_at.desc()).offset(offset).limit(per_page).all()

            # Форматирование результатов
            items = []
            for log in logs:
                changes_before_dict = None
                changes_after_dict = None
                try:
                    if log.changes_before:
                        changes_before_dict = json.loads(log.changes_before)
                    if log.changes_after:
                        changes_after_dict = json.loads(log.changes_after)
                except (json.JSONDecodeError, TypeError):
                    pass

                items.append({
                    'id': log.id,
                    'user_id': log.user_id,
                    'username': log.username,
                    'action_type': log.action_type,
                    'resource_type': log.resource_type,
                    'resource_id': log.resource_id,
                    'description': log.description,
                    'ip_address': log.ip_address,
                    'user_agent': log.user_agent,
                    'changes_before': changes_before_dict,
                    'changes_after': changes_after_dict,
                    'created_at': (
                        log.created_at.strftime('%Y-%m-%d %H:%M:%S')
                        if log.created_at and hasattr(log.created_at, 'strftime')
                        else (str(log.created_at) if log.created_at else None)
                    ),
                })

            total_pages = (total + per_page - 1) // per_page if per_page > 0 else 1

            return {
                'items': items,
                'pagination': {
                    'page': page,
                    'per_page': per_page,
                    'total': total,
                    'pages': max(total_pages, 1),
                },
            }

    def get_daily_activity(self, days: int = 30) -> List[Dict[str, Any]]:
        """Возвращает активность по дням за указанный период."""
        with _session_scope() as session:
            date_from = datetime.now() - timedelta(days=days)
            query = (
                session.query(
                    func.date(AuditLogRecord.created_at).label('date'),
                    func.count(AuditLogRecord.id).label('count'),
                )
                .filter(AuditLogRecord.created_at >= date_from)
                .group_by(func.date(AuditLogRecord.created_at))
                .order_by(func.date(AuditLogRecord.created_at))
            )
            results = query.all()

            return [
                {
                    'date': str(row.date) if row.date else None,
                    'count': row.count,
                }
                for row in results
            ]

    def get_action_stats(self, days: int = 30) -> List[Dict[str, Any]]:
        """Возвращает статистику по типам действий."""
        with _session_scope() as session:
            date_from = datetime.now() - timedelta(days=days)
            query = (
                session.query(
                    AuditLogRecord.action_type,
                    func.count(AuditLogRecord.id).label('count'),
                )
                .filter(AuditLogRecord.created_at >= date_from)
                .group_by(AuditLogRecord.action_type)
                .order_by(func.count(AuditLogRecord.id).desc())
            )
            results = query.all()

            return [
                {
                    'action_type': row.action_type,
                    'count': row.count,
                }
                for row in results
            ]

    def get_top_users(self, limit: int = 10, days: int = 30) -> List[Dict[str, Any]]:
        """Возвращает топ активных пользователей."""
        with _session_scope() as session:
            date_from = datetime.now() - timedelta(days=days)
            query = (
                session.query(
                    AuditLogRecord.username,
                    func.count(AuditLogRecord.id).label('count'),
                )
                .filter(AuditLogRecord.created_at >= date_from)
                .group_by(AuditLogRecord.username)
                .order_by(func.count(AuditLogRecord.id).desc())
                .limit(limit)
            )
            results = query.all()

            return [
                {
                    'username': row.username,
                    'count': row.count,
                }
                for row in results
            ]

    def get_popular_actions(self, limit: int = 10, days: int = 30) -> List[Dict[str, Any]]:
        """Возвращает популярные операции."""
        return self.get_action_stats(days=days)[:limit]

    def get_user_activity(self, user_id: int, days: int = 30) -> Dict[str, Any]:
        """Возвращает активность конкретного пользователя."""
        date_from = datetime.now() - timedelta(days=days)
        logs_data = self.get_logs(
            user_id=user_id,
            date_from=date_from,
            page=1,
            per_page=1000,  # Большое значение для получения всех записей
        )

        # Группировка по типам действий
        action_counts = {}
        for item in logs_data['items']:
            action_type = item['action_type']
            action_counts[action_type] = action_counts.get(action_type, 0) + 1

        return {
            'total_actions': logs_data['pagination']['total'],
            'action_breakdown': action_counts,
            'recent_logs': logs_data['items'][:10],  # Последние 10 действий
        }


class CustomerContactRepository:
    """Репозиторий для работы с контактами заказчиков."""
    
    def create(
        self,
        company_name: str,
        contact_person: Optional[str] = None,
        phone: Optional[str] = None,
        email: Optional[str] = None,
        address: Optional[str] = None,
        notes: Optional[str] = None
    ) -> Optional[int]:
        """Создает новый контакт заказчика."""
        with _session_scope() as session:
            contact = CustomerContactRecord(
                company_name=company_name,
                contact_person=contact_person,
                phone=phone,
                email=email,
                address=address,
                notes=notes
            )
            session.add(contact)
            session.flush()
            return contact.id
    
    def get_by_id(self, contact_id: int) -> Optional[Dict[str, Any]]:
        """Получает контакт по ID."""
        with _session_scope() as session:
            contact = session.query(CustomerContactRecord).filter(
                CustomerContactRecord.id == contact_id
            ).first()
            if not contact:
                return None
            return {
                'id': contact.id,
                'company_name': contact.company_name,
                'contact_person': contact.contact_person,
                'phone': contact.phone,
                'email': contact.email,
                'address': contact.address,
                'notes': contact.notes,
                'created_at': contact.created_at.isoformat() if contact.created_at else None,
                'updated_at': contact.updated_at.isoformat() if contact.updated_at else None,
            }
    
    def get_all(self, search: Optional[str] = None) -> List[Dict[str, Any]]:
        """Получает все контакты с опциональным поиском."""
        with _session_scope() as session:
            query = session.query(CustomerContactRecord)
            if search:
                # ilike уже case-insensitive, но для SQLite может потребоваться явное приведение
                search_term = f'%{search}%'
                query = query.filter(
                    or_(
                        CustomerContactRecord.company_name.ilike(search_term),
                        CustomerContactRecord.contact_person.ilike(search_term),
                        CustomerContactRecord.phone.ilike(search_term),
                        CustomerContactRecord.email.ilike(search_term),
                    )
                )
            contacts = query.order_by(CustomerContactRecord.company_name).all()
            return [
                {
                    'id': c.id,
                    'company_name': c.company_name,
                    'contact_person': c.contact_person,
                    'phone': c.phone,
                    'email': c.email,
                    'address': c.address,
                    'notes': c.notes,
                    'created_at': c.created_at.isoformat() if c.created_at else None,
                    'updated_at': c.updated_at.isoformat() if c.updated_at else None,
                }
                for c in contacts
            ]
    
    def update(
        self,
        contact_id: int,
        company_name: Optional[str] = None,
        contact_person: Optional[str] = None,
        phone: Optional[str] = None,
        email: Optional[str] = None,
        address: Optional[str] = None,
        notes: Optional[str] = None
    ) -> bool:
        """Обновляет контакт заказчика."""
        with _session_scope() as session:
            contact = session.query(CustomerContactRecord).filter(
                CustomerContactRecord.id == contact_id
            ).first()
            if not contact:
                return False
            
            if company_name is not None:
                contact.company_name = company_name
            if contact_person is not None:
                contact.contact_person = contact_person
            if phone is not None:
                contact.phone = phone
            if email is not None:
                contact.email = email
            if address is not None:
                contact.address = address
            if notes is not None:
                contact.notes = notes
            
            return True
    
    def delete(self, contact_id: int) -> bool:
        """Удаляет контакт заказчика."""
        with _session_scope() as session:
            contact = session.query(CustomerContactRecord).filter(
                CustomerContactRecord.id == contact_id
            ).first()
            if not contact:
                return False
            session.delete(contact)
            return True


user_repository = UserRepository()
generation_repository = GenerationRepository()
admin_stats_repository = AdminStatsRepository()
audit_log_repository = AuditLogRepository()
customer_contact_repository = CustomerContactRepository()

__all__ = [
    'user_repository',
    'generation_repository',
    'admin_stats_repository',
    'audit_log_repository',
    'customer_contact_repository',
    'UserRepository',
    'GenerationRepository',
    'AdminStatsRepository',
    'AuditLogRepository',
    'CustomerContactRepository',
]
