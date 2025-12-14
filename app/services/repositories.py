from __future__ import annotations

from typing import Any, Dict, Optional

from app.database import DatabaseService, database_service
from app.models.models import User


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

    def save_history(self, payload: Dict[str, Any], final_price: float, config: Dict[str, Any], user_id: Any) -> bool:
        """Сохраняет расчёт генерации с привязкой к пользователю."""
        return self._db.save_generation_history(payload, final_price, config, int(user_id) if user_id is not None else None)

    def get_history(self, config: Dict[str, Any], *, page: int = 1, per_page: Optional[int] = None, date_from: Optional[str] = None, date_to: Optional[str] = None):
        """Возвращает историю генераций с пагинацией."""
        return self._db.get_generation_history(config, page=page, per_page=per_page, date_from=date_from, date_to=date_to)

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


user_repository = UserRepository()
generation_repository = GenerationRepository()
admin_stats_repository = AdminStatsRepository()

__all__ = [
    'user_repository',
    'generation_repository',
    'admin_stats_repository',
    'UserRepository',
    'GenerationRepository',
    'AdminStatsRepository',
]
