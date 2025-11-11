from __future__ import annotations

from typing import Any, Dict, Optional

from app.database import (
    create_user,
    get_user_by_username,
    get_user_by_id,
    update_last_login,
    get_user_statistics,
    update_user_profile,
    delete_user,
    save_generation_history,
    get_generation_history,
    get_generation_details,
    load_generation_data,
    get_generations_by_drawing,
    get_admin_user_activity,
)
from app.models.models import User


class UserRepository:
    """Инкапсулирует операции с пользователями поверх слоя базы данных."""
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
        user_id = create_user(username, password_hash, last_name, first_name, role, contact_info)
        return self.get_by_id(user_id) if user_id else None

    def get_by_username(self, username: str) -> Optional[User]:
        """Находит пользователя по логину (используется при входе)."""
        row = get_user_by_username(username)
        return User.from_row(row) if row else None

    def get_by_id(self, user_id: Any) -> Optional[User]:
        """Возвращает пользователя по идентификатору с учётом пустых значений."""
        if user_id is None:
            return None
        row = get_user_by_id(int(user_id))
        return User.from_row(row) if row else None

    def record_login(self, user_id: Any) -> bool:
        """Фиксирует время последнего входа пользователя."""
        if user_id is None:
            return False
        return update_last_login(int(user_id))

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
        return update_user_profile(
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
        return delete_user(int(user_id))

    def get_statistics(self, user_id: Any) -> Dict[str, Any]:
        """Возвращает превью активности пользователя для личного кабинета."""
        if user_id is None:
            return {}
        return get_user_statistics(int(user_id)) or {}


class GenerationRepository:
    """Сервис-обёртка для чтения и сохранения историй генераций."""
    def save_history(self, payload: Dict[str, Any], final_price: float, config: Dict[str, Any], user_id: Any) -> bool:
        """Сохраняет расчёт генерации с привязкой к пользователю."""
        return save_generation_history(payload, final_price, config, int(user_id) if user_id is not None else None)

    def get_history(self, config: Dict[str, Any]):
        """Возвращает ограниченную историю генераций согласно конфигурации."""
        return get_generation_history(config)

    def get_details(self, record_id: int):
        """Загружает подробности конкретной генерации по идентификатору."""
        return get_generation_details(record_id)

    def load_generation(self, record_id: int):
        """Возвращает данные генерации для повторного использования в форме."""
        return load_generation_data(record_id)

    def get_by_drawing(self, drawing_number: str):
        """Возвращает список генераций, созданных с указанным номером чертежа."""
        return get_generations_by_drawing(drawing_number)


class AdminStatsRepository:
    """Предоставляет агрегированные данные для административной панели."""

    def get_user_activity(self, limit: int = 10):
        """Возвращает статистику пользовательской активности."""
        return get_admin_user_activity(limit)


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
