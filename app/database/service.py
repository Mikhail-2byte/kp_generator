from __future__ import annotations

from typing import Any, Dict, Optional

from .database import (
    connect_db,
    init_db,
    get_generation_history,
    save_generation_history,
    get_generation_details,
    load_generation_data,
    create_user,
    get_user_by_username,
    get_user_by_id,
    update_last_login,
    get_user_statistics,
    get_admin_user_activity,
    update_user_profile,
    delete_user,
    get_generations_by_drawing,
    get_generations_by_tender,
    get_unique_companies,
    get_users_list,
)


class DatabaseService:
    """Фасад поверх функций модуля database для удобства внедрения."""

    def connect(self):
        return connect_db()

    def init(self) -> None:
        init_db()

    # --- Пользователи -----------------------------------------------------------------

    def create_user(
        self,
        username: str,
        password_hash: str,
        last_name: Optional[str] = None,
        first_name: Optional[str] = None,
        role: str = 'user',
        contact_info: Optional[str] = None,
    ) -> Optional[int]:
        return create_user(username, password_hash, last_name, first_name, role, contact_info)

    def get_user_by_username(self, username: str):
        return get_user_by_username(username)

    def get_user_by_id(self, user_id: int):
        return get_user_by_id(user_id)

    def update_last_login(self, user_id: int) -> bool:
        return update_last_login(user_id)

    def update_user_profile(
        self,
        user_id: int,
        username: str,
        last_name: Optional[str],
        first_name: Optional[str],
        contact_info: Optional[str],
        password_hash: Optional[str],
    ) -> bool:
        return update_user_profile(user_id, username, last_name, first_name, contact_info, password_hash)

    def delete_user(self, user_id: int) -> bool:
        return delete_user(user_id)

    def get_user_statistics(self, user_id: int) -> Dict[str, Any]:
        return get_user_statistics(user_id)

    def get_users_list(
        self,
        *,
        page: int = 1,
        per_page: int = 25,
        search: Optional[str] = None,
        role_filter: Optional[str] = None,
    ) -> Dict[str, Any]:
        return get_users_list(page=page, per_page=per_page, search=search, role_filter=role_filter)

    # --- История генераций -------------------------------------------------------------

    def save_generation_history(
        self,
        payload: Dict[str, Any],
        final_price: float,
        config: Dict[str, Any],
        user_id: Optional[int],
        total_general_price: Optional[float] = None,
    ) -> bool:
        return save_generation_history(payload, final_price, config, user_id, total_general_price)

    def get_generation_history(
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
        return get_generation_history(
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

    def get_generation_details(self, record_id: int):
        return get_generation_details(record_id)

    def load_generation_data(self, record_id: int):
        return load_generation_data(record_id)

    def get_generations_by_drawing(self, drawing_number: str):
        return get_generations_by_drawing(drawing_number)

    def get_generations_by_tender(self, tender_number: str):
        return get_generations_by_tender(tender_number)

    def get_unique_companies(self):
        return get_unique_companies()

    # --- Админская аналитика ----------------------------------------------------------

    def get_admin_user_activity(self, limit: int = 10):
        return get_admin_user_activity(limit)


database_service = DatabaseService()

__all__ = ['DatabaseService', 'database_service']
