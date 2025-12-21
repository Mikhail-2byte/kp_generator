"""Интеграционные тесты для репозиториев."""

import pytest
from datetime import datetime, timedelta

from app.services.repositories import (
    generation_repository,
    user_repository,
    audit_log_repository,
)


class TestGenerationRepositoryIntegration:
    """Интеграционные тесты для GenerationRepository."""
    
    def test_save_and_get_history(self, app):
        """Тест сохранения и получения истории генераций."""
        with app.app_context():
            # Создаем тестового пользователя
            user = user_repository.create_user(
                username=f'test_user_{int(datetime.now().timestamp())}',
                password_hash='test_hash',
                role='user'
            )
            
            # Подготавливаем данные для сохранения
            form_data = {
                'company': 'Тестовая компания',
                'product': 'Тестовый товар',
                'quantity': '10',
                'cost_price': '1000',
                'weight': '5',
                'logistics': '50000',
                'margin_percent': '30',
                'delivery_time': '30',
                'duty_percent': '5',
                'drawing_number': 'Ч-001',
                'material': 'Сталь',
                'delivery_address': 'Москва',
                'positions': [{
                    'quantity': 10,
                    'cost_price': 1000,
                    'weight': 5,
                    'duty_percent': 5
                }]
            }
            
            # Сохраняем генерацию
            saved = generation_repository.save_history(
                form_data, 1500.0, app.config['APP_SETTINGS'], user.id, 15000.0
            )
            
            assert saved is True
            
            # Получаем историю
            history = generation_repository.get_history(
                app.config['APP_SETTINGS'], page=1, per_page=10
            )
            
            assert history['pagination']['total'] > 0
            assert len(history['items']) > 0
    
    def test_get_details(self, app):
        """Тест получения деталей генерации."""
        with app.app_context():
            # Создаем пользователя и генерацию
            user = user_repository.create_user(
                username=f'test_user_{int(datetime.now().timestamp())}',
                password_hash='test_hash'
            )
            
            form_data = {
                'company': 'Компания для деталей',
                'product': 'Товар',
                'quantity': '5',
                'cost_price': '500',
                'weight': '3',
                'logistics': '30000',
                'margin_percent': '25',
                'delivery_time': '20',
                'duty_percent': '0',
                'positions': [{'quantity': 5, 'cost_price': 500, 'weight': 3}]
            }
            
            saved = generation_repository.save_history(
                form_data, 800.0, app.config['APP_SETTINGS'], user.id, 4000.0
            )
            
            assert saved is True
            
            # Получаем последнюю генерацию
            history = generation_repository.get_history(
                app.config['APP_SETTINGS'], page=1, per_page=1
            )
            
            if history['items']:
                record_id = history['items'][0]['id']
                details = generation_repository.get_details(record_id)
                
                assert details is not None
                assert details['company'] == 'Компания для деталей'


class TestAuditLogRepositoryIntegration:
    """Интеграционные тесты для AuditLogRepository."""
    
    def test_create_and_get_logs(self, app):
        """Тест создания и получения логов."""
        with app.app_context():
            # Создаем пользователя
            user = user_repository.create_user(
                username=f'test_user_{int(datetime.now().timestamp())}',
                password_hash='test_hash'
            )
            
            # Создаем лог
            success = audit_log_repository.create_log(
                user_id=user.id,
                username=user.username,
                action_type='test',
                description='Тестовое действие',
                resource_type='test_resource',
                resource_id='1'
            )
            
            assert success is True
            
            # Получаем логи
            logs = audit_log_repository.get_logs(
                action_type='test',
                page=1,
                per_page=10
            )
            
            assert logs['pagination']['total'] > 0
            assert len(logs['items']) > 0
    
    def test_get_daily_activity(self, app):
        """Тест получения ежедневной активности."""
        with app.app_context():
            user = user_repository.create_user(
                username=f'test_user_{int(datetime.now().timestamp())}',
                password_hash='test_hash'
            )
            
            # Создаем несколько логов
            for i in range(3):
                audit_log_repository.create_log(
                    user_id=user.id,
                    username=user.username,
                    action_type='test',
                    description=f'Тест {i}'
                )
            
            # Получаем активность
            activity = audit_log_repository.get_daily_activity(days=7)
            
            assert isinstance(activity, list)
            assert len(activity) > 0


class TestUserRepositoryIntegration:
    """Интеграционные тесты для UserRepository."""
    
    def test_create_and_get_user(self, app):
        """Тест создания и получения пользователя."""
        with app.app_context():
            username = f'test_user_{int(datetime.now().timestamp())}'
            
            user = user_repository.create_user(
                username=username,
                password_hash='test_hash',
                last_name='Тестов',
                first_name='Тест',
                role='user'
            )
            
            assert user is not None
            assert user.username == username
            
            # Получаем пользователя
            retrieved = user_repository.get_by_username(username)
            assert retrieved is not None
            assert retrieved.username == username
    
    def test_update_profile(self, app):
        """Тест обновления профиля пользователя."""
        with app.app_context():
            username = f'test_user_{int(datetime.now().timestamp())}'
            
            user = user_repository.create_user(
                username=username,
                password_hash='test_hash',
                last_name='Старая',
                first_name='Фамилия'
            )
            
            # Обновляем профиль
            success = user_repository.update_profile(
                user.id,
                username=username,
                last_name='Новая',
                first_name='Фамилия',
                contact_info='test@example.com'
            )
            
            assert success is True
            
            # Проверяем обновление
            updated = user_repository.get_by_id(user.id)
            assert updated.last_name == 'Новая'
            assert updated.contact_info == 'test@example.com'

