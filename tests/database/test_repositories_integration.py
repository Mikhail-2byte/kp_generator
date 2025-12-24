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
            
            assert saved is not None
            assert isinstance(saved, int)
            
            # Получаем сохраненную историю
            history = generation_repository.get_by_id(saved)
            
            assert history is not None
            assert history['id'] == saved
            assert history['user_id'] == user.id
            assert history['company'] == 'Тестовая компания'
            assert history['final_price'] == 1500.0
            assert history['general_price'] == 15000.0
    
    def test_get_history_list(self, app):
        """Тест получения списка истории генераций."""
        with app.app_context():
            # Создаем пользователя
            user = user_repository.create_user(
                username=f'test_user_list_{int(datetime.now().timestamp())}',
                password_hash='test_hash',
                role='user'
            )
            
            # Сохраняем несколько генераций
            for i in range(3):
                form_data = {
                    'company': f'Компания {i}',
                    'product': f'Товар {i}',
                    'quantity': '10',
                    'cost_price': '1000',
                    'weight': '5',
                    'logistics': '50000',
                    'margin_percent': '30',
                    'delivery_time': '30',
                    'duty_percent': '5',
                }
                generation_repository.save_history(
                    form_data, 1500.0, app.config['APP_SETTINGS'], user.id, 15000.0
                )
            
            # Получаем список
            history_list = generation_repository.get_user_history(user.id, limit=10, offset=0)
            
            assert len(history_list) >= 3
            assert all(item['user_id'] == user.id for item in history_list)
    
    def test_get_history_with_filters(self, app):
        """Тест получения истории с фильтрами."""
        with app.app_context():
            user = user_repository.create_user(
                username=f'test_user_filters_{int(datetime.now().timestamp())}',
                password_hash='test_hash',
                role='user'
            )
            
            form_data = {
                'company': 'Фильтрованная компания',
                'product': 'Товар',
                'quantity': '10',
                'cost_price': '1000',
                'weight': '5',
                'logistics': '50000',
                'margin_percent': '30',
                'delivery_time': '30',
                'duty_percent': '5',
            }
            
            generation_repository.save_history(
                form_data, 1500.0, app.config['APP_SETTINGS'], user.id, 15000.0
            )
            
            # Фильтруем по компании
            filtered = generation_repository.get_user_history(
                user.id, 
                company_filter='Фильтрованная',
                limit=10,
                offset=0
            )
            
            assert len(filtered) >= 1
            assert any('Фильтрованная' in item['company'] for item in filtered)


class TestAuditLogRepositoryIntegration:
    """Интеграционные тесты для AuditLogRepository."""
    
    def test_create_audit_log(self, app):
        """Тест создания записи аудита."""
        with app.app_context():
            user = user_repository.create_user(
                username=f'test_audit_user_{int(datetime.now().timestamp())}',
                password_hash='test_hash',
                role='user'
            )
            
            log_id = audit_log_repository.create(
                user_id=user.id,
                action='test_action',
                resource_type='test_resource',
                resource_id=123,
                details={'key': 'value'}
            )
            
            assert log_id is not None
            assert isinstance(log_id, int)
            
            # Получаем запись
            log = audit_log_repository.get_by_id(log_id)
            
            assert log is not None
            assert log['user_id'] == user.id
            assert log['action'] == 'test_action'
            assert log['resource_type'] == 'test_resource'
            assert log['resource_id'] == 123
    
    def test_get_user_audit_logs(self, app):
        """Тест получения логов пользователя."""
        with app.app_context():
            user = user_repository.create_user(
                username=f'test_audit_user_list_{int(datetime.now().timestamp())}',
                password_hash='test_hash',
                role='user'
            )
            
            # Создаем несколько записей
            for i in range(3):
                audit_log_repository.create(
                    user_id=user.id,
                    action=f'action_{i}',
                    resource_type='test_resource',
                    resource_id=i
                )
            
            # Получаем логи пользователя
            logs = audit_log_repository.get_user_logs(user.id, limit=10, offset=0)
            
            assert len(logs) >= 3
            assert all(log['user_id'] == user.id for log in logs)


class TestUserRepositoryIntegration:
    """Интеграционные тесты для UserRepository."""
    
    def test_create_and_get_user(self, app):
        """Тест создания и получения пользователя."""
        with app.app_context():
            user = user_repository.create_user(
                username=f'test_repo_user_{int(datetime.now().timestamp())}',
                password_hash='test_hash',
                role='user'
            )
            
            assert user is not None
            assert user.id is not None
            
            # Получаем пользователя по ID
            retrieved = user_repository.get_by_id(user.id)
            
            assert retrieved is not None
            assert retrieved.id == user.id
            assert retrieved.username == user.username
    
    def test_get_user_by_username(self, app):
        """Тест получения пользователя по имени."""
        with app.app_context():
            username = f'test_username_{int(datetime.now().timestamp())}'
            user = user_repository.create_user(
                username=username,
                password_hash='test_hash',
                role='user'
            )
            
            retrieved = user_repository.get_by_username(username)
            
            assert retrieved is not None
            assert retrieved.id == user.id
            assert retrieved.username == username
    
    def test_update_user(self, app):
        """Тест обновления пользователя."""
        with app.app_context():
            user = user_repository.create_user(
                username=f'test_update_user_{int(datetime.now().timestamp())}',
                password_hash='test_hash',
                role='user'
            )
            
            updated = user_repository.update_user(
                user.id,
                last_name='Иванов',
                first_name='Иван'
            )
            
            assert updated is True
            
            retrieved = user_repository.get_by_id(user.id)
            assert retrieved.last_name == 'Иванов'
            assert retrieved.first_name == 'Иван'

