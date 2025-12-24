"""Тесты для REST API endpoints."""

import json
from typing import Any, Dict

import pytest
from flask import Flask
from flask.testing import FlaskClient

from app.database.service import database_service


@pytest.fixture
def test_user(app: Flask) -> Dict[str, Any]:
    """Создает тестового пользователя для аутентификации."""
    user_id = database_service.create_user(
        username='api_test_user',
        password_hash='test_hash',
        last_name='Тестов',
        first_name='Тест',
        role='user',
        contact_info='test@example.com'
    )
    return {
        'id': user_id,
        'username': 'api_test_user'
    }


@pytest.fixture
def authenticated_client(client: FlaskClient, app: Flask, test_user: Dict[str, Any]) -> FlaskClient:
    """Создает аутентифицированного клиента."""
    # Логинимся через веб-интерфейс для получения сессии
    with app.app_context():
        from app.models.models import User
        user_row = database_service.get_user_by_id(test_user['id'])
        if user_row:
            user = User.from_row(user_row)
            with client.session_transaction() as sess:
                sess['_user_id'] = str(user.id)
                sess['_fresh'] = True
    
    return client


class TestAPIHealth:
    """Тесты для health check endpoint."""
    
    def test_health_endpoint(self, client: FlaskClient):
        """Тест проверки работоспособности API."""
        response = client.get('/api/v1/health')
        assert response.status_code == 200
        data = response.get_json()
        assert data['status'] == 'ok'
        assert data['version'] == '1.0'
        assert data['service'] == 'kp_generator'


class TestAPIGenerations:
    """Тесты для endpoints генераций."""
    
    def test_list_generations_requires_auth(self, client: FlaskClient):
        """Тест что список генераций требует аутентификации."""
        response = client.get('/api/v1/generations')
        assert response.status_code == 401 or response.status_code == 302  # Redirect to login
    
    def test_list_generations_with_auth(self, authenticated_client: FlaskClient, app: Flask):
        """Тест получения списка генераций с аутентификацией."""
        # Создаем тестовую генерацию
        with app.app_context():
            from app.services.repositories import generation_repository
            config = app.config.get('APP_SETTINGS', {})
            
            # Создаем генерацию через репозиторий
            gen_data = {
                'company': 'Тестовая компания',
                'product': 'Тестовый товар',
                'quantity': 10,
                'cost_price': 1000.0,
                'final_price': 1500.0,
                'margin_percent': 30.0,
                'logistics': 50000.0,
                'weight': 5.0,
                'duty_percent': 5.0,
                'delivery_time': 30,
                'tender_number': 'TEST-001'
            }
            
            # Используем прямой вызов БД для создания генерации
            from app.database.database import save_generation_history
            config = app.config.get('APP_SETTINGS', {})
            save_generation_history(gen_data, 1500.0, config, None, 15000.0)
            
            response = authenticated_client.get('/api/v1/generations?page=1&per_page=10')
            assert response.status_code == 200
            data = response.get_json()
            assert 'items' in data
            assert 'pagination' in data
    
    def test_get_generation_requires_auth(self, client: FlaskClient):
        """Тест что получение генерации требует аутентификации."""
        response = client.get('/api/v1/generations/1')
        assert response.status_code == 401 or response.status_code == 302
    
    def test_get_generation_not_found(self, authenticated_client: FlaskClient):
        """Тест получения несуществующей генерации."""
        response = authenticated_client.get('/api/v1/generations/99999')
        assert response.status_code == 404


class TestAPICalculate:
    """Тесты для endpoint расчета цен."""
    
    def test_calculate_requires_auth(self, client: FlaskClient):
        """Тест что расчет требует аутентификации."""
        response = client.post(
            '/api/v1/calculate',
            json={'positions': []},
            content_type='application/json'
        )
        assert response.status_code == 401 or response.status_code == 302
    
    def test_calculate_with_valid_data(self, authenticated_client: FlaskClient, app: Flask):
        """Тест расчета цен с валидными данными."""
        with app.app_context():
            data = {
                'positions': [
                    {
                        'quantity': 10,
                        'cost_price': 1000,
                        'weight': 5,
                        'duty_percent': 5
                    }
                ],
                'logistics_rub': 50000,
                'delivery_time': 30,
                'margin_percent': 30
            }
            
            response = authenticated_client.post(
                '/api/v1/calculate',
                json=data,
                content_type='application/json'
            )
            
            assert response.status_code == 200
            result = response.get_json()
            assert 'positions' in result
            assert 'total_costs' in result
            assert 'total_revenue' in result
            assert 'target_margin' in result
            assert 'actual_margin' in result
    
    def test_calculate_without_positions(self, authenticated_client: FlaskClient):
        """Тест расчета без позиций."""
        data = {
            'logistics_rub': 50000,
            'delivery_time': 30,
            'margin_percent': 30
        }
        
        response = authenticated_client.post(
            '/api/v1/calculate',
            json=data,
            content_type='application/json'
        )
        
        assert response.status_code == 400


class TestAPILogistics:
    """Тесты для endpoint расчета логистики."""
    
    def test_calculate_logistics(self, client: FlaskClient):
        """Тест расчета логистики."""
        data = {
            'weight_kg': 1000,
            'city_price': 50000,
            'transport_type': 'truck',
            'city_name': 'Москва',
            'is_main_route': True,
            'distance_from_ekb_km': 1800
        }
        
        response = client.post(
            '/api/v1/logistics/calculate',
            json=data,
            content_type='application/json'
        )
        
        assert response.status_code == 200
        result = response.get_json()
        # calculate_logistics возвращает разные структуры в зависимости от условий
        assert 'basis' in result or 'total_price' in result or 'error' in result
    
    def test_calculate_logistics_invalid_data(self, client: FlaskClient):
        """Тест расчета логистики с невалидными данными."""
        data = {
            'weight_kg': -100,  # Отрицательный вес
            'city_price': 50000
        }
        
        response = client.post(
            '/api/v1/logistics/calculate',
            json=data,
            content_type='application/json'
        )
        
        # Должен обработать или вернуть ошибку
        assert response.status_code in [200, 400]


class TestAPIDuty:
    """Тесты для endpoint поиска пошлин."""
    
    def test_search_duty(self, client: FlaskClient):
        """Тест поиска пошлин."""
        response = client.get('/api/v1/duty/search?query=сталь&limit=10')
        assert response.status_code == 200
        data = response.get_json()
        assert 'items' in data
        assert 'total' in data
        assert 'query' in data
    
    def test_search_duty_empty_query(self, client: FlaskClient):
        """Тест поиска пошлин с пустым запросом."""
        response = client.get('/api/v1/duty/search')
        assert response.status_code == 200
        data = response.get_json()
        assert 'items' in data


class TestAPIGBMaterials:
    """Тесты для endpoint материалов GB."""
    
    def test_get_gb_materials(self, client: FlaskClient):
        """Тест получения материалов GB."""
        response = client.get('/api/v1/materials/gb')
        assert response.status_code == 200
        data = response.get_json()
        assert 'items' in data
        assert 'total' in data
    
    def test_get_gb_materials_with_search(self, client: FlaskClient):
        """Тест поиска материалов GB."""
        response = client.get('/api/v1/materials/gb?search=сталь')
        assert response.status_code == 200
        data = response.get_json()
        assert 'items' in data


class TestAPILogisticsCities:
    """Тесты для endpoint городов логистики."""
    
    def test_get_logistics_cities(self, client: FlaskClient):
        """Тест получения городов логистики."""
        response = client.get('/api/v1/logistics/cities')
        assert response.status_code == 200
        data = response.get_json()
        assert 'items' in data
        assert 'total' in data


class TestAPICustomers:
    """Тесты для endpoints контактов заказчиков."""
    
    def test_list_customers_requires_auth(self, client: FlaskClient):
        """Тест что список контактов требует аутентификации."""
        response = client.get('/api/v1/customers')
        assert response.status_code == 401 or response.status_code == 302
    
    def test_create_customer_requires_auth(self, client: FlaskClient):
        """Тест что создание контакта требует аутентификации."""
        response = client.post(
            '/api/v1/customers',
            json={'company_name': 'Тест'},
            content_type='application/json'
        )
        assert response.status_code == 401 or response.status_code == 302
    
    def test_create_customer_with_auth(self, authenticated_client: FlaskClient):
        """Тест создания контакта с аутентификацией."""
        data = {
            'company_name': 'Тестовая компания API',
            'contact_person': 'Иван Иванов',
            'phone': '+7 999 123-45-67',
            'email': 'test@example.com'
        }
        
        response = authenticated_client.post(
            '/api/v1/customers',
            json=data,
            content_type='application/json'
        )
        
        assert response.status_code == 201
        result = response.get_json()
        assert result['company_name'] == 'Тестовая компания API'
        assert 'id' in result
    
    def test_create_customer_without_name(self, authenticated_client: FlaskClient):
        """Тест создания контакта без названия компании."""
        data = {
            'contact_person': 'Иван Иванов'
        }
        
        response = authenticated_client.post(
            '/api/v1/customers',
            json=data,
            content_type='application/json'
        )
        
        assert response.status_code == 400
    
    def test_get_customer_requires_auth(self, client: FlaskClient):
        """Тест что получение контакта требует аутентификации."""
        response = client.get('/api/v1/customers/1')
        assert response.status_code == 401 or response.status_code == 302
    
    def test_get_customer_not_found(self, authenticated_client: FlaskClient):
        """Тест получения несуществующего контакта."""
        response = authenticated_client.get('/api/v1/customers/99999')
        assert response.status_code == 404
    
    def test_update_customer_requires_auth(self, client: FlaskClient):
        """Тест что обновление контакта требует аутентификации."""
        response = client.put(
            '/api/v1/customers/1',
            json={'phone': '+7 999 999-99-99'},
            content_type='application/json'
        )
        assert response.status_code == 401 or response.status_code == 302
    
    def test_delete_customer_requires_auth(self, client: FlaskClient):
        """Тест что удаление контакта требует аутентификации."""
        response = client.delete('/api/v1/customers/1')
        assert response.status_code == 401 or response.status_code == 302

