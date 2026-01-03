"""
Интеграционные тесты для API endpoints логистики.

Проверяет работу API endpoints:
- /api/logistics/calculate
- /api/logistics/save-distance
"""

import sys
from pathlib import Path

# Добавляем корень проекта в путь для возможности прямого запуска
project_root = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(project_root))

import pytest
import json
from flask import url_for


class TestLogisticsCalculateAPI:
    """Тесты для API расчета логистики."""
    
    def test_calculate_basic_request(self, client):
        """Базовый тест расчета логистики через API."""
        response = client.post(
            '/api/logistics/calculate',
            json={
                'weight_kg': 10000,
                'city_price': 600000,
                'transport_type': 'truck'
            },
            content_type='application/json'
        )
        
        assert response.status_code == 200
        data = response.get_json()
        assert 'total_price' in data
        assert data['basis'] == 'weight'
        assert data['total_price'] == 300000  # (600000/20000) * 10000
    
    def test_calculate_with_city_name_and_main_route(self, client):
        """Тест расчета с указанием города и флага основного маршрута."""
        response = client.post(
            '/api/logistics/calculate',
            json={
                'weight_kg': 8000,
                'city_price': 800000,
                'transport_type': 'truck',
                'city_name': 'Красноярск',
                'is_main_route': True
            },
            content_type='application/json'
        )
        
        assert response.status_code == 200
        data = response.get_json()
        assert data['calculation_type'] == 'direct'
        assert data['total_price'] == 320000  # (800000/20000) * 8000
    
    def test_calculate_ekb_plus_rf_route(self, client):
        """Тест расчета по алгоритму 'До ЕКБ + по РФ'."""
        response = client.post(
            '/api/logistics/calculate',
            json={
                'weight_kg': 3000,
                'city_price': 500000,  # не используется для алгоритма ЕКБ+РФ
                'transport_type': 'truck',
                'city_name': 'Сызрань',
                'is_main_route': False,
                'distance_from_ekb_km': 1050
            },
            content_type='application/json'
        )
        
        assert response.status_code == 200
        data = response.get_json()
        assert data['calculation_type'] == 'ekb_plus_rf'
        assert 'route_details' in data
        assert data['route_details']['china_to_ekb']['price'] == 165000
        assert abs(data['route_details']['ekb_to_destination']['price'] - 72450) < 1
        assert abs(data['total_price'] - 237450) < 1
    
    def test_calculate_needs_distance_error(self, client):
        """Тест ошибки при отсутствии расстояния для легкого груза."""
        response = client.post(
            '/api/logistics/calculate',
            json={
                'weight_kg': 3000,
                'city_price': 500000,
                'transport_type': 'truck',
                'city_name': 'Сызрань',
                'is_main_route': False
                # distance_from_ekb_km не указано
            },
            content_type='application/json'
        )
        
        assert response.status_code == 200
        data = response.get_json()
        assert 'error' in data
        assert data['needs_distance'] is True
        assert 'расстояние' in data['error'].lower()
    
    def test_calculate_trail_restriction(self, client):
        """Тест ограничения трала для недоступных городов."""
        response = client.post(
            '/api/logistics/calculate',
            json={
                'weight_kg': 25000,
                'city_price': 900000,
                'transport_type': 'trail',
                'city_name': 'Новосибирск'
            },
            content_type='application/json'
        )
        
        assert response.status_code == 200
        data = response.get_json()
        assert 'error' in data
        assert data['trail_not_available'] is True
    
    def test_calculate_trail_allowed_for_ekaterinburg(self, client):
        """Тест что трал доступен для Екатеринбурга."""
        response = client.post(
            '/api/logistics/calculate',
            json={
                'weight_kg': 30000,
                'city_price': 1100000,
                'transport_type': 'trail',
                'city_name': 'Екатеринбург'
            },
            content_type='application/json'
        )
        
        assert response.status_code == 200
        data = response.get_json()
        assert 'error' not in data
        assert data['basis'] == 'weight'
    
    def test_calculate_small_cargo_recommendation(self, client):
        """Тест рекомендации для мелкого груза."""
        response = client.post(
            '/api/logistics/calculate',
            json={
                'weight_kg': 500,
                'city_price': 400000,
                'transport_type': 'truck'
            },
            content_type='application/json'
        )
        
        assert response.status_code == 200
        data = response.get_json()
        assert data['recommendation'] == 'dellin'
        assert 'Деловые линии' in data['message']
    
    def test_calculate_with_dimensions(self, client):
        """Тест расчета с габаритами."""
        response = client.post(
            '/api/logistics/calculate',
            json={
                'weight_kg': 5000,
                'city_price': 820000,
                'transport_type': 'truck',
                'city_name': 'Новосибирск',
                'is_main_route': True,
                'length_mm': 5000,
                'width_mm': 2000,
                'height_mm': 2000
            },
            content_type='application/json'
        )
        
        assert response.status_code == 200
        data = response.get_json()
        assert 'dimensions' in data
        assert data['volume_m3'] == 20.0
        assert data['basis'] == 'weight'  # вес превалирует
    
    def test_calculate_invalid_weight(self, client):
        """Тест с некорректным весом."""
        response = client.post(
            '/api/logistics/calculate',
            json={
                'weight_kg': -100,
                'city_price': 400000,
                'transport_type': 'truck'
            },
            content_type='application/json'
        )
        
        assert response.status_code == 400
        data = response.get_json()
        assert 'error' in data
    
    def test_calculate_missing_parameters(self, client):
        """Тест с отсутствующими обязательными параметрами."""
        response = client.post(
            '/api/logistics/calculate',
            json={
                'transport_type': 'truck'
            },
            content_type='application/json'
        )
        
        assert response.status_code == 400
        data = response.get_json()
        assert 'error' in data


class TestDatasetsFunctions:
    """Тесты для функций работы с данными."""
    
    def test_logistics_cities_have_required_fields(self, app):
        """Тест что все города имеют необходимые поля."""
        from app.services import datasets
        
        with app.app_context():
            cities = datasets.load_logistics_cities()
            if not cities:
                pytest.skip("No cities in database")
            
            required_fields = ['name', 'region', 'truck_price', 'trail_price', 
                             'is_main_route', 'allows_trail', 'distance_from_ekb_km']
            
            for city in cities:
                for field in required_fields:
                    assert field in city, f"City {city.get('name')} missing field {field}"


class TestLogisticsPageAccess:
    """Тесты доступа к странице логистики."""
    
    def test_logistics_page_loads(self, client):
        """Тест загрузки страницы логистики."""
        response = client.get('/logistics')
        assert response.status_code == 200
        assert 'Калькулятор логистики'.encode('utf-8') in response.data
    
    def test_logistics_page_has_cities_data(self, client):
        """Тест что страница содержит данные о городах."""
        response = client.get('/logistics')
        assert response.status_code == 200
        # Проверяем наличие скрытого элемента с данными
        assert b'cities-data' in response.data


class TestEdgeCases:
    """Тесты граничных случаев."""
    
    def test_calculate_exactly_5_tons_main_route(self, client):
        """Тест расчета для груза ровно 5 тонн на основном маршруте."""
        response = client.post(
            '/api/logistics/calculate',
            json={
                'weight_kg': 5000,
                'city_price': 600000,
                'transport_type': 'truck',
                'is_main_route': True
            },
            content_type='application/json'
        )
        
        assert response.status_code == 200
        data = response.get_json()
        # При 5 тоннах на основном маршруте - прямой расчет
        assert data['calculation_type'] == 'direct'
    
    def test_calculate_exactly_5_tons_not_main_route_with_distance(self, client):
        """Тест расчета для груза ровно 5 тонн не на основном маршруте с расстоянием."""
        response = client.post(
            '/api/logistics/calculate',
            json={
                'weight_kg': 5000,
                'city_price': 600000,
                'transport_type': 'truck',
                'is_main_route': False,
                'distance_from_ekb_km': 1000
            },
            content_type='application/json'
        )
        
        assert response.status_code == 200
        data = response.get_json()
        # При 5 тоннах используется прямой расчет независимо от расстояния
        assert data['calculation_type'] == 'direct'
    
    def test_calculate_exactly_600_kg(self, client):
        """Тест расчета для груза ровно 600 кг."""
        response = client.post(
            '/api/logistics/calculate',
            json={
                'weight_kg': 600,
                'city_price': 400000,
                'transport_type': 'truck'
            },
            content_type='application/json'
        )
        
        assert response.status_code == 200
        data = response.get_json()
        # 600 кг - граница, должна быть рекомендация Деловых линий
        assert data['recommendation'] == 'dellin'
    
    def test_calculate_601_kg(self, client):
        """Тест расчета для груза 601 кг (чуть больше порога)."""
        response = client.post(
            '/api/logistics/calculate',
            json={
                'weight_kg': 601,
                'city_price': 400000,
                'transport_type': 'truck'
            },
            content_type='application/json'
        )
        
        assert response.status_code == 200
        data = response.get_json()
        # 601 кг - нормальный расчет
        assert 'recommendation' not in data or data['recommendation'] != 'dellin'
        assert 'total_price' in data


if __name__ == '__main__':
    pytest.main([__file__, '-v'])

