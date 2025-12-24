"""
Тесты для административной панели логистики.

Проверяет работу админки для управления городами с новыми полями.
"""

import pytest
from flask import url_for


class TestLogisticsAdminPanel:
    """Тесты административной панели логистики."""
    
    def test_admin_logistics_page_requires_auth(self, client):
        """Тест что страница админки требует авторизации."""
        response = client.get('/admin/logistics')
        # Должен быть редирект на страницу логина
        assert response.status_code == 302


class TestNewFieldsDisplay:
    """Тесты отображения новых полей."""
    
    def test_admin_page_exists(self, client):
        """Тест что страница админки существует."""
        response = client.get('/admin/logistics', follow_redirects=False)
        # Страница должна перенаправить на логин (302) или загрузиться (200)
        assert response.status_code in [200, 302]


class TestDataMigration:
    """Тесты миграции данных."""
    
    def test_migrated_cities_have_new_fields(self, app):
        """Тест что мигрированные города имеют новые поля."""
        from app.services import datasets
        
        with app.app_context():
            cities = datasets.load_logistics_cities()
            
            if not cities:
                pytest.skip("No cities to test")
            
            # Проверяем что у всех городов есть новые поля
            for city in cities:
                assert 'is_main_route' in city
                assert 'allows_trail' in city
                assert 'distance_from_ekb_km' in city
                
                # Проверяем типы
                assert isinstance(city['is_main_route'], bool)
                assert isinstance(city['allows_trail'], bool)
                assert city['distance_from_ekb_km'] is None or isinstance(city['distance_from_ekb_km'], (int, float))
    
    def test_main_route_cities_marked_correctly(self, app):
        """Тест что основные города маршрута помечены корректно."""
        from app.services import datasets
        
        EXPECTED_MAIN_CITIES = [
            'Чита', 'Улан-Удэ', 'Иркутск', 'Красноярск', 
            'Новосибирск', 'Омск', 'Екатеринбург', 'Москва', 'Санкт-Петербург'
        ]
        
        with app.app_context():
            cities = datasets.load_logistics_cities()
            
            for expected_city in EXPECTED_MAIN_CITIES:
                city = next((c for c in cities if c['name'] == expected_city), None)
                if city:
                    assert city['is_main_route'] is True, f"{expected_city} должен быть основным маршрутом"
    
    def test_trail_cities_marked_correctly(self, app):
        """Тест что города с тралом помечены корректно."""
        from app.services import datasets
        
        EXPECTED_TRAIL_CITIES = ['Екатеринбург', 'Москва']
        
        with app.app_context():
            cities = datasets.load_logistics_cities()
            
            for city in cities:
                if city['name'] in EXPECTED_TRAIL_CITIES:
                    assert city['allows_trail'] is True, f"{city['name']} должен иметь доступный трал"
                    assert city['trail_price'] is not None, f"{city['name']} должен иметь цену трала"


if __name__ == '__main__':
    pytest.main([__file__, '-v'])

