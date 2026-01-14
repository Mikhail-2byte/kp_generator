"""Тесты для модуля обработки ошибок app/core/errors.py."""

from unittest.mock import patch

import pytest
from werkzeug.exceptions import RequestEntityTooLarge

from app.core.errors import _is_api_request, register_error_handlers
from app.core.exceptions import (
    CalculationError,
    DatabaseError,
    DocumentGenerationError,
    NotFoundError,
    PermissionError,
    ValidationError,
)


class TestIsApiRequest:
    """Тесты для функции _is_api_request."""

    def test_is_api_request_with_api_path(self, app):
        """Проверяет, что API путь определяется правильно."""
        with app.test_request_context('/api/health'):
            assert _is_api_request() is True

    def test_is_api_request_with_non_api_path(self, app):
        """Проверяет, что не-API путь определяется правильно."""
        with app.test_request_context('/'):
            assert _is_api_request() is False

    def test_is_api_request_outside_context(self):
        """Проверяет, что вне контекста запроса возвращается False."""
        assert _is_api_request() is False


class TestErrorHandlers:
    """Тесты для обработчиков ошибок."""

    def test_register_error_handlers(self, app):
        """Проверяет, что обработчики ошибок регистрируются."""
        register_error_handlers(app)
        # Проверяем, что обработчики зарегистрированы
        assert 404 in app.error_handler_spec[None]
        assert 500 in app.error_handler_spec[None]

    def test_404_handler_api(self, app, client):
        """Проверяет обработку 404 для API запроса."""
        register_error_handlers(app)
        response = client.get('/api/nonexistent')
        assert response.status_code == 404
        assert response.is_json
        data = response.get_json()
        assert 'error' in data
        assert data['error'] == 'not_found'

    def test_404_handler_ui(self, app, client):
        """Проверяет обработку 404 для UI запроса."""
        register_error_handlers(app)
        response = client.get('/nonexistent-page')
        assert response.status_code == 404
        # В UI режиме должен возвращаться HTML
        assert response.content_type.startswith('text/html')

    def test_500_handler_api(self, app):
        """Проверяет обработку 500 для API запроса."""
        register_error_handlers(app)

        @app.route('/test-error')
        def test_error():
            raise ValueError('Test error')

        with app.test_client() as client:
            response = client.get('/test-error')
            assert response.status_code == 500
            assert response.is_json
            data = response.get_json()
            assert 'error' in data
            assert data['error'] == 'internal_error'

    def test_validation_error_handler_api(self, app, client):
        """Проверяет обработку ValidationError для API."""
        register_error_handlers(app)

        @app.route('/test-validation')
        def test_validation():
            raise ValidationError('Invalid data', field='test_field', value='test_value')

        with app.test_client() as client:
            response = client.get('/test-validation')
            assert response.status_code == 400
            assert response.is_json
            data = response.get_json()
            assert data['error'] == 'validation_error'
            assert 'field' in data
            assert data['field'] == 'test_field'

    def test_calculation_error_handler_api(self, app, client):
        """Проверяет обработку CalculationError для API."""
        register_error_handlers(app)

        @app.route('/test-calculation')
        def test_calculation():
            raise CalculationError('Calculation failed', calculation_type='test')

        with app.test_client() as client:
            response = client.get('/test-calculation')
            assert response.status_code == 500
            assert response.is_json
            data = response.get_json()
            assert data['error'] == 'calculation_error'

    def test_document_generation_error_handler_api(self, app, client):
        """Проверяет обработку DocumentGenerationError для API."""
        register_error_handlers(app)

        @app.route('/test-document')
        def test_document():
            raise DocumentGenerationError('Document generation failed', document_type='test')

        with app.test_client() as client:
            response = client.get('/test-document')
            assert response.status_code == 500
            assert response.is_json
            data = response.get_json()
            assert data['error'] == 'document_generation_error'

    def test_not_found_error_handler_api(self, app, client):
        """Проверяет обработку NotFoundError для API."""
        register_error_handlers(app)

        @app.route('/test-not-found')
        def test_not_found():
            raise NotFoundError('Resource not found', resource_type='test', resource_id='123')

        with app.test_client() as client:
            response = client.get('/test-not-found')
            assert response.status_code == 404
            assert response.is_json
            data = response.get_json()
            assert data['error'] == 'not_found'
            assert 'resource_type' in data

    def test_database_error_handler_api(self, app, client):
        """Проверяет обработку DatabaseError для API."""
        register_error_handlers(app)

        @app.route('/test-database')
        def test_database():
            raise DatabaseError('Database error', operation='test')

        with app.test_client() as client:
            response = client.get('/test-database')
            assert response.status_code == 500
            assert response.is_json
            data = response.get_json()
            assert data['error'] == 'database_error'

    def test_permission_error_handler_api(self, app, client):
        """Проверяет обработку PermissionError для API."""
        register_error_handlers(app)

        @app.route('/test-permission')
        def test_permission():
            raise PermissionError('Permission denied', required_permission='admin')

        with app.test_client() as client:
            response = client.get('/test-permission')
            assert response.status_code == 403
            assert response.is_json
            data = response.get_json()
            assert data['error'] == 'permission_denied'

    def test_request_entity_too_large_handler_api(self, app, client):
        """Проверяет обработку RequestEntityTooLarge для API."""
        register_error_handlers(app)

        @app.route('/test-large')
        def test_large():
            raise RequestEntityTooLarge()

        with app.test_client() as client:
            response = client.get('/test-large')
            assert response.status_code == 413
            assert response.is_json
            data = response.get_json()
            assert data['error'] == 'file_too_large'

