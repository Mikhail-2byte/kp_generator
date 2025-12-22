from typing import Any

from flask import Response, flash, jsonify, render_template, request

from app.core.exceptions import (
    CalculationError,
    DatabaseError,
    DocumentGenerationError,
    KPGeneratorError,
    NotFoundError,
    PermissionError,
    ValidationError,
)
from app.presentation.ui import build_context
from app.core.extensions import login_manager


def _is_api_request() -> bool:
    """
    Определяет, является ли текущий запрос запросом к REST API.

    Используется для выбора между HTML-страницей и JSON-ответом.
    """
    try:
        return request.path.startswith('/api/')
    except RuntimeError:
        # Вне контекста запроса
        return False


def register_error_handlers(app) -> None:
    """Подключает шаблоны отображения для стандартных ошибок HTTP и кастомных исключений."""

    @app.errorhandler(404)
    def not_found_error(_error: Any) -> tuple[str, int]:
        """Возвращает страницу 404 с дружественным описанием."""
        if _is_api_request():
            return jsonify({'error': 'not_found', 'message': 'Ресурс не найден'}), 404
        return render_template(
            '404.html',
            **build_context('index', 'Страница не найдена')
        ), 404

    @app.errorhandler(500)
    def internal_error(_error: Any) -> tuple[str, int]:
        """Отображает страницу 500 при непредвиденных исключениях."""
        if _is_api_request():
            return jsonify({'error': 'internal_error', 'message': 'Внутренняя ошибка сервера'}), 500
        return render_template(
            '500.html',
            **build_context('index', 'Внутренняя ошибка')
        ), 500
    
    @app.errorhandler(NotFoundError)
    def handle_not_found(error: NotFoundError) -> tuple[str, int]:
        """Обрабатывает ошибки NotFoundError."""
        if _is_api_request():
            return jsonify(
                {
                    'error': 'not_found',
                    'message': error.message,
                    'resource_type': error.resource_type,
                    'resource_id': error.resource_id,
                }
            ), 404
        flash(f'Ресурс не найден: {error.message}', 'danger')
        return render_template(
            '404.html',
            **build_context('index', 'Ресурс не найден')
        ), 404
    
    @app.errorhandler(ValidationError)
    def handle_validation_error(error: ValidationError) -> tuple[str, int]:
        """Обрабатывает ошибки валидации."""
        if _is_api_request():
            payload = {
                'error': 'validation_error',
                'message': error.message,
            }
            if getattr(error, 'field', None):
                payload['field'] = error.field
            if getattr(error, 'value', None) is not None:
                payload['value'] = error.value
            if error.details:
                payload['details'] = error.details
            return jsonify(payload), 400
        flash(f'Ошибка валидации: {error.message}', 'danger')
        # Перенаправляем на главную страницу или возвращаем форму с ошибками
        return render_template(
            'index.html',
            **build_context('index', 'Ошибка валидации')
        ), 400
    
    @app.errorhandler(CalculationError)
    def handle_calculation_error(error: CalculationError) -> tuple[str, int]:
        """Обрабатывает ошибки расчета."""
        if _is_api_request():
            payload = {
                'error': 'calculation_error',
                'message': error.message,
            }
            if error.calculation_type:
                payload['calculation_type'] = error.calculation_type
            if error.details:
                payload['details'] = error.details
            return jsonify(payload), 500
        flash(f'Ошибка расчета: {error.message}', 'danger')
        return render_template(
            'index.html',
            **build_context('index', 'Ошибка расчета')
        ), 500
    
    @app.errorhandler(DocumentGenerationError)
    def handle_document_error(error: DocumentGenerationError) -> tuple[str, int]:
        """Обрабатывает ошибки генерации документов."""
        if _is_api_request():
            payload = {
                'error': 'document_generation_error',
                'message': error.message,
            }
            if error.document_type:
                payload['document_type'] = error.document_type
            if error.template_path:
                payload['template_path'] = error.template_path
            if error.details:
                payload['details'] = error.details
            return jsonify(payload), 500
        flash(f'Ошибка генерации документа: {error.message}', 'danger')
        return render_template(
            'index.html',
            **build_context('index', 'Ошибка генерации')
        ), 500
    
    @app.errorhandler(DatabaseError)
    def handle_database_error(error: DatabaseError) -> tuple[str, int]:
        """Обрабатывает ошибки базы данных."""
        app.logger.error(f'Database error: {error.message}', extra=error.details)
        if _is_api_request():
            payload = {
                'error': 'database_error',
                'message': error.message,
            }
            if error.operation:
                payload['operation'] = error.operation
            if error.details:
                payload['details'] = error.details
            return jsonify(payload), 500
        flash('Ошибка базы данных. Попробуйте позже.', 'danger')
        return render_template(
            '500.html',
            **build_context('index', 'Ошибка базы данных')
        ), 500
    
    @app.errorhandler(PermissionError)
    def handle_permission_error(error: PermissionError) -> tuple[str, int]:
        """Обрабатывает ошибки доступа."""
        if _is_api_request():
            payload = {
                'error': 'permission_denied',
                'message': error.message,
            }
            if error.required_permission:
                payload['required_permission'] = error.required_permission
            if error.details:
                payload['details'] = error.details
            return jsonify(payload), 403
        flash(f'Недостаточно прав: {error.message}', 'danger')
        return render_template(
            '500.html',
            **build_context('index', 'Доступ запрещен')
        ), 403
    
    @app.errorhandler(KPGeneratorError)
    def handle_generator_error(error: KPGeneratorError) -> tuple[str, int]:
        """Обрабатывает общие ошибки приложения."""
        app.logger.error(f'Application error: {error.message}', extra=error.details)
        if _is_api_request():
            payload = {
                'error': 'application_error',
                'message': error.message,
            }
            if error.details:
                payload['details'] = error.details
            return jsonify(payload), 500
        flash(f'Ошибка: {error.message}', 'danger')
        return render_template(
            '500.html',
            **build_context('index', 'Ошибка приложения')
        ), 500

    @login_manager.unauthorized_handler
    def handle_unauthorized() -> Response:
        """
        Обработчик неавторизованного доступа для Flask-Login.

        Для API возвращает JSON с кодом 401, для обычных страниц — редирект на login_view.
        """
        if _is_api_request():
            return jsonify({'error': 'unauthorized', 'message': 'Требуется аутентификация'}), 401
        # Для обычного веб-интерфейса используем стандартный редирект
        from flask import redirect, url_for

        login_view = login_manager.login_view or 'auth.profile'
        return redirect(url_for(login_view))
