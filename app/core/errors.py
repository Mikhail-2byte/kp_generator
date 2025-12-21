from typing import Any

from flask import Response, flash, render_template

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


def register_error_handlers(app) -> None:
    """Подключает шаблоны отображения для стандартных ошибок HTTP и кастомных исключений."""
    
    @app.errorhandler(404)
    def not_found_error(_error: Any) -> tuple[str, int]:
        """Возвращает страницу 404 с дружественным описанием."""
        return render_template(
            '404.html',
            **build_context('index', 'Страница не найдена')
        ), 404

    @app.errorhandler(500)
    def internal_error(_error: Any) -> tuple[str, int]:
        """Отображает страницу 500 при непредвиденных исключениях."""
        return render_template(
            '500.html',
            **build_context('index', 'Внутренняя ошибка')
        ), 500
    
    @app.errorhandler(NotFoundError)
    def handle_not_found(error: NotFoundError) -> tuple[str, int]:
        """Обрабатывает ошибки NotFoundError."""
        flash(f'Ресурс не найден: {error.message}', 'danger')
        return render_template(
            '404.html',
            **build_context('index', 'Ресурс не найден')
        ), 404
    
    @app.errorhandler(ValidationError)
    def handle_validation_error(error: ValidationError) -> tuple[str, int]:
        """Обрабатывает ошибки валидации."""
        flash(f'Ошибка валидации: {error.message}', 'danger')
        # Перенаправляем на главную страницу или возвращаем форму с ошибками
        return render_template(
            'index.html',
            **build_context('index', 'Ошибка валидации')
        ), 400
    
    @app.errorhandler(CalculationError)
    def handle_calculation_error(error: CalculationError) -> tuple[str, int]:
        """Обрабатывает ошибки расчета."""
        flash(f'Ошибка расчета: {error.message}', 'danger')
        return render_template(
            'index.html',
            **build_context('index', 'Ошибка расчета')
        ), 500
    
    @app.errorhandler(DocumentGenerationError)
    def handle_document_error(error: DocumentGenerationError) -> tuple[str, int]:
        """Обрабатывает ошибки генерации документов."""
        flash(f'Ошибка генерации документа: {error.message}', 'danger')
        return render_template(
            'index.html',
            **build_context('index', 'Ошибка генерации')
        ), 500
    
    @app.errorhandler(DatabaseError)
    def handle_database_error(error: DatabaseError) -> tuple[str, int]:
        """Обрабатывает ошибки базы данных."""
        app.logger.error(f'Database error: {error.message}', extra=error.details)
        flash('Ошибка базы данных. Попробуйте позже.', 'danger')
        return render_template(
            '500.html',
            **build_context('index', 'Ошибка базы данных')
        ), 500
    
    @app.errorhandler(PermissionError)
    def handle_permission_error(error: PermissionError) -> tuple[str, int]:
        """Обрабатывает ошибки доступа."""
        flash(f'Недостаточно прав: {error.message}', 'danger')
        return render_template(
            '500.html',
            **build_context('index', 'Доступ запрещен')
        ), 403
    
    @app.errorhandler(KPGeneratorError)
    def handle_generator_error(error: KPGeneratorError) -> tuple[str, int]:
        """Обрабатывает общие ошибки приложения."""
        app.logger.error(f'Application error: {error.message}', extra=error.details)
        flash(f'Ошибка: {error.message}', 'danger')
        return render_template(
            '500.html',
            **build_context('index', 'Ошибка приложения')
        ), 500
