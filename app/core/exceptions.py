"""Иерархия кастомных исключений для приложения."""

from typing import Any, Dict, Optional


class KPGeneratorError(Exception):
    """Базовое исключение для всех ошибок приложения."""
    
    def __init__(self, message: str, details: Optional[Dict[str, Any]] = None):
        super().__init__(message)
        self.message = message
        self.details = details or {}


class ValidationError(KPGeneratorError):
    """Ошибка валидации входных данных."""
    
    def __init__(
        self,
        message: str,
        field: Optional[str] = None,
        value: Any = None,
        details: Optional[Dict[str, Any]] = None
    ):
        super().__init__(message, details)
        self.field = field
        self.value = value


class CalculationError(KPGeneratorError):
    """Ошибка при расчете цен или логистики."""
    
    def __init__(
        self,
        message: str,
        calculation_type: Optional[str] = None,
        details: Optional[Dict[str, Any]] = None
    ):
        super().__init__(message, details)
        self.calculation_type = calculation_type


class DocumentGenerationError(KPGeneratorError):
    """Ошибка при генерации документов."""
    
    def __init__(
        self,
        message: str,
        document_type: Optional[str] = None,
        template_path: Optional[str] = None,
        details: Optional[Dict[str, Any]] = None
    ):
        super().__init__(message, details)
        self.document_type = document_type
        self.template_path = template_path


class DatabaseError(KPGeneratorError):
    """Ошибка при работе с базой данных."""
    
    def __init__(
        self,
        message: str,
        operation: Optional[str] = None,
        details: Optional[Dict[str, Any]] = None
    ):
        super().__init__(message, details)
        self.operation = operation


class NotFoundError(KPGeneratorError):
    """Ресурс не найден."""
    
    def __init__(
        self,
        message: str,
        resource_type: Optional[str] = None,
        resource_id: Optional[str] = None,
        details: Optional[Dict[str, Any]] = None
    ):
        super().__init__(message, details)
        self.resource_type = resource_type
        self.resource_id = resource_id


class PermissionError(KPGeneratorError):
    """Ошибка доступа (недостаточно прав)."""
    
    def __init__(
        self,
        message: str,
        required_permission: Optional[str] = None,
        details: Optional[Dict[str, Any]] = None
    ):
        super().__init__(message, details)
        self.required_permission = required_permission


class ConfigurationError(KPGeneratorError):
    """Ошибка конфигурации приложения."""
    
    def __init__(
        self,
        message: str,
        config_key: Optional[str] = None,
        details: Optional[Dict[str, Any]] = None
    ):
        super().__init__(message, details)
        self.config_key = config_key


class ImportError(KPGeneratorError):
    """Ошибка при импорте данных (Excel, CSV и т.д.)."""
    
    def __init__(
        self,
        message: str,
        file_type: Optional[str] = None,
        row_number: Optional[int] = None,
        details: Optional[Dict[str, Any]] = None
    ):
        super().__init__(message, details)
        self.file_type = file_type
        self.row_number = row_number


__all__ = [
    'KPGeneratorError',
    'ValidationError',
    'CalculationError',
    'DocumentGenerationError',
    'DatabaseError',
    'NotFoundError',
    'PermissionError',
    'ConfigurationError',
    'ImportError',
]

