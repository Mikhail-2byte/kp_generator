"""
Конфигурация для AI агента.
"""

import logging
import os
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv

# Настраиваем логирование
logger = logging.getLogger(__name__)

# Путь к корню проекта (на уровень выше ai_agent/)
BASE_DIR = Path(__file__).resolve().parent.parent

# Загружаем переменные окружения из .env файла в корне проекта
env_path = BASE_DIR / '.env'
load_dotenv(dotenv_path=env_path)

# API конфигурация
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")
OPENROUTER_API_URL = "https://openrouter.ai/api/v1/chat/completions"
# Модель читается из .env (OPENROUTER_MODEL); при работе в Flask — из app.config, заполненного при load_config
_DEFAULT_MODEL = "nvidia/nemotron-3-nano-30b-a3b:free"
MODEL_NAME = os.getenv("OPENROUTER_MODEL", _DEFAULT_MODEL)

# Пути к документации (теперь в папке ai_agent/data/)
AI_AGENT_DIR = Path(__file__).resolve().parent
INSTRUCTIONS_DIR = AI_AGENT_DIR / "data" / "instructions"
DOCS_DIR = AI_AGENT_DIR / "data" / "documentation"

# Проверяем существование папок при импорте (для диагностики)
if not INSTRUCTIONS_DIR.exists():
    import logging
    logging.warning(f'Instructions directory not found: {INSTRUCTIONS_DIR}')
if not DOCS_DIR.exists():
    import logging
    logging.warning(f'Documentation directory not found: {DOCS_DIR}')

# Конфигурация для reasoning
ENABLE_REASONING = os.getenv("OPENROUTER_REASONING_ENABLED", "true").lower() == "true"

# Timeout для запросов к API
API_TIMEOUT = int(os.getenv("OPENROUTER_TIMEOUT", "60"))

# Fallback режим
FALLBACK_ENABLED = os.getenv("AI_FALLBACK_ENABLED", "true").lower() == "true"

# Максимальная длина истории
MAX_HISTORY_LENGTH = int(os.getenv("AI_MAX_HISTORY_LENGTH", "20"))

# Мониторинг использования
USAGE_MONITORING = os.getenv("AI_USAGE_MONITORING", "true").lower() == "true"


class ConfigurationError(Exception):
    """Исключение для ошибок конфигурации."""
    pass


def validate_api_key(api_key: Optional[str]) -> None:
    """
    Валидирует API ключ OpenRouter.
    
    Args:
        api_key: API ключ для валидации
        
    Raises:
        ConfigurationError: Если ключ невалиден
    """
    if not api_key:
        raise ConfigurationError(
            "OPENROUTER_API_KEY не установлен! "
            "Установите переменную окружения OPENROUTER_API_KEY или добавьте её в .env файл. "
            "Получить ключ можно на https://openrouter.ai/keys"
        )
    
    if not api_key.startswith("sk-or-v1-"):
        raise ConfigurationError(
            f"Неверный формат OPENROUTER_API_KEY. "
            f"Ключ должен начинаться с 'sk-or-v1-', получено: {api_key[:15]}..."
        )
    
    if len(api_key) < 40:
        raise ConfigurationError(
            f"OPENROUTER_API_KEY слишком короткий. "
            f"Ожидается минимум 40 символов, получено: {len(api_key)}"
        )


def get_api_key() -> str:
    """
    Возвращает и валидирует API ключ OpenRouter.
    
    Returns:
        Валидный API ключ
        
    Raises:
        ConfigurationError: Если ключ не установлен или невалиден
    """
    validate_api_key(OPENROUTER_API_KEY)
    return OPENROUTER_API_KEY


def get_model_name() -> str:
    """Возвращает название модели из переменной окружения OPENROUTER_MODEL (или из app.config при работе в Flask)."""
    try:
        from flask import current_app
        from flask import has_app_context
        if has_app_context():
            model = (current_app.config.get('APP_SETTINGS') or {}).get('openrouter_model')
            if model:
                return model
    except (ImportError, RuntimeError):
        pass
    return os.getenv("OPENROUTER_MODEL", _DEFAULT_MODEL)


def get_api_url() -> str:
    """Возвращает URL API OpenRouter."""
    return OPENROUTER_API_URL


def is_reasoning_enabled() -> bool:
    """Проверяет, включен ли reasoning."""
    return ENABLE_REASONING


def get_timeout() -> int:
    """Возвращает timeout для API запросов в секундах."""
    return API_TIMEOUT


def is_fallback_enabled() -> bool:
    """Проверяет, включен ли fallback режим."""
    return FALLBACK_ENABLED


def get_max_history_length() -> int:
    """Возвращает максимальную длину истории диалога."""
    return MAX_HISTORY_LENGTH


def is_usage_monitoring_enabled() -> bool:
    """Проверяет, включен ли мониторинг использования API."""
    return USAGE_MONITORING

