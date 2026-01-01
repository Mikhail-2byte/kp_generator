"""
Конфигурация для AI агента.
"""

import os
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv

# Загружаем переменные окружения
load_dotenv()

# Путь к корню проекта (на уровень выше ai_agent/)
BASE_DIR = Path(__file__).resolve().parent.parent

# API конфигурация
OPENROUTER_API_KEY = os.getenv(
    "OPENROUTER_API_KEY",
    "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"
    #"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"
)
OPENROUTER_API_URL = "https://openrouter.ai/api/v1/chat/completions"
MODEL_NAME = "xiaomi/mimo-v2-flash:free"

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
ENABLE_REASONING = True


def get_api_key() -> str:
    """Возвращает API ключ OpenRouter."""
    return OPENROUTER_API_KEY


def get_model_name() -> str:
    """Возвращает название модели."""
    return MODEL_NAME


def get_api_url() -> str:
    """Возвращает URL API OpenRouter."""
    return OPENROUTER_API_URL


def is_reasoning_enabled() -> bool:
    """Проверяет, включен ли reasoning."""
    return ENABLE_REASONING

