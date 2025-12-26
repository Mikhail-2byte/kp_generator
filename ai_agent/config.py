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
)
OPENROUTER_API_URL = "https://openrouter.ai/api/v1/chat/completions"
MODEL_NAME = "xiaomi/mimo-v2-flash:free"

# Пути к документации
INSTRUCTIONS_DIR = BASE_DIR / "static" / "instructions"
DOCS_DIR = BASE_DIR / "docs"

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

