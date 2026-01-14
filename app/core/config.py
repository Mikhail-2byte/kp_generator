import json
import logging
import os
import secrets
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any, Dict

from dotenv import load_dotenv

# Загружаем переменные окружения
load_dotenv()

BASE_DIR = Path(__file__).resolve().parents[2]
CONFIG_DIR = BASE_DIR / 'config'
ENVIRONMENTS_DIR = CONFIG_DIR / 'environments'
DEFAULT_PROFILE = 'development'


def _resolve_log_level_value(value: Any) -> int:
    """
    Приводит значение уровня логирования к числу, устойчиво к строкам в любом регистре.
    Некорректные значения откатываются на INFO, чтобы не рушить инициализацию.
    """
    if isinstance(value, int):
        return value

    if isinstance(value, str):
        normalized = value.strip().upper()
        resolved = logging.getLevelName(normalized)
        if isinstance(resolved, int):
            return resolved

    logging.getLogger(__name__).warning('Unknown log level "%s", fallback to INFO', value)
    return logging.INFO


def _deep_update(target: Dict[str, Any], updates: Dict[str, Any]) -> Dict[str, Any]:
    """Рекурсивно обновляет словарь, не теряя вложенные структуры."""
    for key, value in updates.items():
        if key in target and isinstance(target[key], dict) and isinstance(value, dict):
            _deep_update(target[key], value)
        else:
            target[key] = value
    return target


def _load_json(path: Path, *, logger) -> Dict[str, Any]:
    if not path.exists():
        logger.warning('Config file not found: %s', path.as_posix())
        return {}
    try:
        with path.open('r', encoding='utf-8') as file:
            data = json.load(file)
        logger.info('Config loaded successfully: %s', path.as_posix())
        return data
    except json.JSONDecodeError as exc:
        logger.error('Failed to parse config file %s: %s', path.as_posix(), exc)
        return {}


def _resolve_profile() -> str:
    profile = os.environ.get('APP_ENV') or os.environ.get('FLASK_ENV') or DEFAULT_PROFILE
    return profile.strip().lower()


def load_config(app):
    """Загружает базовую конфигурацию, профиль и переопределения из окружения."""
    try:
        base_path = CONFIG_DIR / 'settings.json'
        config_data = _load_json(base_path, logger=app.logger)

        profile = _resolve_profile()
        profile_path = ENVIRONMENTS_DIR / f'{profile}.json'
        profile_config = _load_json(profile_path, logger=app.logger) if ENVIRONMENTS_DIR.exists() else {}

        _deep_update(config_data, profile_config)
        config_data['profile'] = profile

        # Переопределяем чувствительные данные из переменных окружения
        debug_flag = os.environ.get('DEBUG')
        if debug_flag is None:
            debug_flag = os.environ.get('FLASK_DEBUG')

        env_config = {
            'secret_key': os.environ.get('SECRET_KEY'),
            'database_url': os.environ.get('DATABASE_URL') or profile_config.get('database_url') or 'sqlite:///kp_generator.db',
            'debug': str(debug_flag).lower() == 'true' if debug_flag is not None else profile_config.get('debug', False),
            'log_level': os.environ.get('LOG_LEVEL') or profile_config.get('log_level', 'INFO'),
        }

        _deep_update(config_data, env_config)

        database_url = config_data.get('database_url')
        if database_url:
            os.environ.setdefault('DATABASE_URL', database_url)

        log_level = config_data.get('log_level')
        if log_level:
            os.environ.setdefault('LOG_LEVEL', str(log_level))

        return config_data

    except Exception as exc:
        app.logger.error('Error loading config: %s', exc)
        return {}


def setup_app_security(app, config):
    """Настраивает безопасность приложения"""
    secret_key = config.get('secret_key')

    # В продакшене запрещаем автогенерацию ключа, чтобы не ронять сессии между рестартами
    if (config.get('profile') or '').lower() == 'production' and not secret_key:
        raise RuntimeError('SECRET_KEY обязателен в production. Задайте переменную окружения SECRET_KEY')

    secret_key = secret_key or secrets.token_hex(32)
    app.secret_key = secret_key
    
    # Настройка Flask на основе конфигурации
    app.config['DEBUG'] = config.get('debug', False)
    app.config['TESTING'] = False
    
    return secret_key

def setup_logging(app, config):
    """Настраивает логирование"""
    if not os.path.exists('logs'):
        os.makedirs('logs')
    
    log_level = _resolve_log_level_value(config.get('log_level', 'INFO'))
    
    file_handler = RotatingFileHandler('logs/kp_generator.log', maxBytes=10240, backupCount=10)
    file_handler.setFormatter(logging.Formatter(
        '%(asctime)s %(levelname)s: %(message)s [in %(pathname)s:%(lineno)d]'))
    file_handler.setLevel(log_level)
    
    error_handler = RotatingFileHandler('logs/kp_generator_errors.log', maxBytes=10240, backupCount=5)
    error_handler.setFormatter(logging.Formatter(
        '%(asctime)s %(levelname)s: %(message)s [in %(pathname)s:%(lineno)d]'))
    error_handler.setLevel(logging.ERROR)
    
    app.logger.addHandler(file_handler)
    app.logger.addHandler(error_handler)
    app.logger.setLevel(log_level)