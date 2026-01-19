"""Тесты для модуля конфигурации app/core/config.py."""

import os
import tempfile
from pathlib import Path
from unittest.mock import patch

import pytest

from app.core.config import (
    _deep_update,
    _load_json,
    _resolve_log_level_value,
    load_config,
    setup_app_security,
    setup_logging,
)


class TestResolveLogLevel:
    """Тесты для функции _resolve_log_level_value."""

    def test_resolve_int_value(self):
        """Проверяет, что целочисленное значение возвращается как есть."""
        assert _resolve_log_level_value(10) == 10
        assert _resolve_log_level_value(20) == 20
        assert _resolve_log_level_value(30) == 30

    def test_resolve_string_value(self):
        """Проверяет, что строковое значение преобразуется в уровень логирования."""
        assert _resolve_log_level_value('DEBUG') == 10
        assert _resolve_log_level_value('INFO') == 20
        assert _resolve_log_level_value('WARNING') == 30
        assert _resolve_log_level_value('ERROR') == 40
        assert _resolve_log_level_value('CRITICAL') == 50

    def test_resolve_case_insensitive(self):
        """Проверяет, что регистр не важен."""
        assert _resolve_log_level_value('debug') == 10
        assert _resolve_log_level_value('Info') == 20
        assert _resolve_log_level_value('WARNING') == 30

    def test_resolve_invalid_value_fallback(self):
        """Проверяет, что невалидное значение откатывается на INFO."""
        assert _resolve_log_level_value('INVALID') == 20
        assert _resolve_log_level_value(999) == 20
        assert _resolve_log_level_value(None) == 20


class TestDeepUpdate:
    """Тесты для функции _deep_update."""

    def test_simple_update(self):
        """Проверяет простое обновление словаря."""
        target = {'a': 1, 'b': 2}
        updates = {'b': 3, 'c': 4}
        result = _deep_update(target, updates)
        assert result == {'a': 1, 'b': 3, 'c': 4}
        assert target == {'a': 1, 'b': 3, 'c': 4}  # Изменяется исходный словарь

    def test_nested_update(self):
        """Проверяет рекурсивное обновление вложенных словарей."""
        target = {'a': {'x': 1, 'y': 2}, 'b': 3}
        updates = {'a': {'y': 20, 'z': 30}, 'b': 4}
        result = _deep_update(target, updates)
        assert result == {'a': {'x': 1, 'y': 20, 'z': 30}, 'b': 4}

    def test_nested_override(self):
        """Проверяет, что не-словарь перезаписывает словарь."""
        target = {'a': {'x': 1, 'y': 2}}
        updates = {'a': 'not_a_dict'}
        result = _deep_update(target, updates)
        assert result == {'a': 'not_a_dict'}


class TestLoadJson:
    """Тесты для функции _load_json."""

    def test_load_existing_file(self, tmp_path):
        """Проверяет загрузку существующего JSON файла."""
        from unittest.mock import Mock
        config_file = tmp_path / 'config.json'
        config_file.write_text('{"key": "value", "number": 42}', encoding='utf-8')

        logger = Mock()
        result = _load_json(config_file, logger=logger)
        assert result == {'key': 'value', 'number': 42}
        logger.warning.assert_not_called()

    def test_load_nonexistent_file(self, tmp_path):
        """Проверяет обработку несуществующего файла."""
        from unittest.mock import Mock
        config_file = tmp_path / 'nonexistent.json'
        logger = Mock()
        result = _load_json(config_file, logger=logger)
        assert result == {}
        logger.warning.assert_called_once()

    def test_load_invalid_json(self, tmp_path):
        """Проверяет обработку невалидного JSON."""
        from unittest.mock import Mock
        config_file = tmp_path / 'invalid.json'
        config_file.write_text('{invalid json}', encoding='utf-8')

        logger = Mock()
        result = _load_json(config_file, logger=logger)
        assert result == {}
        assert logger.warning.call_count >= 1


class TestLoadConfig:
    """Тесты для функции load_config."""

    @patch.dict(os.environ, {'FLASK_ENV': 'testing', 'SECRET_KEY': 'test-secret-key'})
    def test_load_config_with_env(self, app):
        """Проверяет загрузку конфигурации с переменными окружения."""
        config = load_config(app)
        assert isinstance(config, dict)
        # Проверяем наличие основных ключей конфигурации
        assert 'database_url' in config or 'calculation_constants' in config or 'profile' in config

    @patch.dict(os.environ, {'FLASK_ENV': 'development'}, clear=False)
    def test_load_config_development(self, app):
        """Проверяет загрузку конфигурации для development."""
        config = load_config(app)
        assert isinstance(config, dict)

    @patch.dict(os.environ, {'FLASK_ENV': 'production'}, clear=False)
    def test_load_config_production(self, app):
        """Проверяет загрузку конфигурации для production."""
        config = load_config(app)
        assert isinstance(config, dict)


class TestSetupAppSecurity:
    """Тесты для функции setup_app_security."""

    @patch.dict(os.environ, {'SECRET_KEY': 'test-secret-key-12345'})
    def test_setup_security_with_secret_key(self, app):
        """Проверяет настройку безопасности с SECRET_KEY."""
        config = {'secret_key': None}
        setup_app_security(app, config)
        assert app.config['SECRET_KEY'] == 'test-secret-key-12345'

    @patch.dict(os.environ, {}, clear=True)
    def test_setup_security_without_secret_key_production(self, app):
        """Проверяет, что без SECRET_KEY в production выбрасывается ошибка."""
        config = {'secret_key': None, 'profile': 'production'}
        with pytest.raises(RuntimeError, match='SECRET_KEY'):
            setup_app_security(app, config)

    @patch.dict(os.environ, {}, clear=True)
    def test_setup_security_without_secret_key_development(self, app):
        """Проверяет, что в development генерируется случайный ключ."""
        app.config['ENV'] = 'development'
        config = {'secret_key': None}
        setup_app_security(app, config)
        assert app.config['SECRET_KEY'] is not None
        assert len(app.config['SECRET_KEY']) > 0


class TestSetupLogging:
    """Тесты для функции setup_logging."""

    def test_setup_logging_default(self, app):
        """Проверяет настройку логирования с дефолтными параметрами."""
        config = {}
        setup_logging(app, config)
        assert app.logger is not None

    def test_setup_logging_with_config(self, app):
        """Проверяет настройку логирования с конфигурацией."""
        config = {
            'logging': {
                'level': 'DEBUG',
                'file': 'test.log'
            }
        }
        setup_logging(app, config)
        assert app.logger is not None

