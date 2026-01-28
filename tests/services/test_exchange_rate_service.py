"""
Тесты для сервиса получения курса валют.
"""
import sys
from pathlib import Path

# Добавляем корень проекта в путь для возможности прямого запуска
project_root = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(project_root))

import pytest
from app.services.exchange_rate_service import (
    get_currency_rate,
    get_exchange_rate_info,
    get_cached_rate,
    clear_cache,
    _get_default_rate
)


@pytest.mark.integration
def test_get_currency_rate_success():
    """
    Тест успешного получения курса валют из API.
    Этот тест требует подключения к интернету.
    """
    result = get_currency_rate()
    
    assert result is not None, "Курс валют должен быть получен"
    rate, date_str = result
    
    assert isinstance(rate, float), "Курс должен быть числом"
    assert rate > 0, "Курс должен быть положительным"
    assert 5.0 < rate < 20.0, f"Курс CNY к RUB должен быть в разумных пределах (получен: {rate})"
    assert isinstance(date_str, str), "Дата должна быть строкой"
    assert len(date_str) > 0, "Дата не должна быть пустой"


@pytest.mark.integration
def test_get_exchange_rate_info():
    """
    Тест получения информации о курсе валют через API endpoint.
    """
    # Очищаем кэш перед тестом
    clear_cache()
    
    info = get_exchange_rate_info()
    
    assert isinstance(info, dict), "Информация должна быть словарем"
    assert 'rate' in info, "Должен быть ключ 'rate'"
    assert 'date' in info, "Должен быть ключ 'date'"
    assert 'source' in info, "Должен быть ключ 'source'"
    
    assert isinstance(info['rate'], (int, float)), "Курс должен быть числом"
    assert info['rate'] > 0, "Курс должен быть положительным"
    assert info['source'] in ['api', 'cache', 'config'], f"Источник должен быть одним из: api, cache, config (получен: {info['source']})"


@pytest.mark.integration
def test_get_cached_rate():
    """
    Тест получения кэшированного курса.
    """
    # Очищаем кэш перед тестом
    clear_cache()
    
    # Первый вызов должен получить курс из API
    rate1 = get_cached_rate()
    assert isinstance(rate1, float), "Курс должен быть числом"
    assert rate1 > 0, "Курс должен быть положительным"
    
    # Второй вызов должен использовать кэш
    rate2 = get_cached_rate()
    assert rate1 == rate2, "Второй вызов должен вернуть тот же курс из кэша"


def test_get_default_rate():
    """
    Тест получения курса по умолчанию из конфига.
    """
    default_rate = _get_default_rate()
    
    assert isinstance(default_rate, float), "Курс по умолчанию должен быть числом"
    assert default_rate > 0, "Курс по умолчанию должен быть положительным"
    assert default_rate == 11.5, "Курс по умолчанию должен быть 11.5 из конфига"


def test_clear_cache():
    """
    Тест очистки кэша.
    """
    # Устанавливаем кэш (если его нет)
    get_cached_rate()
    
    # Очищаем кэш
    clear_cache()
    
    # Проверяем, что кэш очищен (следующий вызов должен получить курс заново)
    # Это косвенная проверка, так как мы не можем напрямую проверить внутреннее состояние
    rate = get_cached_rate()
    assert isinstance(rate, float), "После очистки кэша курс должен быть получен"
    assert rate > 0, "Курс должен быть положительным"


@pytest.mark.integration
def test_exchange_rate_api_endpoint_format():
    """
    Тест формата ответа API endpoint для курса валют.
    Проверяет, что все необходимые поля присутствуют.
    """
    clear_cache()
    
    info = get_exchange_rate_info()
    
    # Проверяем обязательные поля
    required_fields = ['rate', 'date', 'source', 'cached']
    for field in required_fields:
        assert field in info, f"В ответе должен быть ключ '{field}'"
    
    # Проверяем типы данных
    assert isinstance(info['rate'], (int, float)), "rate должен быть числом"
    assert isinstance(info['date'], str), "date должна быть строкой"
    assert isinstance(info['source'], str), "source должна быть строкой"
    assert isinstance(info['cached'], bool), "cached должна быть булевым значением"
    
    # Проверяем значения
    assert info['rate'] > 0, "rate должен быть положительным"
    assert len(info['date']) > 0, "date не должна быть пустой"
    assert info['source'] in ['api', 'cache', 'config'], "source должен быть одним из: api, cache, config"


if __name__ == "__main__":
    # Запуск тестов при прямом выполнении файла
    pytest.main([__file__, "-v", "-m", "integration"])
