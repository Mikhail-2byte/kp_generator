# -*- coding: utf-8 -*-
"""
Сервис для получения актуального курса китайского юаня (CNY) к российскому рублю (RUB)
"""

import json
import logging
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional, Tuple

import requests

logger = logging.getLogger(__name__)

# Кэш для хранения курса валют
_cache: Optional[Tuple[float, str, datetime]] = None
_CACHE_TTL_HOURS = 1  # Время жизни кэша в часах


def _get_default_rate() -> float:
    """
    Получает значение курса по умолчанию из конфига.
    
    Returns:
        Курс валют из конфига или 11.5 по умолчанию
    """
    try:
        config_path = Path(__file__).resolve().parents[2] / 'config' / 'settings.json'
        if config_path.exists():
            with config_path.open('r', encoding='utf-8') as f:
                config = json.load(f)
                conversion_rate = config.get('calculation_constants', {}).get('conversion_rate', 11.5)
                return float(conversion_rate)
    except Exception as e:
        logger.warning('Не удалось загрузить курс из конфига: %s', e)
    
    return 11.5  # Значение по умолчанию


def _try_api_source(url: str, headers: dict, timeout: int = 10) -> Optional[Tuple[float, str]]:
    """
    Пытается получить курс из одного источника API.
    
    Returns:
        Кортеж (курс, дата) или None в случае ошибки
    """
    try:
        # Используем session для лучшей производительности
        session = requests.Session()
        session.headers.update(headers)
        
        response = session.get(url, timeout=timeout, stream=False)
        response.raise_for_status()
        
        data = response.json()
        
        # Проверяем наличие курса RUB (поддерживаем разные форматы API)
        if 'rates' not in data or 'RUB' not in data['rates']:
            return None
        
        rub_rate = data['rates']['RUB']
        
        # Получаем дату курса (разные API используют разные поля)
        date_str = (
            data.get('date') or 
            data.get('time_last_update_utc', '').split(' ')[0] if 'time_last_update_utc' in data else None or
            datetime.now().strftime('%Y-%m-%d')
        )
        
        # Если date_str содержит полную дату-время, извлекаем только дату
        if ' ' in str(date_str):
            date_str = str(date_str).split(' ')[0]
        
        return float(rub_rate), str(date_str)
    except Exception as e:
        logger.debug("Ошибка при запросе к API источнику %s: %s", url, e)
        return None


def get_currency_rate() -> Optional[Tuple[float, str]]:
    """
    Получает актуальный курс CNY к RUB через бесплатный API.
    Пробует несколько источников с fallback.
    
    Returns:
        Кортеж (курс, дата) или None в случае ошибки
    """
    # Список альтернативных API источников (пробуем в порядке приоритета)
    api_sources = [
        {
            'url': 'https://open.er-api.com/v6/latest/CNY',
            'name': 'open.er-api.com',
            'timeout': 8  # Более быстрый и надежный источник
        },
        {
            'url': 'https://api.exchangerate-api.com/v4/latest/CNY',
            'name': 'exchangerate-api.com',
            'timeout': 8  # Резервный источник
        }
    ]
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'application/json',
    }
    
    logger.info("Попытка получить курс валют из API (пробуем %d источников)", len(api_sources))
    print(f"[EXCHANGE_RATE] Попытка получить курс из {len(api_sources)} источников...")
    request_start_time = datetime.now()
    
    # Пробуем каждый источник по очереди
    for idx, source in enumerate(api_sources, 1):
        url = source['url']
        name = source['name']
        timeout = source['timeout']
        
        logger.info("Попытка %d/%d: %s", idx, len(api_sources), name)
        print(f"[EXCHANGE_RATE] Попытка {idx}/{len(api_sources)}: {name}")
        
        try:
            result = _try_api_source(url, headers, timeout)
            if result:
                rate, date_str = result
                duration = (datetime.now() - request_start_time).total_seconds()
                logger.info("✓ Курс валют успешно получен из %s: 1 CNY = %.4f RUB (дата: %s, за %.2f сек)", 
                           name, rate, date_str, duration)
                print(f"[EXCHANGE_RATE] ✓ Курс получен из {name}: 1 CNY = {rate:.4f} RUB (дата: {date_str}, за {duration:.2f} сек)")
                return rate, date_str
            else:
                logger.warning("Не удалось получить курс из %s", name)
                print(f"[EXCHANGE_RATE] Не удалось получить курс из {name}")
        except Exception as e:
            duration = (datetime.now() - request_start_time).total_seconds()
            logger.warning("Ошибка при запросе к %s после %.2f сек: %s", name, duration, e)
            print(f"[EXCHANGE_RATE] Ошибка при запросе к {name}: {type(e).__name__}")
            continue
    
    # Если все источники не сработали
    total_duration = (datetime.now() - request_start_time).total_seconds()
    logger.error("Не удалось получить курс ни из одного источника (всего попыток: %d, время: %.2f сек)", 
                len(api_sources), total_duration)
    print(f"[EXCHANGE_RATE] ✗ Не удалось получить курс ни из одного источника (время: {total_duration:.2f} сек)")
    
    return None


def get_cached_rate() -> float:
    """
    Возвращает кэшированный курс или значение по умолчанию.
    Обновляет кэш, если он устарел или отсутствует.
    
    Returns:
        Актуальный курс CNY к RUB
    """
    global _cache
    
    # Проверяем, нужно ли обновить кэш
    if _cache is None:
        # Кэш пуст, пытаемся получить курс
        result = get_currency_rate()
        if result:
            rate, date_str = result
            _cache = (rate, date_str, datetime.now())
            logger.info("Курс валют загружен из API: 1 CNY = %.4f RUB (дата: %s)", rate, date_str)
            return rate
        else:
            # Не удалось получить курс, используем значение по умолчанию
            default_rate = _get_default_rate()
            logger.warning("Не удалось получить курс из API, используется значение по умолчанию: %.2f", default_rate)
            return default_rate
    
    rate, date_str, cached_at = _cache
    
    # Проверяем, не устарел ли кэш
    if datetime.now() - cached_at > timedelta(hours=_CACHE_TTL_HOURS):
        # Кэш устарел, пытаемся обновить
        result = get_currency_rate()
        if result:
            new_rate, new_date_str = result
            _cache = (new_rate, new_date_str, datetime.now())
            logger.info("Курс валют обновлен из API: 1 CNY = %.4f RUB (дата: %s)", new_rate, new_date_str)
            return new_rate
        else:
            # Не удалось обновить, используем старый кэш
            logger.warning("Не удалось обновить курс из API, используется кэшированное значение: %.2f", rate)
            return rate
    
    # Кэш актуален, возвращаем его
    return rate


def get_exchange_rate_info() -> dict:
    """
    Возвращает информацию о текущем курсе валют для API endpoint.
    Использует кэш и обновляет его только если он устарел (старше 1 часа).
    Это предотвращает перегрузку внешнего API.
    
    Returns:
        Словарь с информацией о курсе: rate, date, source, cached
    """
    global _cache
    
    # Сначала проверяем кэш
    if _cache is not None:
        rate, date_str, cached_at = _cache
        cache_age = datetime.now() - cached_at
        
        # Если кэш свежий (меньше 1 часа), возвращаем его БЕЗ запроса к API
        if cache_age < timedelta(hours=_CACHE_TTL_HOURS):
            cache_age_minutes = int(cache_age.total_seconds() / 60)
            logger.debug("Используется кэшированный курс: %.4f RUB (возраст кэша: %d мин)", 
                        rate, cache_age_minutes)
            return {
                'rate': round(rate, 4),
                'date': date_str,
                'source': 'cache',
                'cached': True,
                'cached_at': cached_at.isoformat(),
                'cache_age_minutes': cache_age_minutes
            }
        else:
            # Кэш устарел, нужно обновить
            cache_age_hours = cache_age.total_seconds() / 3600
            logger.info("Кэш устарел (возраст: %.1f часов), обновляем из API...", cache_age_hours)
    
    # Кэш отсутствует или устарел - пытаемся получить актуальный курс из API
    logger.info("Запрос информации о курсе валют (кэш: %s)", "устарел" if _cache else "отсутствует")
    result = get_currency_rate()
    
    if result:
        rate, date_str = result
        # Обновляем кэш
        _cache = (rate, date_str, datetime.now())
        logger.info("Курс валют обновлен в кэше: %.4f RUB (будет кэширован на %d часов)", 
                   rate, _CACHE_TTL_HOURS)
        return {
            'rate': round(rate, 4),
            'date': date_str,
            'source': 'api',
            'cached': False
        }
    
    # Если не удалось получить из API, проверяем старый кэш (даже если устарел)
    if _cache:
        rate, date_str, cached_at = _cache
        cache_age = datetime.now() - cached_at
        cache_age_hours = cache_age.total_seconds() / 3600
        logger.warning("Не удалось обновить курс из API, используется устаревший кэш (возраст: %.1f часов)", 
                      cache_age_hours)
        return {
            'rate': round(rate, 4),
            'date': date_str,
            'source': 'cache',
            'cached': True,
            'cached_at': cached_at.isoformat(),
            'cache_age_minutes': int(cache_age.total_seconds() / 60),
            'stale': True  # Помечаем как устаревший
        }
    
    # Используем значение по умолчанию
    default_rate = _get_default_rate()
    logger.warning("Используется курс по умолчанию из конфига: %.2f RUB", default_rate)
    print(f"[EXCHANGE_RATE] Используется значение по умолчанию: {default_rate} RUB")
    return {
        'rate': default_rate,
        'date': datetime.now().strftime('%Y-%m-%d'),
        'source': 'config',
        'cached': False
    }


def clear_cache() -> None:
    """
    Очищает кэш курса валют (для тестирования или принудительного обновления).
    """
    global _cache
    _cache = None
    logger.info("Кэш курса валют очищен")
