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
_CACHE_TTL_MINUTES = 30 # Время жизни кэша в минутах


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
        
        # Парсим JSON ответ
        data = response.json()
        rates = data.get('rates', {})
        rate = rates.get('RUB')
        if rate is None:
            return None
        
        rub_rate = float(rate)
        
        # Получаем дату курса (разные API используют разные поля)
        date_str = (
            data.get('date') or 
            data.get('time_last_update_utc', '').split(' ')[0] if 'time_last_update_utc' in data else None or
            datetime.now().strftime('%Y-%m-%d')
        )
        
        # Если date_str содержит полную дату-время, извлекаем только дату
        if ' ' in str(date_str):
            date_str = str(date_str).split(' ')[0]
        
        return rub_rate, str(date_str)
    except Exception as e:
        logger.debug("Ошибка при запросе к API источнику %s: %s", url, e)
        return None


def get_currency_rate():
    """
    Получает актуальный курс CNY к RUB от ЦБ РФ
    """
    import requests
    import xml.etree.ElementTree as ET
    from datetime import datetime
    
    try:
        # Получаем курс от ЦБ РФ (основной источник)
        cbr_url = "https://www.cbr.ru/scripts/XML_daily.asp"
        logger.debug("Запрос курса к ЦБ РФ...")
        
        response = requests.get(cbr_url, timeout=10)
        if response.status_code == 200:
            root = ET.fromstring(response.content)
            
            # Ищем курс CNY
            for valute in root.findall('Valute'):
                char_code = valute.find('CharCode')
                if char_code is not None and char_code.text == 'CNY':
                    name = valute.find('Name')
                    value = valute.find('Value')
                    nominal = valute.find('Nominal')
                    
                    if name is not None and value is not None and nominal is not None:
                        rate = float(value.text.replace(',', '.')) / int(nominal.text)
                        date_str = root.attrib.get('Date', '')
                        logger.info("Курс получен из ЦБ РФ: 1 CNY = %.4f RUB (дата: %s)", rate, date_str)
                        return round(rate, 4), date_str
            
            logger.error("Курс CNY не найден в ответе ЦБ РФ")
            
    except Exception as e:
        logger.error("Ошибка при получении курса от ЦБ РФ: %s", e)
    
    # Fallback: пробуем Open ER API
    logger.debug("Пробуем резервный источник - Open ER API...")
    try:
        url = "https://open.er-api.com/v6/latest/CNY"
        response = requests.get(url, timeout=5)
        if response.status_code == 200:
            data = response.json()
            rates = data.get('rates', {})
            rate = rates.get('RUB')
            date_str = data.get('time_last_update_utc', '')
            if rate:
                logger.info("Курс получен из Open ER API: 1 CNY = %.4f RUB (дата: %s)", rate, date_str)
                return round(rate, 4), date_str
    except Exception as e:
        logger.error("Ошибка при получении курса от Open ER API: %s", e)
    
    logger.error("Не удалось получить курс из всех источников")
    return None, None


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
    if datetime.now() - cached_at > timedelta(minutes=_CACHE_TTL_MINUTES):
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


def get_exchange_rate_info():
    """
    Получает информацию о курсе валют с кэшированием
    """
    global _cache, cached_at
    
    # Проверяем, есть ли кэш
    if _cache is not None and cached_at is not None:
        rate, date_str = _cache
        
        # Проверяем, не устарел ли кэш
        cache_age = datetime.now() - cached_at
        
        # Если кэш свежий (меньше 10 минут), возвращаем его БЕЗ запроса к API
        if cache_age < timedelta(minutes=_CACHE_TTL_MINUTES):
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
            cache_age_minutes = cache_age.total_seconds() / 60
            logger.info("Кэш устарел (возраст: %.1f мин), обновляем из API...", cache_age_minutes)
    
    # Кэша нет или он устарел - получаем новый курс
    rate, date_str = get_currency_rate()
    
    if rate is not None:
        # Обновляем кэш
        _cache = (rate, date_str, datetime.now())
        cached_at = _cache[2]
        
        return {
            'rate': round(rate, 4),
            'date': date_str,
            'source': 'ЦБ РФ + резервный API',
            'cached': False,
            'cached_at': cached_at.isoformat(),
            'cache_age_minutes': 0
        }
    else:
        # Не удалось получить курс
        logger.error("Не удалось получить курс валют")
        return {
            'rate': 11.05,  # Значение по умолчанию
            'date': datetime.now().strftime('%d.%m.%Y'),
            'source': 'default',
            'cached': False,
            'cached_at': None,
            'cache_age_minutes': None
        }
    

def clear_cache() -> None:
    """
    Очищает кэш курса валют (для тестирования или принудительного обновления).
    """
    global _cache
    _cache = None
    logger.info("Кэш курса валют очищен")


def get_cached_exchange_rate():
    """
    Получает курс валют с автоматическим обновлением устаревшего кэша
    """
    global _cache, cached_at
    
    if _cache is None or cached_at is None:
        # Нет кэша - получаем курс
        result = get_currency_rate()
        if result[0] is not None:
            rate, date_str = result
            _cache = (rate, date_str, datetime.now())
            cached_at = _cache[2]
            logger.info("Курс валют получен и закэширован: %.4f RUB", rate)
            return rate
        else:
            logger.error("Не удалось получить курс валют")
            return 11.05  # Значение по умолчанию
    
    # Проверяем, не устарел ли кэш
    rate, date_str, cache_time = _cache
    if datetime.now() - cache_time > timedelta(minutes=_CACHE_TTL_MINUTES):
        # Кэш устарел, пытаемся обновить
        result = get_currency_rate()
        if result[0] is not None:
            new_rate, new_date_str = result
            _cache = (new_rate, new_date_str, datetime.now())
            logger.info("Курс валют обновлен из API: 1 CNY = %.4f RUB (дата: %s)", new_rate, new_date_str)
            return new_rate
        else:
            # Не удалось обновить, используем старый кэш
            logger.warning("Не удалось обновить курс из API, используется кэшированное значение: %.2f", rate)
            return rate
    else:
        # Кэш свежий
        cache_age_minutes = int((datetime.now() - cache_time).total_seconds() / 60)
        logger.debug("Используется свежий кэш курса: %.4f RUB (возраст: %d мин)", rate, cache_age_minutes)
        return rate