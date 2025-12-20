from __future__ import annotations

import csv
import json
import os
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional
from uuid import uuid4

from flask import current_app


BASE_DIR = Path(__file__).resolve().parents[2]
CONFIG_DIR = BASE_DIR / 'config'
VERSIONS_DIR = CONFIG_DIR / 'versions'

GB_ANALOGS: List[Dict[str, Any]] = []
DUTY_RATES: List[Dict[str, Any]] = []
TNVED_CATALOG: List[Dict[str, Any]] = []
ORDERS_REGISTRY: List[Dict[str, Any]] = []
TASK_TEMPLATES: List[Dict[str, Any]] = []
TASK_INSTRUCTIONS: List[Dict[str, Any]] = []


def _log_error(message: str):
    """Пишет сообщение об ошибке в лог приложения, если он доступен."""
    logger = getattr(current_app, 'logger', None)
    if logger:
        logger.error(message)


def load_gb_materials() -> List[Dict[str, Any]]:
    """Читает аналоги материалов из конфигурационного JSON."""
    materials_path = CONFIG_DIR / 'gb_materials.json'
    try:
        with materials_path.open('r', encoding='utf-8') as file:
            data = json.load(file)
        materials = data.get('materials', [])

        for material in materials:
            composition = material.get('composition', [])
            material['composition_search'] = ' '.join(
                f"{item.get('element', '')} {item.get('content', '')}"
                for item in composition
            )

        return materials
    except FileNotFoundError:
        _log_error(f'GB materials file not found at {materials_path.as_posix()}')
    except json.JSONDecodeError as exc:
        _log_error(f'Failed to parse GB materials file: {exc}')
    return []


def save_gb_materials(materials: List[Dict[str, Any]], *, actor: Optional[str] = None):
    """Сохраняет список аналогов материалов обратно в файл."""
    materials_path = CONFIG_DIR / 'gb_materials.json'
    materials_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        'materials': [
            {
                'russian': material.get('russian', ''),
                'gb': material.get('gb', ''),
                'notes': material.get('notes', ''),
                'composition': [
                    {
                        'element': component.get('element', ''),
                        'content': component.get('content', '')
                    }
                    for component in material.get('composition', [])
                ]
            }
            for material in materials
        ]
    }
    _snapshot_version('gb_materials', payload, actor)

    with materials_path.open('w', encoding='utf-8') as file:
        json.dump(payload, file, ensure_ascii=False, indent=2)


def refresh_gb_analogs():
    """Перечитывает аналоги материалов в память для быстрого доступа."""
    global GB_ANALOGS
    GB_ANALOGS = load_gb_materials()


def get_gb_materials() -> List[Dict[str, Any]]:
    """Возвращает копию кэшированного списка аналогов материалов."""
    return list(GB_ANALOGS)


def load_duty_rates() -> List[Dict[str, Any]]:
    """Загружает ставки пошлин из JSON и готовит поля для поиска."""
    duty_path = CONFIG_DIR / 'duty_rates.json'
    try:
        with duty_path.open('r', encoding='utf-8') as file:
            data = json.load(file)
        items = data.get('items', [])

        for item in items:
            item['product_search'] = str(item.get('product', '')).lower()
            item['category_search'] = str(item.get('category', '')).lower()
            item['duty_search'] = str(item.get('duty_percent', '')).lower()

        return items
    except FileNotFoundError:
        _log_error(f'Duty rates file not found at {duty_path.as_posix()}')
    except json.JSONDecodeError as exc:
        _log_error(f'Failed to parse duty rates file: {exc}')
    return []


def _tnved_catalog_path() -> Optional[Path]:
    """Определяет путь к CSV каталогу ТН ВЭД."""
    data_dir = BASE_DIR / 'data'
    candidates = [
        data_dir / 'tnved_catalog.csv',
        CONFIG_DIR / 'tnved_catalog.csv',
        BASE_DIR / 'ТН-ВЭД-ТД-для-менеджеров-с-ключевыми-словами.csv',
    ]

    for candidate in candidates:
        if candidate.exists():
            return candidate

    try:
        match = next(
            (
                path
                for path in BASE_DIR.glob('*ТН*ВЭД*csv')
                if path.is_file()
            ),
            None,
        )
    except OSError:
        match = None

    return match


def _split_keywords(raw_text: str) -> List[str]:
    """Разбивает строку ключевых слов на список значений."""
    if not raw_text:
        return []
    parts = re.split(r'[;,/]|(?:\s{2,})', raw_text)
    keywords = []
    for part in parts:
        cleaned = part.strip()
        if cleaned:
            keywords.append(cleaned)
    if not keywords and raw_text.strip():
        keywords.append(raw_text.strip())
    return keywords


def _extract_percent_value(raw_text: str) -> Optional[float]:
    """Извлекает числовое значение процента из текстового описания."""
    if not raw_text:
        return None
    match = re.search(r'(\d+(?:[.,]\d+)?)\s*%', raw_text)
    if not match:
        return None
    value = match.group(1).replace(',', '.')
    try:
        return float(value)
    except ValueError:
        return None


def _format_percent_display(value: Optional[float]) -> str:
    """Формирует строковое представление значения процента."""
    if value is None:
        return '—'
    formatted = f'{value:.2f}'.rstrip('0').rstrip('.')
    return f'{formatted}%'


def load_tnved_catalog() -> List[Dict[str, Any]]:
    """Загружает расширенный каталог ставок пошлин из CSV."""
    catalog_path = _tnved_catalog_path()
    if catalog_path is None:
        _log_error('TNVED catalog file not found.')
        return []

    items: List[Dict[str, Any]] = []
    try:
        with catalog_path.open('r', encoding='utf-8-sig', newline='') as csv_file:
            reader = csv.reader(csv_file)
            for row in reader:
                if not row or not any(cell.strip() for cell in row):
                    continue

                if len(row) < 6:
                    continue

                index_raw = row[0].strip().strip('"')
                if not index_raw or not index_raw.isdigit():
                    continue

                code = row[1].strip()
                description = row[2].strip()
                keywords_text = row[3].strip()
                examples_text = row[4].strip()
                duty_text = row[5].strip() or '—'

                keywords_list = _split_keywords(keywords_text)
                duty_percent = _extract_percent_value(duty_text)

                search_chunks = [
                    code.lower(),
                    keywords_text.lower(),
                    description.lower(),
                    examples_text.lower(),
                    duty_text.lower(),
                ]

                items.append({
                    'code': code,
                    'description': description,
                    'keywords_display': keywords_text,
                    'keywords_list': keywords_list,
                    'examples': examples_text,
                    'duty_text': duty_text,
                    'duty_percent': duty_percent,
                    'title': keywords_text or description or code,
                    'title_search': (keywords_text or description or code).lower(),
                    'description_search': description.lower(),
                    'examples_search': examples_text.lower(),
                    'duty_search': duty_text.lower(),
                    'search_blob': ' '.join(filter(None, search_chunks)),
                    'source': 'tnved',
                })
    except FileNotFoundError:
        _log_error(f'TNVED catalog file not found: {catalog_path}')
    except Exception as exc:  # pragma: no cover - defensive logging
        _log_error(f'Failed to parse TNVED catalog: {exc}')

    return items


def save_duty_rates(items: List[Dict[str, Any]], *, actor: Optional[str] = None):
    """Сохраняет изменённый список ставок пошлин в конфигурационный файл."""
    duty_path = CONFIG_DIR / 'duty_rates.json'
    duty_path.parent.mkdir(parents=True, exist_ok=True)

    def _coerce_float(value):
        try:
            return float(value)
        except (TypeError, ValueError):
            return 0.0

    payload = {
        'items': [
            {
                'product': item.get('product', ''),
                'category': item.get('category', ''),
                'duty_percent': _coerce_float(item.get('duty_percent', 0))
            }
            for item in items
        ]
    }

    _snapshot_version('duty_rates', payload, actor)

    with duty_path.open('w', encoding='utf-8') as file:
        json.dump(payload, file, ensure_ascii=False, indent=2)


def refresh_duty_rates():
    """Обновляет кэш ставок пошлин после изменения файлов."""
    global DUTY_RATES
    DUTY_RATES = load_duty_rates()


def get_duty_rates() -> List[Dict[str, Any]]:
    """Возвращает копию кэшированного списка ставок пошлин."""
    return list(DUTY_RATES)


def refresh_tnved_catalog():
    """Обновляет кэш расширенного каталога ставок пошлин."""
    global TNVED_CATALOG
    TNVED_CATALOG = load_tnved_catalog()


def get_tnved_catalog() -> List[Dict[str, Any]]:
    """Возвращает копию каталога ставок из CSV."""
    return list(TNVED_CATALOG)


def _normalize_manual_duty_items(manual_items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Преобразует ручные записи пошлин в общий формат каталога."""
    normalized: List[Dict[str, Any]] = []
    for item in manual_items:
        product = str(item.get('product', '')).strip()
        category = str(item.get('category', '')).strip()
        duty_percent = item.get('duty_percent')
        duty_text = _format_percent_display(duty_percent if duty_percent is not None else None)
        title = product or category or 'Без названия'
        title_search = title.lower()
        description_search = category.lower()
        duty_search = duty_text.lower()

        normalized.append({
            'code': '',
            'description': category,
            'keywords_display': product,
            'keywords_list': [product] if product else [],
            'examples': '',
            'duty_text': duty_text,
            'duty_percent': duty_percent,
            'title': title,
            'title_search': title_search,
            'description_search': description_search,
            'examples_search': '',
            'duty_search': duty_search,
            'search_blob': ' '.join(filter(None, [title_search, description_search, duty_search])),
            'product': product,
            'category': category,
            'source': 'manual',
        })
    return normalized


def _normalize_tnved_items(catalog_items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Нормализует записи каталога ТН ВЭД (здесь данные уже стандартизированы)."""
    normalized: List[Dict[str, Any]] = []
    for item in catalog_items:
        normalized.append({
            **item,
            'product': item.get('keywords_display', ''),
            'category': item.get('description', ''),
        })
    return normalized


def _duty_catalog_sort_key(item: Dict[str, Any]) -> tuple:
    """Используется для устойчивой сортировки общего каталога."""
    source_priority = 0 if item.get('source') == 'manual' else 1
    code = item.get('code') or ''
    title = item.get('title') or ''
    return (source_priority, code, title.lower())


def get_duty_catalog() -> List[Dict[str, Any]]:
    """Возвращает объединённый каталог пошлин (ручные записи + ТН ВЭД)."""
    manual = _normalize_manual_duty_items(get_duty_rates())
    tnved = _normalize_tnved_items(get_tnved_catalog())
    combined = manual + tnved
    return sorted(combined, key=_duty_catalog_sort_key)


def load_logistics_cities() -> List[Dict[str, Any]]:
    """
    Загружает справочник городов и тарифов логистики.
    Объединяет все три справочника (основные города, ЕКБ+РФ, трал).
    """
    # Всегда используем объединение новых справочников
    # Старый файл logistics_cities.json больше не используется напрямую
    return load_all_logistics_cities()


def save_logistics_cities(cities: List[Dict[str, Any]], *, actor: Optional[str] = None):
    """Сохраняет обновлённый список тарифов логистики в JSON."""
    logistics_path = CONFIG_DIR / 'logistics_cities.json'
    logistics_path.parent.mkdir(parents=True, exist_ok=True)

    def _coerce_float(value):
        try:
            return float(value) if value is not None else None
        except (TypeError, ValueError):
            return None
    
    def _coerce_int(value):
        try:
            return int(value) if value is not None else None
        except (TypeError, ValueError):
            return None

    payload = {
        'cities': [
            {
                'name': city.get('name', ''),
                'region': city.get('region', ''),
                'truck_price': _coerce_float(city.get('truck_price', 0)),
                'trail_price': _coerce_float(city.get('trail_price')),
                'is_main_route': bool(city.get('is_main_route', False)),
                'allows_trail': bool(city.get('allows_trail', False)),
                'distance_from_ekb_km': _coerce_int(city.get('distance_from_ekb_km')),
                **({'main_city': city['main_city']} if 'main_city' in city else {})
            }
            for city in cities
        ]
    }

    _snapshot_version('logistics_cities', payload, actor)

    with logistics_path.open('w', encoding='utf-8') as file:
        json.dump(payload, file, ensure_ascii=False, indent=2)


def load_main_cities() -> List[Dict[str, Any]]:
    """
    Загружает справочник основных городов и городов в радиусе 300км.
    Включает основные города (is_main_route=True) и города с полем main_city.
    """
    main_cities_path = CONFIG_DIR / 'logistics_main_cities.json'
    try:
        with main_cities_path.open('r', encoding='utf-8') as file:
            data = json.load(file)
        return data.get('cities', [])
    except FileNotFoundError:
        _log_error(f'Main cities file not found at {main_cities_path.as_posix()}')
    except json.JSONDecodeError as exc:
        _log_error(f'Failed to parse main cities file: {exc}')
    return []


def save_main_cities(cities: List[Dict[str, Any]], *, actor: Optional[str] = None):
    """Сохраняет справочник основных городов и городов в радиусе 300км."""
    main_cities_path = CONFIG_DIR / 'logistics_main_cities.json'
    main_cities_path.parent.mkdir(parents=True, exist_ok=True)

    def _coerce_float(value):
        try:
            return float(value) if value is not None else None
        except (TypeError, ValueError):
            return None

    payload = {
        'cities': [
            {
                'name': city.get('name', ''),
                'region': city.get('region', ''),
                'truck_price': _coerce_float(city.get('truck_price', 0)),
                'is_main_route': bool(city.get('is_main_route', False)),
                'main_city': city.get('main_city')  # Для городов в радиусе 300км
            }
            for city in cities
        ]
    }

    _snapshot_version('logistics_main_cities', payload, actor)

    with main_cities_path.open('w', encoding='utf-8') as file:
        json.dump(payload, file, ensure_ascii=False, indent=2)


def load_ekb_rf_cities() -> List[Dict[str, Any]]:
    """
    Загружает справочник городов за пределами 300км (используется алгоритм ЕКБ+РФ).
    """
    ekb_rf_path = CONFIG_DIR / 'logistics_ekb_rf_cities.json'
    try:
        with ekb_rf_path.open('r', encoding='utf-8') as file:
            data = json.load(file)
        return data.get('cities', [])
    except FileNotFoundError:
        _log_error(f'EKB+RF cities file not found at {ekb_rf_path.as_posix()}')
    except json.JSONDecodeError as exc:
        _log_error(f'Failed to parse EKB+RF cities file: {exc}')
    return []


def save_ekb_rf_cities(cities: List[Dict[str, Any]], *, actor: Optional[str] = None):
    """Сохраняет справочник городов для алгоритма ЕКБ+РФ."""
    ekb_rf_path = CONFIG_DIR / 'logistics_ekb_rf_cities.json'
    ekb_rf_path.parent.mkdir(parents=True, exist_ok=True)

    def _coerce_int(value):
        try:
            return int(value) if value is not None else None
        except (TypeError, ValueError):
            return None

    payload = {
        'cities': [
            {
                'name': city.get('name', ''),
                'region': city.get('region', ''),
                'distance_from_ekb_km': _coerce_int(city.get('distance_from_ekb_km'))
            }
            for city in cities
        ]
    }

    _snapshot_version('logistics_ekb_rf_cities', payload, actor)

    with ekb_rf_path.open('w', encoding='utf-8') as file:
        json.dump(payload, file, ensure_ascii=False, indent=2)


def load_trail_cities() -> List[Dict[str, Any]]:
    """Загружает справочник городов для трала."""
    trail_path = CONFIG_DIR / 'logistics_trail_cities.json'
    try:
        with trail_path.open('r', encoding='utf-8') as file:
            data = json.load(file)
        return data.get('cities', [])
    except FileNotFoundError:
        _log_error(f'Trail cities file not found at {trail_path.as_posix()}')
    except json.JSONDecodeError as exc:
        _log_error(f'Failed to parse trail cities file: {exc}')
    return []


def save_trail_cities(cities: List[Dict[str, Any]], *, actor: Optional[str] = None):
    """Сохраняет справочник городов для трала."""
    trail_path = CONFIG_DIR / 'logistics_trail_cities.json'
    trail_path.parent.mkdir(parents=True, exist_ok=True)

    def _coerce_float(value):
        try:
            return float(value) if value is not None else None
        except (TypeError, ValueError):
            return None

    payload = {
        'cities': [
            {
                'name': city.get('name', ''),
                'region': city.get('region', ''),
                'trail_price': _coerce_float(city.get('trail_price', 0))
            }
            for city in cities
        ]
    }

    _snapshot_version('logistics_trail_cities', payload, actor)

    with trail_path.open('w', encoding='utf-8') as file:
        json.dump(payload, file, ensure_ascii=False, indent=2)


def load_all_logistics_cities() -> List[Dict[str, Any]]:
    """
    Объединяет все три справочника логистики в один список для обратной совместимости.
    Используется в основном приложении для работы с логистикой.
    """
    all_cities = []
    
    # Основные города
    main_cities = load_main_cities()
    for city in main_cities:
        all_cities.append({
            'name': city.get('name', ''),
            'region': city.get('region', ''),
            'truck_price': city.get('truck_price', 0),
            'trail_price': None,
            'is_main_route': city.get('is_main_route', False),
            'allows_trail': False,
            'distance_from_ekb_km': None,
            **({'main_city': city['main_city']} if 'main_city' in city else {})
        })
    
    # ЕКБ+РФ города
    ekb_rf_cities = load_ekb_rf_cities()
    for city in ekb_rf_cities:
        all_cities.append({
            'name': city.get('name', ''),
            'region': city.get('region', ''),
            'truck_price': 0,  # Для ЕКБ+РФ используется алгоритм, а не прямая цена
            'trail_price': None,
            'is_main_route': False,
            'allows_trail': False,
            'distance_from_ekb_km': city.get('distance_from_ekb_km')
        })
    
    # Трал города
    trail_cities = load_trail_cities()
    for city in trail_cities:
        all_cities.append({
            'name': city.get('name', ''),
            'region': city.get('region', ''),
            'truck_price': 0,  # Для трала используется отдельная цена
            'trail_price': city.get('trail_price', 0),
            'is_main_route': False,
            'allows_trail': True,
            'distance_from_ekb_km': None
        })
    
    return all_cities


def is_city_in_ekb_rf_catalog(city_name: str) -> bool:
    """
    Проверяет, находится ли город в справочнике ЕКБ+РФ.
    
    Args:
        city_name: Название города для проверки
    
    Returns:
        True если город найден в справочнике ЕКБ+РФ, False иначе
    """
    ekb_rf_cities = load_ekb_rf_cities()
    for city in ekb_rf_cities:
        if city.get('name', '').strip() == city_name.strip():
            return True
    return False


def get_ekb_rf_city_distance(city_name: str) -> Optional[int]:
    """
    Получает расстояние от ЕКБ для города из справочника ЕКБ+РФ.
    
    Args:
        city_name: Название города
    
    Returns:
        Расстояние от ЕКБ в километрах или None если город не найден
    """
    ekb_rf_cities = load_ekb_rf_cities()
    for city in ekb_rf_cities:
        if city.get('name', '').strip() == city_name.strip():
            return city.get('distance_from_ekb_km')
    return None


def update_city_distance(city_name: str, distance_km: int, actor: Optional[str] = None) -> bool:
    """
    Обновляет расстояние от ЕКБ для указанного города.
    Ищет город во всех трех справочниках и обновляет в соответствующем.
    
    Args:
        city_name: Название города
        distance_km: Расстояние от Екатеринбурга в километрах
        actor: Кто вносит изменение (username или None)
    
    Returns:
        True если город найден и обновлен, False иначе
    """
    # Ищем в справочнике ЕКБ+РФ (там хранится distance_from_ekb_km)
    ekb_rf_cities = load_ekb_rf_cities()
    for city in ekb_rf_cities:
        if city.get('name') == city_name:
            city['distance_from_ekb_km'] = distance_km
            save_ekb_rf_cities(ekb_rf_cities, actor=actor)
            return True
    
    # Если город не найден в ЕКБ+РФ, но есть расстояние, возможно нужно переместить город
    # или добавить его в справочник ЕКБ+РФ. Для простоты просто возвращаем False.
    return False


def load_orders_documents() -> List[Dict[str, Any]]:
    """Читает список распоряжений и нормализует структуру файлов."""
    orders_path = CONFIG_DIR / 'orders_documents.json'
    try:
        with orders_path.open('r', encoding='utf-8') as file:
            data = json.load(file)
    except FileNotFoundError:
        _log_error(f'Orders registry file not found at {orders_path.as_posix()}')
        return []
    except json.JSONDecodeError as exc:
        _log_error(f'Failed to parse orders registry file: {exc}')
        return []

    normalized_orders: List[Dict[str, Any]] = []
    for entry in data.get('orders', []):
        files: List[Dict[str, str]] = []
        for file_entry in entry.get('files', []):
            filename = str(file_entry.get('filename', '')).strip()
            if not filename:
                continue
            label = str(
                file_entry.get('label')
                or file_entry.get('format')
                or file_entry.get('name')
                or Path(filename).suffix.replace('.', '').upper()
            ).strip()
            files.append({
                'label': label or 'Скачать',
                'filename': filename
            })

        normalized_orders.append({
            'id': entry.get('id') or entry.get('identifier'),
            'title': entry.get('title') or entry.get('name') or 'Распоряжение',
            'brief': str(entry.get('brief', '')).strip(),
            'summary': str(entry.get('summary', '')).strip(),
            'files': files,
            'updated_at': entry.get('updated_at') or entry.get('date')
        })

    return normalized_orders


def refresh_orders_documents():
    """Обновляет кэш распоряжений из конфигурационного файла."""
    global ORDERS_REGISTRY
    ORDERS_REGISTRY = load_orders_documents()


def get_orders_documents() -> List[Dict[str, Any]]:
    """Возвращает копию списка распоряжений."""
    return list(ORDERS_REGISTRY)


def load_task_templates() -> List[Dict[str, Any]]:
    """Читает список шаблонов задач из конфигурационного файла."""
    templates_path = CONFIG_DIR / 'task_templates.json'
    try:
        with templates_path.open('r', encoding='utf-8') as file:
            data = json.load(file)
    except FileNotFoundError:
        _log_error(f'Task templates file not found at {templates_path.as_posix()}')
        return []
    except json.JSONDecodeError as exc:
        _log_error(f'Failed to parse task templates file: {exc}')
        return []

    normalized_templates: List[Dict[str, Any]] = []
    for entry in data.get('templates', []):
        files: List[Dict[str, str]] = []
        for file_entry in entry.get('files', []):
            filename = str(file_entry.get('filename', '')).strip()
            if not filename:
                continue
            label = str(
                file_entry.get('label')
                or file_entry.get('format')
                or file_entry.get('name')
                or Path(filename).suffix.replace('.', '').upper()
            ).strip()
            files.append({
                'label': label or 'Скачать',
                'filename': filename
            })

        normalized_templates.append({
            'id': entry.get('id') or entry.get('identifier'),
            'title': entry.get('title') or entry.get('name') or 'Шаблон',
            'summary': str(entry.get('summary', '')).strip(),
            'files': files,
            'updated_at': entry.get('updated_at') or entry.get('date')
        })

    return normalized_templates


def refresh_task_templates():
    """Обновляет кэш шаблонов задач."""
    global TASK_TEMPLATES
    TASK_TEMPLATES = load_task_templates()


def get_task_templates() -> List[Dict[str, Any]]:
    """Возвращает копию списка шаблонов задач."""
    return list(TASK_TEMPLATES)


def load_task_instructions() -> List[Dict[str, Any]]:
    """Читает список инструкций из конфигурационного файла."""
    instructions_path = CONFIG_DIR / 'instructions_tasks.json'
    try:
        with instructions_path.open('r', encoding='utf-8') as file:
            data = json.load(file)
    except FileNotFoundError:
        _log_error(f'Instructions file not found at {instructions_path.as_posix()}')
        return []
    except json.JSONDecodeError as exc:
        _log_error(f'Failed to parse instructions file: {exc}')
        return []

    normalized_instructions: List[Dict[str, Any]] = []
    for entry in data.get('instructions', []):
        files: List[Dict[str, str]] = []
        
        for file_entry in entry.get('files', []):
            filename = str(file_entry.get('filename', '')).strip()
            if not filename:
                continue
            label = str(
                file_entry.get('label')
                or file_entry.get('format')
                or file_entry.get('name')
                or Path(filename).suffix.replace('.', '').upper()
            ).strip()
            files.append({
                'label': label or 'Скачать',
                'filename': filename
            })

        normalized_instructions.append({
            'id': entry.get('id') or entry.get('identifier'),
            'title': entry.get('title') or entry.get('name') or 'Инструкция',
            'brief': str(entry.get('brief', '')).strip(),
            'summary': str(entry.get('summary', '')).strip(),
            'files': files,
            'updated_at': entry.get('updated_at') or entry.get('date')
        })

    return normalized_instructions


def refresh_task_instructions():
    """Обновляет кэш инструкций."""
    global TASK_INSTRUCTIONS
    TASK_INSTRUCTIONS = load_task_instructions()


def get_task_instructions() -> List[Dict[str, Any]]:
    """
    Возвращает список инструкций с содержимым TXT файлов.
    Содержимое файлов загружается каждый раз при запросе (не кэшируется),
    чтобы изменения в TXT файлах сразу отображались на сайте.
    """
    instructions_dir = BASE_DIR / 'static' / 'instructions'
    result = []
    
    for instruction in TASK_INSTRUCTIONS:
        # Копируем инструкцию
        instruction_copy = dict(instruction)
        
        # Загружаем содержимое TXT файлов каждый раз при запросе
        content_text: str | None = None
        for file_entry in instruction.get('files', []):
            filename = str(file_entry.get('filename', '')).strip()
            if filename.lower().endswith('.txt'):
                file_path = instructions_dir / filename
                if file_path.exists():
                    try:
                        with file_path.open('r', encoding='utf-8') as txt_file:
                            content_text = txt_file.read()
                            break  # Берем первый найденный TXT файл
                    except Exception as exc:
                        _log_error(f'Failed to read instruction file {filename}: {exc}')
        
        instruction_copy['content_text'] = content_text
        result.append(instruction_copy)
    
    return result


def parse_composition_input(raw_text: str):
    """Преобразует текстовое описание состава материала в структуру данных."""
    if not raw_text:
        return []

    composition = []
    for line in raw_text.splitlines():
        cleaned = line.strip()
        if not cleaned:
            continue
        if ':' in cleaned:
            element, content = cleaned.split(':', 1)
        elif '=' in cleaned:
            element, content = cleaned.split('=', 1)
        else:
            parts = cleaned.split(maxsplit=1)
            element = parts[0]
            content = parts[1] if len(parts) > 1 else ''
        composition.append({'element': element.strip(), 'content': content.strip()})

    return composition


def save_orders_documents(orders: List[Dict[str, Any]], *, actor: Optional[str] = None):
    """Сохраняет список распоряжений и обновляет кэш."""
    orders_path = CONFIG_DIR / 'orders_documents.json'
    orders_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {'orders': orders}
    _snapshot_version('orders_documents', payload, actor)

    with orders_path.open('w', encoding='utf-8') as file:
        json.dump(payload, file, ensure_ascii=False, indent=2)

    refresh_orders_documents()


def save_task_templates(templates: List[Dict[str, Any]], *, actor: Optional[str] = None):
    """Сохраняет перечень шаблонов и обновляет кэш."""
    templates_path = CONFIG_DIR / 'task_templates.json'
    templates_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {'templates': templates}
    _snapshot_version('task_templates', payload, actor)

    with templates_path.open('w', encoding='utf-8') as file:
        json.dump(payload, file, ensure_ascii=False, indent=2)

    refresh_task_templates()


def save_task_instructions(instructions: List[Dict[str, Any]], *, actor: Optional[str] = None):
    """Сохраняет список инструкций и обновляет кэш."""
    instructions_path = CONFIG_DIR / 'instructions_tasks.json'
    instructions_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {'instructions': instructions}
    _snapshot_version('instructions_tasks', payload, actor)

    with instructions_path.open('w', encoding='utf-8') as file:
        json.dump(payload, file, ensure_ascii=False, indent=2)

    refresh_task_instructions()


def save_with_version(collection: str, data: Dict[str, Any], *, actor: Optional[str] = None):
    """Обёртка для сохранения произвольных данных с версионированием."""
    _snapshot_version(collection, data, actor)


def parse_files_input(raw_text: str) -> List[Dict[str, str]]:
    """Преобразует текстовое представление файлов контента в структуру."""
    if not raw_text:
        return []

    files: List[Dict[str, str]] = []
    for line in raw_text.splitlines():
        cleaned = line.strip()
        if not cleaned:
            continue

        if '|' in cleaned:
            label, filename = cleaned.split('|', 1)
        elif ';' in cleaned:
            label, filename = cleaned.split(';', 1)
        else:
            parts = cleaned.split(maxsplit=1)
            label = parts[0]
            filename = parts[1] if len(parts) > 1 else parts[0]

        files.append({'label': label.strip(), 'filename': filename.strip()})

    return files


def format_files_output(files: List[Dict[str, Any]]) -> str:
    """Формирует читабельное представление файлов для textarea."""
    lines = []
    for file_item in files:
        label = str(file_item.get('label', '')).strip()
        filename = str(file_item.get('filename', '')).strip()
        if label and filename:
            lines.append(f'{label} | {filename}')
        elif filename:
            lines.append(filename)
    return '\n'.join(lines)


def make_content_id(prefix: str) -> str:
    """Создаёт уникальный идентификатор контента."""
    return f'{prefix}-{uuid4().hex[:8]}'


def list_content_versions(collection: str) -> List[Dict[str, Any]]:
    """Возвращает список версий для указанной коллекции."""
    versions_path = VERSIONS_DIR / collection
    if not versions_path.exists():
        return []

    versions: List[Dict[str, Any]] = []
    for entry in sorted(versions_path.iterdir(), reverse=True):
        if entry.suffix != '.json' or not entry.is_file():
            continue
        try:
            with entry.open('r', encoding='utf-8') as file:
                payload = json.load(file)
        except (OSError, json.JSONDecodeError):
            continue

        versions.append({
            'filename': entry.name,
            'saved_at': payload.get('saved_at'),
            'actor': payload.get('actor'),
        })

    return versions


def load_content_version(collection: str, filename: str) -> Optional[Dict[str, Any]]:
    """Загружает выбранную версию контента."""
    target = VERSIONS_DIR / collection / filename
    if not target.exists():
        return None
    try:
        with target.open('r', encoding='utf-8') as file:
            return json.load(file)
    except (OSError, json.JSONDecodeError):
        return None


def init_app(_app):
    """Инициализирует кэшированные данные при старте приложения."""
    refresh_gb_analogs()
    refresh_duty_rates()
    refresh_tnved_catalog()
    refresh_orders_documents()
    refresh_task_templates()
    refresh_task_instructions()


def _snapshot_version(collection: str, payload: Dict[str, Any], actor: Optional[str] = None):
    """Сохраняет резервную копию данных в каталоге версий."""
    if not payload:
        return

    timestamp = datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S')
    versions_path = VERSIONS_DIR / collection
    versions_path.mkdir(parents=True, exist_ok=True)

    filename = f'{timestamp}.json'
    meta = {
        'saved_at': timestamp,
        'actor': actor,
        'data': payload,
    }

    try:
        with (versions_path / filename).open('w', encoding='utf-8') as file:
            json.dump(meta, file, ensure_ascii=False, indent=2)
    except OSError:
        _log_error(f'Failed to persist version snapshot for {collection}')
