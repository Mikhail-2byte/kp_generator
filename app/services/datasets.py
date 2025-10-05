from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, Dict, List

from flask import current_app


BASE_DIR = Path(__file__).resolve().parents[2]
CONFIG_DIR = BASE_DIR / 'config'

GB_ANALOGS: List[Dict[str, Any]] = []
DUTY_RATES: List[Dict[str, Any]] = []
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


def save_gb_materials(materials: List[Dict[str, Any]]):
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


def save_duty_rates(items: List[Dict[str, Any]]):
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

    with duty_path.open('w', encoding='utf-8') as file:
        json.dump(payload, file, ensure_ascii=False, indent=2)


def refresh_duty_rates():
    """Обновляет кэш ставок пошлин после изменения файлов."""
    global DUTY_RATES
    DUTY_RATES = load_duty_rates()


def get_duty_rates() -> List[Dict[str, Any]]:
    """Возвращает копию кэшированного списка ставок пошлин."""
    return list(DUTY_RATES)


def load_logistics_cities() -> List[Dict[str, Any]]:
    """Загружает справочник городов и тарифов логистики."""
    logistics_path = CONFIG_DIR / 'logistics_cities.json'
    try:
        with logistics_path.open('r', encoding='utf-8') as file:
            data = json.load(file)
        return data.get('cities', [])
    except FileNotFoundError:
        _log_error(f'Logistics file not found at {logistics_path.as_posix()}')
    except json.JSONDecodeError as exc:
        _log_error(f'Failed to parse logistics file: {exc}')
    return []


def save_logistics_cities(cities: List[Dict[str, Any]]):
    """Сохраняет обновлённый список тарифов логистики в JSON."""
    logistics_path = CONFIG_DIR / 'logistics_cities.json'
    logistics_path.parent.mkdir(parents=True, exist_ok=True)

    def _coerce_float(value):
        try:
            return float(value)
        except (TypeError, ValueError):
            return 0.0

    payload = {
        'cities': [
            {
                'name': city.get('name', ''),
                'region': city.get('region', ''),
                'truck_price': _coerce_float(city.get('truck_price', 0)),
                'trail_price': _coerce_float(city.get('trail_price', 0))
            }
            for city in cities
        ]
    }

    with logistics_path.open('w', encoding='utf-8') as file:
        json.dump(payload, file, ensure_ascii=False, indent=2)


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
    """Возвращает копию списка инструкций."""
    return list(TASK_INSTRUCTIONS)


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


def init_app(_app):
    """Инициализирует кэшированные данные при старте приложения."""
    refresh_gb_analogs()
    refresh_duty_rates()
    refresh_orders_documents()
    refresh_task_templates()
    refresh_task_instructions()
