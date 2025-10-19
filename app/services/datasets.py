from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional
from uuid import uuid4

from flask import current_app


BASE_DIR = Path(__file__).resolve().parents[2]
CONFIG_DIR = BASE_DIR / 'config'
VERSIONS_DIR = CONFIG_DIR / 'versions'

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


def save_logistics_cities(cities: List[Dict[str, Any]], *, actor: Optional[str] = None):
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

    _snapshot_version('logistics_cities', payload, actor)

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
    refresh_orders_documents()
    refresh_task_templates()
    refresh_task_instructions()


def _snapshot_version(collection: str, payload: Dict[str, Any], actor: Optional[str] = None):
    """Сохраняет резервную копию данных в каталоге версий."""
    if not payload:
        return

    timestamp = datetime.utcnow().strftime('%Y%m%dT%H%M%S')
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
