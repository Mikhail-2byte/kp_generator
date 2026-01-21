"""Тесты для работы с каталогом ТН ВЭД в JSON формате."""

import json
import sys
from pathlib import Path
from tempfile import TemporaryDirectory

# Добавляем корень проекта в путь для возможности прямого запуска
project_root = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(project_root))

import pytest

from app.services import datasets


@pytest.fixture
def sample_tnved_items():
    """Создает тестовые данные каталога ТН ВЭД."""
    return [
        {
            'code': '9031908500',
            'description': 'ЧАСТИ И ПРИНАДЛЕЖНОСТИ ИЗМЕРИТЕЛЬНЫХ ИЛИ КОНТРОЛЬНЫХ ПРИБОРОВ',
            'keywords_display': 'ШТОК (КОНТРОЛЬ И ИЗМЕРЕНИЕ ПРОЧНОСТНЫХ ПАРАМЕТРОВ)',
            'examples': 'ШТОК',
            'duty_text': '0%',
            'duty_percent': 0.0
        },
        {
            'code': '8709900000',
            'description': 'ЧАСТИ ТРАНСПОРТНЫХ СРЕДСТВ ПРОМЫШЛЕННОГО НАЗНАЧЕНИЯ',
            'keywords_display': 'ОСЬ ТОРМОЗНАЯ',
            'examples': 'ОСЬ ТОРМОЗНАЯ',
            'duty_text': '5%',
            'duty_percent': 5.0
        },
        {
            'code': '8505200000',
            'description': 'ЭЛЕКТРОМАГНИТНЫЕ СЦЕПЛЕНИЯ, МУФТЫ И ТОРМОЗА',
            'keywords_display': 'ЭЛЕКТРОМАГНИТНАЯ МУФТА, ТОРМОЗ',
            'examples': 'ЭЛЕКТРОМАГНИТНАЯ МУФТА, ТОРМОЗ',
            'duty_text': '10%',
            'duty_percent': 10.0
        }
    ]


def test_save_tnved_catalog(tmp_path: Path, sample_tnved_items, app, monkeypatch):
    """Тест сохранения каталога ТН ВЭД в JSON."""
    config_dir = tmp_path / 'config'
    config_dir.mkdir()
    
    original_config_dir = datasets.CONFIG_DIR
    monkeypatch.setattr(datasets, 'CONFIG_DIR', config_dir)
    
    try:
        with app.app_context():
            datasets.save_tnved_catalog(sample_tnved_items, actor='test_user')
        
        # Проверяем, что файл создан
        tnved_file = config_dir / 'tnved_catalog.json'
        assert tnved_file.exists()
        
        # Проверяем содержимое файла
        with tnved_file.open('r', encoding='utf-8') as f:
            data = json.load(f)
        
        assert 'items' in data
        assert len(data['items']) == 3
        
        # Проверяем структуру первой записи
        item = data['items'][0]
        assert item['code'] == '9031908500'
        assert item['description'] == 'ЧАСТИ И ПРИНАДЛЕЖНОСТИ ИЗМЕРИТЕЛЬНЫХ ИЛИ КОНТРОЛЬНЫХ ПРИБОРОВ'
        assert item['keywords_display'] == 'ШТОК (КОНТРОЛЬ И ИЗМЕРЕНИЕ ПРОЧНОСТНЫХ ПАРАМЕТРОВ)'
        assert item['examples'] == 'ШТОК'
        assert item['duty_text'] == '0%'
        assert item['duty_percent'] == 0.0
        
    finally:
        monkeypatch.setattr(datasets, 'CONFIG_DIR', original_config_dir)


def test_load_tnved_catalog_from_json(tmp_path: Path, sample_tnved_items, app, monkeypatch):
    """Тест загрузки каталога ТН ВЭД из JSON."""
    config_dir = tmp_path / 'config'
    config_dir.mkdir()
    
    original_config_dir = datasets.CONFIG_DIR
    monkeypatch.setattr(datasets, 'CONFIG_DIR', config_dir)
    
    try:
        # Создаем JSON файл
        tnved_file = config_dir / 'tnved_catalog.json'
        payload = {
            'items': [
                {
                    'code': item['code'],
                    'description': item['description'],
                    'keywords_display': item['keywords_display'],
                    'examples': item['examples'],
                    'duty_text': item['duty_text'],
                    'duty_percent': item['duty_percent']
                }
                for item in sample_tnved_items
            ]
        }
        with tnved_file.open('w', encoding='utf-8') as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        
        # Очищаем кэш
        datasets.load_tnved_catalog.cache_clear()
        
        with app.app_context():
            items = datasets.load_tnved_catalog()
        
        assert len(items) == 3
        
        # Проверяем, что данные загружены корректно
        first_item = items[0]
        assert first_item['code'] == '9031908500'
        assert first_item['description'] == 'ЧАСТИ И ПРИНАДЛЕЖНОСТИ ИЗМЕРИТЕЛЬНЫХ ИЛИ КОНТРОЛЬНЫХ ПРИБОРОВ'
        assert first_item['keywords_display'] == 'ШТОК (КОНТРОЛЬ И ИЗМЕРЕНИЕ ПРОЧНОСТНЫХ ПАРАМЕТРОВ)'
        assert first_item['duty_percent'] == 0.0
        assert first_item['source'] == 'tnved'
        
    finally:
        monkeypatch.setattr(datasets, 'CONFIG_DIR', original_config_dir)


def test_save_tnved_catalog_with_none_duty_percent(tmp_path: Path, app, monkeypatch):
    """Тест сохранения каталога с None в duty_percent."""
    config_dir = tmp_path / 'config'
    config_dir.mkdir()
    
    original_config_dir = datasets.CONFIG_DIR
    monkeypatch.setattr(datasets, 'CONFIG_DIR', config_dir)
    
    try:
        items = [
            {
                'code': '1234567890',
                'description': 'Тестовое описание',
                'keywords_display': 'Ключевые слова',
                'examples': 'Примеры',
                'duty_text': 'Не указано',
                'duty_percent': None
            }
        ]
        
        with app.app_context():
            datasets.save_tnved_catalog(items, actor='test_user')
        
        tnved_file = config_dir / 'tnved_catalog.json'
        assert tnved_file.exists()
        
        with tnved_file.open('r', encoding='utf-8') as f:
            data = json.load(f)
        
        assert data['items'][0]['duty_percent'] is None
        
    finally:
        monkeypatch.setattr(datasets, 'CONFIG_DIR', original_config_dir)


def test_save_tnved_catalog_empty_list(tmp_path: Path, app, monkeypatch):
    """Тест сохранения пустого каталога."""
    config_dir = tmp_path / 'config'
    config_dir.mkdir()
    
    original_config_dir = datasets.CONFIG_DIR
    monkeypatch.setattr(datasets, 'CONFIG_DIR', config_dir)
    
    try:
        with app.app_context():
            datasets.save_tnved_catalog([], actor='test_user')
        
        tnved_file = config_dir / 'tnved_catalog.json'
        assert tnved_file.exists()
        
        with tnved_file.open('r', encoding='utf-8') as f:
            data = json.load(f)
        
        assert data['items'] == []
        
    finally:
        monkeypatch.setattr(datasets, 'CONFIG_DIR', original_config_dir)


if __name__ == "__main__":
    pytest.main([__file__, "-v"])

