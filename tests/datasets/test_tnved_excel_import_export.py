"""Тесты для экспорта и импорта каталога ТН ВЭД в Excel."""

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
            'description': 'ЧАСТИ ТРАНСПОРТНЫХ СРЕДСТВ',
            'keywords_display': 'ОСЬ ТОРМОЗНАЯ',
            'examples': 'ОСЬ ТОРМОЗНАЯ',
            'duty_text': '5%',
            'duty_percent': 5.0
        }
    ]


def test_export_tnved_catalog_to_excel(tmp_path: Path, sample_tnved_items, app, monkeypatch):
    """Тест экспорта каталога ТН ВЭД в Excel."""
    config_dir = tmp_path / 'config'
    config_dir.mkdir()
    
    original_config_dir = datasets.CONFIG_DIR
    monkeypatch.setattr(datasets, 'CONFIG_DIR', config_dir)
    
    try:
        # Сохраняем тестовые данные
        with app.app_context():
            datasets.save_tnved_catalog(sample_tnved_items, actor='test_user')
            datasets.load_tnved_catalog.cache_clear()
            
            # Экспортируем
            excel_data = datasets.export_tnved_catalog_to_excel()
        
        # Проверяем, что данные не пустые
        assert len(excel_data) > 0
        
        # Сохраняем во временный файл для проверки
        excel_file = tmp_path / 'exported.xlsx'
        excel_file.write_bytes(excel_data)
        
        # Проверяем, что файл создан
        assert excel_file.exists()
        assert excel_file.stat().st_size > 0
        
    finally:
        monkeypatch.setattr(datasets, 'CONFIG_DIR', original_config_dir)


def test_import_tnved_catalog_from_excel(tmp_path: Path, app, monkeypatch):
    """Тест импорта каталога ТН ВЭД из Excel."""
    config_dir = tmp_path / 'config'
    config_dir.mkdir()
    
    original_config_dir = datasets.CONFIG_DIR
    monkeypatch.setattr(datasets, 'CONFIG_DIR', config_dir)
    
    try:
        # Создаем тестовые данные и экспортируем их
        sample_items = [
            {
                'code': '9031908500',
                'description': 'ЧАСТИ И ПРИНАДЛЕЖНОСТИ',
                'keywords_display': 'ШТОК',
                'examples': 'ШТОК',
                'duty_text': '0%',
                'duty_percent': 0.0
            }
        ]
        
        with app.app_context():
            datasets.save_tnved_catalog(sample_items, actor='test_user')
            # Очищаем кеш для всех функций, которые могут использоваться
            datasets.load_tnved_catalog.cache_clear()
            datasets.load_duty_rates.cache_clear()
            
            # Экспортируем
            excel_data = datasets.export_tnved_catalog_to_excel()
        
        # Сохраняем Excel файл
        excel_file = tmp_path / 'test_import.xlsx'
        excel_file.write_bytes(excel_data)
        
        # Импортируем обратно
        with app.app_context():
            # Очищаем кеш перед импортом
            datasets.load_tnved_catalog.cache_clear()
            datasets.load_duty_rates.cache_clear()
            count = datasets.import_tnved_catalog_from_excel(excel_file, actor='test_user')
        
        assert count == 1
        
        # Проверяем, что данные загружены
        with app.app_context():
            # Очищаем кеш перед проверкой
            datasets.load_tnved_catalog.cache_clear()
            datasets.load_duty_rates.cache_clear()
            items = datasets.load_tnved_catalog()
        
        assert len(items) == 1
        assert items[0]['code'] == '9031908500'
        assert items[0]['description'] == 'ЧАСТИ И ПРИНАДЛЕЖНОСТИ'
        
    finally:
        monkeypatch.setattr(datasets, 'CONFIG_DIR', original_config_dir)


def test_import_tnved_catalog_from_excel_with_duty_percent(tmp_path: Path, app, monkeypatch):
    """Тест импорта каталога с числовым значением пошлины."""
    config_dir = tmp_path / 'config'
    config_dir.mkdir()
    
    original_config_dir = datasets.CONFIG_DIR
    monkeypatch.setattr(datasets, 'CONFIG_DIR', config_dir)
    
    try:
        # Создаем Excel файл вручную с помощью openpyxl
        try:
            from openpyxl import Workbook
        except ImportError:
            pytest.skip("openpyxl не установлен")
        
        wb = Workbook()
        ws = wb.active
        ws.append(['Код ТН ВЭД', 'Описание', 'Ключевые слова', 'Примеры', 'Пошлина', 'Пошлина, %'])
        ws.append(['1234567890', 'Тестовое описание', 'Ключевые слова', 'Примеры', '10%', 10.0])
        
        excel_file = tmp_path / 'test.xlsx'
        wb.save(excel_file)
        
        # Импортируем
        with app.app_context():
            count = datasets.import_tnved_catalog_from_excel(excel_file, actor='test_user')
        
        assert count == 1
        
        # Проверяем данные
        datasets.load_tnved_catalog.cache_clear()
        with app.app_context():
            items = datasets.load_tnved_catalog()
        
        assert items[0]['code'] == '1234567890'
        assert items[0]['duty_percent'] == 10.0
        
    finally:
        monkeypatch.setattr(datasets, 'CONFIG_DIR', original_config_dir)


def test_import_tnved_catalog_from_excel_nonexistent_file(tmp_path: Path, app):
    """Тест импорта из несуществующего файла."""
    nonexistent_file = tmp_path / 'nonexistent.xlsx'
    
    with app.app_context():
        with pytest.raises(FileNotFoundError):
            datasets.import_tnved_catalog_from_excel(nonexistent_file, actor='test_user')


def test_export_tnved_catalog_empty(tmp_path: Path, app, monkeypatch):
    """Тест экспорта пустого каталога."""
    config_dir = tmp_path / 'config'
    config_dir.mkdir()
    
    original_config_dir = datasets.CONFIG_DIR
    monkeypatch.setattr(datasets, 'CONFIG_DIR', config_dir)
    
    try:
        with app.app_context():
            datasets.save_tnved_catalog([], actor='test_user')
            datasets.load_tnved_catalog.cache_clear()
            
            excel_data = datasets.export_tnved_catalog_to_excel()
        
        # Даже пустой каталог должен создавать файл с заголовками
        assert len(excel_data) > 0
        
    finally:
        monkeypatch.setattr(datasets, 'CONFIG_DIR', original_config_dir)


if __name__ == "__main__":
    pytest.main([__file__, "-v"])

