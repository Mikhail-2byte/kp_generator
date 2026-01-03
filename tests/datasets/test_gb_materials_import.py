"""Тесты для импорта GB материалов из CSV."""

import csv
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
def sample_csv_file(tmp_path: Path) -> Path:
    """Создает тестовый CSV файл с данными материалов."""
    csv_path = tmp_path / 'test_steel_prices.csv'
    
    # Создаем CSV файл с правильной структурой
    rows = [
        ['Отчет по материалам [по назначению]', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['Родитель', '', '', '', '', '', '', '', 'Итого'],
        ['№ в группе', 'Материал', '', '', 'Наименование мир', '', 'ГОСТ', 'Материал.Цена', 'Материал.Цена.Вид заготовки', 'Количество'],
        [' 1. Углеродистая (нелегир)', '', '', '', '', '', '', '', '', '115'],
        ['1', '10', '', '', '10', '', 'ГОСТ 1050-2013', '4.30', 'Лист', '1'],
        ['2', '20', '', '', '20', '', 'ГОСТ 1050-2013', '4.30', 'Прокат', '1'],
        ['3', '45', '', '', '45', '', 'ГОСТ 1050-2013', '4.40', 'Труба', '1'],
        [' 2. Низколегированная', '', '', '', '', '', '', '', '', '64'],
        ['1', '09Г2С', '', '', '12Mn', '', 'ГОСТ 19281-2014', '4.00', 'Лист', '1'],
        ['2', '40Х', '', '', '40Cr', '', 'ГОСТ 4543-2016', '4.90', 'Прокат', '1'],
    ]
    
    with csv_path.open('w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerows(rows)
    
    return csv_path


def test_parse_steel_prices_csv(sample_csv_file: Path, app):
    """Тест парсинга CSV файла со сталями."""
    with app.app_context():
        materials = datasets.parse_steel_prices_csv(sample_csv_file)
    
    assert len(materials) == 5  # Должно быть 5 материалов (группирующие строки пропускаются)
    
    # Проверяем первый материал
    material = materials[0]
    assert material['russian'] == '10'
    assert material['gb'] == '10'
    assert material['gost'] == 'ГОСТ 1050-2013'
    assert material['price'] == '4.30'
    assert material['workpiece_type'] == 'Лист'
    assert material['notes'] == ''
    
    # Проверяем последний материал
    last_material = materials[-1]
    assert last_material['russian'] == '40Х'
    assert last_material['gb'] == '40Cr'
    assert last_material['gost'] == 'ГОСТ 4543-2016'
    assert last_material['price'] == '4.90'
    assert last_material['workpiece_type'] == 'Прокат'


def test_parse_steel_prices_csv_empty_file(tmp_path: Path, app):
    """Тест парсинга пустого CSV файла."""
    csv_path = tmp_path / 'empty.csv'
    csv_path.write_text('', encoding='utf-8')
    
    with app.app_context():
        materials = datasets.parse_steel_prices_csv(csv_path)
    assert materials == []


def test_parse_steel_prices_csv_missing_file(tmp_path: Path, app):
    """Тест парсинга несуществующего файла."""
    csv_path = tmp_path / 'nonexistent.csv'
    
    with app.app_context():
        materials = datasets.parse_steel_prices_csv(csv_path)
    assert materials == []


def test_import_gb_materials_from_csv(sample_csv_file: Path, tmp_path: Path, app, monkeypatch):
    """Тест импорта материалов из CSV."""
    # Создаем временный config директорий
    config_dir = tmp_path / 'config'
    config_dir.mkdir()
    
    # Сохраняем оригинальный CONFIG_DIR
    original_config_dir = datasets.CONFIG_DIR
    monkeypatch.setattr(datasets, 'CONFIG_DIR', config_dir)
    
    try:
        # Импортируем материалы
        count = datasets.import_gb_materials_from_csv(sample_csv_file, actor='test_user')
        
        assert count == 5
        
        # Проверяем, что файл создан
        materials_file = config_dir / 'gb_materials.json'
        assert materials_file.exists()
        
        # Проверяем содержимое файла
        with materials_file.open('r', encoding='utf-8') as f:
            data = json.load(f)
        
        assert 'materials' in data
        assert len(data['materials']) == 5
        
        # Проверяем структуру первого материала
        material = data['materials'][0]
        assert 'russian' in material
        assert 'gb' in material
        assert 'gost' in material
        assert 'price' in material
        assert 'workpiece_type' in material
        assert 'notes' in material
        assert 'composition' not in material  # composition должен быть удален
        
        # Проверяем, что данные загружаются корректно
        loaded_materials = datasets.load_gb_materials()
        assert len(loaded_materials) == 5
        
    finally:
        # Восстанавливаем оригинальный CONFIG_DIR
        monkeypatch.setattr(datasets, 'CONFIG_DIR', original_config_dir)


def test_import_gb_materials_from_csv_replaces_existing(tmp_path: Path, app, monkeypatch):
    """Тест, что импорт заменяет существующие данные."""
    config_dir = tmp_path / 'config'
    config_dir.mkdir()
    
    original_config_dir = datasets.CONFIG_DIR
    monkeypatch.setattr(datasets, 'CONFIG_DIR', config_dir)
    
    try:
        # Создаем существующий файл с материалами
        materials_file = config_dir / 'gb_materials.json'
        existing_data = {
            'materials': [
                {
                    'russian': 'Сталь 3',
                    'gb': 'Q235',
                    'notes': 'Старый материал',
                    'gost': '',
                    'price': '',
                    'workpiece_type': ''
                }
            ]
        }
        with materials_file.open('w', encoding='utf-8') as f:
            json.dump(existing_data, f, ensure_ascii=False, indent=2)
        
        # Создаем CSV файл
        csv_path = tmp_path / 'test.csv'
        rows = [
            ['', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', ''],
            ['Родитель', '', '', '', '', '', '', '', 'Итого'],
            ['№ в группе', 'Материал', '', '', 'Наименование мир', '', 'ГОСТ', 'Материал.Цена', 'Материал.Цена.Вид заготовки', 'Количество'],
            ['1', '10', '', '', '10', '', 'ГОСТ 1050-2013', '4.30', 'Лист', '1'],
        ]
        
        with csv_path.open('w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(rows)
        
        # Импортируем
        count = datasets.import_gb_materials_from_csv(csv_path, actor='test_user')
        
        assert count == 1
        
        # Проверяем, что старые данные заменены
        loaded_materials = datasets.load_gb_materials()
        assert len(loaded_materials) == 1
        assert loaded_materials[0]['russian'] == '10'
        assert loaded_materials[0]['gb'] == '10'
        
    finally:
        monkeypatch.setattr(datasets, 'CONFIG_DIR', original_config_dir)


if __name__ == "__main__":
    # Запуск тестов при прямом выполнении файла
    pytest.main([__file__, "-v"])

