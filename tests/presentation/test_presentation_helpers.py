"""Тесты для presentation helpers."""

import os
import sys
import tempfile
from pathlib import Path

# Добавляем корень проекта в путь для возможности прямого запуска
project_root = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(project_root))

import pytest

from app.presentation.helpers import (
    check_templates_exist,
    extract_positions_from_form,
    get_safe_filename
)


class TestCheckTemplatesExist:
    """Тесты для проверки существования шаблонов."""
    
    def test_check_templates_exist_when_missing(self, monkeypatch):
        """Тест проверки когда шаблоны отсутствуют."""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Меняем рабочую директорию на временную
            original_cwd = os.getcwd()
            try:
                os.chdir(tmpdir)
                errors = check_templates_exist()
                assert len(errors) > 0
                assert any('Excel template' in error for error in errors)
                assert any('Word template' in error for error in errors)
            finally:
                os.chdir(original_cwd)
    
    def test_check_templates_exist_when_present(self, monkeypatch):
        """Тест проверки когда шаблоны присутствуют."""
        with tempfile.TemporaryDirectory() as tmpdir:
            original_cwd = os.getcwd()
            try:
                os.chdir(tmpdir)
                # Создаем директорию и файлы
                templates_dir = Path(tmpdir) / 'templates_docs'
                templates_dir.mkdir()
                (templates_dir / 'template.xlsx').touch()
                (templates_dir / 'template.docx').touch()
                
                errors = check_templates_exist()
                assert len(errors) == 0
            finally:
                os.chdir(original_cwd)


class TestGetSafeFilename:
    """Тесты для создания безопасных имен файлов."""
    
    def test_get_safe_filename_basic(self):
        """Тест базового создания безопасного имени."""
        filename = get_safe_filename('ООО Компания')
        assert filename == 'ООО_Компания'
        assert len(filename) <= 50
    
    def test_get_safe_filename_with_special_chars(self):
        """Тест с специальными символами."""
        filename = get_safe_filename('Компания "Тест" & Co.')
        assert '"' not in filename
        assert '&' not in filename
        assert '.' not in filename
    
    def test_get_safe_filename_with_spaces(self):
        """Тест с пробелами."""
        filename = get_safe_filename('Компания с пробелами')
        assert ' ' not in filename
        assert '_' in filename
    
    def test_get_safe_filename_long_name(self):
        """Тест с длинным именем."""
        long_name = 'О' * 100
        filename = get_safe_filename(long_name)
        assert len(filename) <= 50
    
    def test_get_safe_filename_empty(self):
        """Тест с пустым именем."""
        filename = get_safe_filename('')
        assert filename == 'file'  # Пустое имя заменяется на 'file'


class TestExtractPositionsFromForm:
    """Тесты для извлечения позиций из формы."""
    
    def test_extract_single_position(self):
        """Тест извлечения одной позиции."""
        form_data = {
            'product': 'Товар 1',
            'quantity': '10',
            'cost_price': '1000',
            'weight': '5',
            'duty_percent': '5'
        }
        
        positions = extract_positions_from_form(form_data)
        assert len(positions) == 1
        assert positions[0]['product'] == 'Товар 1'
        assert positions[0]['quantity'] == '10'
    
    def test_extract_multiple_positions(self):
        """Тест извлечения множественных позиций."""
        form_data = {
            'product': 'Товар 1',
            'quantity': '10',
            'cost_price': '1000',
            'weight': '5',
            'duty_percent': '5',
            'product_2': 'Товар 2',
            'quantity_2': '20',
            'cost_price_2': '2000',
            'weight_2': '10',
            'duty_percent_2': '10'
        }
        
        positions = extract_positions_from_form(form_data)
        assert len(positions) == 2
        assert positions[0]['product'] == 'Товар 1'
        assert positions[1]['product'] == 'Товар 2'
    
    def test_extract_positions_with_field_keys(self):
        """Тест извлечения позиций с ключами полей."""
        form_data = {
            'product': 'Товар 1',
            'quantity': '10',
            'product_2': 'Товар 2',
            'quantity_2': '20'
        }
        
        positions, field_keys = extract_positions_from_form(form_data, include_field_keys=True)
        assert len(positions) == 2
        assert len(field_keys) == 2
        assert field_keys[0]['product'] == 'product'
        assert field_keys[1]['product'] == 'product_2'
    
    def test_extract_positions_empty_form(self):
        """Тест извлечения из пустой формы."""
        form_data = {}
        positions = extract_positions_from_form(form_data)
        assert len(positions) == 0
    
    def test_extract_positions_partial_data(self):
        """Тест извлечения с частичными данными."""
        form_data = {
            'product': 'Товар 1',
            'quantity': '10'
            # Отсутствуют cost_price, weight, duty_percent
        }
        
        positions = extract_positions_from_form(form_data)
        assert len(positions) == 1
        assert 'product' in positions[0]
        assert 'quantity' in positions[0]
        assert 'cost_price' not in positions[0]
    
    def test_extract_positions_mixed_numbering(self):
        """Тест извлечения с смешанной нумерацией."""
        form_data = {
            'product': 'Товар 1',
            'quantity': '10',
            'product_3': 'Товар 3',
            'quantity_3': '30',
            'product_2': 'Товар 2',
            'quantity_2': '20'
        }
        
        positions = extract_positions_from_form(form_data)
        assert len(positions) == 3
        # Должны быть отсортированы по номеру
        assert positions[0]['product'] == 'Товар 1'
        assert positions[1]['product'] == 'Товар 2'
        assert positions[2]['product'] == 'Товар 3'
    
    def test_extract_positions_with_drawing_number(self):
        """Тест извлечения с номерами чертежей."""
        form_data = {
            'product': 'Товар 1',
            'drawing_number': 'DWG-001',
            'quantity': '10',
            'product_2': 'Товар 2',
            'drawing_number_2': 'DWG-002',
            'quantity_2': '20'
        }
        
        positions = extract_positions_from_form(form_data)
        assert len(positions) == 2
        assert positions[0]['drawing_number'] == 'DWG-001'
        assert positions[1]['drawing_number'] == 'DWG-002'
    
    def test_extract_positions_with_material(self):
        """Тест извлечения с материалами."""
        form_data = {
            'product': 'Товар 1',
            'material': 'Сталь',
            'quantity': '10',
            'product_2': 'Товар 2',
            'material_2': 'Алюминий',
            'quantity_2': '20'
        }
        
        positions = extract_positions_from_form(form_data)
        assert len(positions) == 2
        assert positions[0]['material'] == 'Сталь'
        assert positions[1]['material'] == 'Алюминий'
    
    def test_extract_positions_with_cost_price_per_kg(self):
        """Тест извлечения с ценой за кг."""
        form_data = {
            'product': 'Товар 1',
            'cost_price_per_kg': '200',
            'quantity': '10',
            'weight': '5'
        }
        
        positions = extract_positions_from_form(form_data)
        assert len(positions) == 1
        assert positions[0]['cost_price_per_kg'] == '200'

    def test_extract_positions_duty_percent_zero(self):
        """Пошлина 0% должна извлекаться и попадать в позиции (для записи в Excel)."""
        form_data = {
            'product': 'Товар без пошлины',
            'quantity': '5',
            'cost_price': '500',
            'weight': '2',
            'duty_percent': '0',
        }
        positions = extract_positions_from_form(form_data)
        assert len(positions) == 1
        assert positions[0].get('duty_percent') == '0'

    def test_extract_positions_multiple_with_duty_zero_second(self):
        """Вторая позиция с duty_percent=0 должна извлекаться."""
        form_data = {
            'product': 'Товар 1',
            'quantity': '10',
            'cost_price': '1000',
            'weight': '5',
            'duty_percent': '5',
            'product_2': 'Товар 2',
            'quantity_2': '20',
            'cost_price_2': '2000',
            'weight_2': '10',
            'duty_percent_2': '0',
        }
        positions = extract_positions_from_form(form_data)
        assert len(positions) == 2
        assert positions[0]['duty_percent'] == '5'
        assert positions[1]['duty_percent'] == '0'

    def test_extract_positions_with_payload(self):
        """Тест извлечения позиций из positions_payload."""
        positions_data = [
            {'product': 'Товар 1', 'quantity': '10', 'cost_price': '1000'},
            {'product': 'Товар 2', 'quantity': '20', 'cost_price': '2000'}
        ]
        import json
        form_data = {
            'positions_payload': json.dumps(positions_data, ensure_ascii=False)
        }
        
        positions = extract_positions_from_form(form_data)
        assert len(positions) == 2
        assert positions[0]['product'] == 'Товар 1'
        assert positions[1]['product'] == 'Товар 2'

    def test_extract_positions_with_path_traversal_protection(self):
        """Тест защиты от path traversal в get_safe_filename."""
        dangerous_names = [
            '../../../etc/passwd',
            '..\\..\\windows\\system32',
            '/etc/shadow',
            'C:\\Windows\\System32',
            '../../templates',
        ]
        
        for dangerous_name in dangerous_names:
            safe_name = get_safe_filename(dangerous_name)
            assert '..' not in safe_name
            assert '/' not in safe_name or safe_name == Path(safe_name).name
            assert '\\' not in safe_name or safe_name == Path(safe_name).name
            assert len(safe_name) <= 50


if __name__ == "__main__":
    # Запуск тестов при прямом выполнении файла
    pytest.main([__file__, "-v"])

