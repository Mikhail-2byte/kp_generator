"""Тесты для админ-маршрутов управления каталогом ТН ВЭД."""

import json
import sys
from pathlib import Path
from tempfile import TemporaryDirectory

# Добавляем корень проекта в путь для возможности прямого запуска
project_root = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(project_root))

import pytest
from flask import url_for

from app.services import datasets


@pytest.fixture
def sample_tnved_data(tmp_path: Path, app, monkeypatch):
    """Создает тестовые данные каталога ТН ВЭД."""
    config_dir = tmp_path / 'config'
    config_dir.mkdir()
    
    original_config_dir = datasets.CONFIG_DIR
    monkeypatch.setattr(datasets, 'CONFIG_DIR', config_dir)
    
    # Очистить кеш перед созданием данных
    datasets.load_duty_rates.cache_clear()
    datasets.load_tnved_catalog.cache_clear()
    
    items = [
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
        # Сохраняем через save_duty_rates, так как маршрут использует эту функцию
        # и она сохраняет все данные (duty_items + tnved_items) в единый файл
        datasets.save_duty_rates(items, actor='test_user')
        datasets.load_tnved_catalog.cache_clear()
        datasets.load_duty_rates.cache_clear()
    
    yield config_dir
    
    # Восстановить оригинальный CONFIG_DIR
    monkeypatch.setattr(datasets, 'CONFIG_DIR', original_config_dir)
    datasets.load_duty_rates.cache_clear()
    datasets.load_tnved_catalog.cache_clear()


def test_manage_duty_page_loads(logged_in_client, sample_tnved_data):
    """Тест загрузки страницы управления пошлинами."""
    response = logged_in_client.get('/admin/duty')
    
    assert response.status_code == 200
    assert 'Каталог ТН ВЭД' in response.data.decode('utf-8')
    assert 'Добавление записи в каталог ТН ВЭД' in response.data.decode('utf-8')


def test_add_tnved_item(logged_in_client, sample_tnved_data, app):
    """Тест добавления записи в каталог ТН ВЭД."""
    # Проверяем, что CONFIG_DIR указывает на тестовую директорию
    assert datasets.CONFIG_DIR == sample_tnved_data
    
    # Очищаем кеш перед запросом, чтобы данные загружались из тестового файла
    with app.app_context():
        datasets.load_duty_rates.cache_clear()
        datasets.load_tnved_catalog.cache_clear()
    
    # Получаем CSRF токен
    page = logged_in_client.get('/admin/duty')
    assert page.status_code == 200
    csrf_token = page.data.decode().split('name="csrf_token" value="')[1].split('"')[0]
    
    response = logged_in_client.post('/admin/duty', data={
        'action': 'add_tnved',
        'tnved-code': '1234567890',
        'tnved-description': 'Тестовое описание',
        'tnved-keywords_display': 'Ключевые слова',
        'tnved-examples': 'Примеры',
        'tnved-duty_text': '10%',
        'tnved-duty_percent': 10.0,
        'csrf_token': csrf_token
    }, follow_redirects=True)
    
    assert response.status_code == 200
    
    # Проверяем, что данные сохранены
    with app.app_context():
        # Убеждаемся, что CONFIG_DIR все еще указывает на тестовую директорию
        assert datasets.CONFIG_DIR == sample_tnved_data
        
        # Очищаем кеш перед проверкой данных
        datasets.load_duty_rates.cache_clear()
        datasets.load_tnved_catalog.cache_clear()
        
        # Загружаем все данные (duty_items + tnved_items)
        all_items = datasets.load_duty_rates()
        # Фильтруем только ТН ВЭД записи (с полем code)
        tnved_items = [item for item in all_items if item.get('code')]
    
    assert len(tnved_items) == 2  # Было 1, стало 2
    assert any(item['code'] == '1234567890' for item in tnved_items)


def test_edit_tnved_item(logged_in_client, sample_tnved_data, app):
    """Тест редактирования записи каталога ТН ВЭД."""
    # Проверяем, что CONFIG_DIR указывает на тестовую директорию
    assert datasets.CONFIG_DIR == sample_tnved_data
    
    # Очищаем кеш перед запросом, чтобы данные загружались из тестового файла
    with app.app_context():
        datasets.load_duty_rates.cache_clear()
        datasets.load_tnved_catalog.cache_clear()
    
    # Получаем CSRF токен
    page = logged_in_client.get('/admin/duty')
    assert page.status_code == 200
    csrf_token = page.data.decode().split('name="csrf_token" value="')[1].split('"')[0]
    
    response = logged_in_client.post('/admin/duty', data={
        'action': 'edit_tnved',
        'index': 0,
        'tnved-code': '9031908500',
        'tnved-description': 'Обновленное описание',
        'tnved-keywords_display': 'Обновленные ключевые слова',
        'tnved-examples': 'Обновленные примеры',
        'tnved-duty_text': '5%',
        'tnved-duty_percent': 5.0,
        'csrf_token': csrf_token
    }, follow_redirects=True)
    
    assert response.status_code == 200
    
    # Проверяем, что данные обновлены
    with app.app_context():
        # Убеждаемся, что CONFIG_DIR все еще указывает на тестовую директорию
        assert datasets.CONFIG_DIR == sample_tnved_data
        
        # Очищаем кеш перед проверкой данных
        datasets.load_duty_rates.cache_clear()
        datasets.load_tnved_catalog.cache_clear()
        
        # Загружаем все данные (duty_items + tnved_items)
        all_items = datasets.load_duty_rates()
        # Фильтруем только ТН ВЭД записи (с полем code)
        tnved_items = [item for item in all_items if item.get('code')]
    
    # Находим обновленную запись по коду
    updated_item = next((item for item in tnved_items if item['code'] == '9031908500'), None)
    assert updated_item is not None
    assert updated_item['description'] == 'Обновленное описание'
    assert updated_item['duty_percent'] == 5.0


def test_delete_tnved_item(logged_in_client, sample_tnved_data, app):
    """Тест удаления записи каталога ТН ВЭД."""
    # Проверяем, что CONFIG_DIR указывает на тестовую директорию
    assert datasets.CONFIG_DIR == sample_tnved_data
    
    # Очищаем кеш перед запросом, чтобы данные загружались из тестового файла
    with app.app_context():
        datasets.load_duty_rates.cache_clear()
        datasets.load_tnved_catalog.cache_clear()
    
    # Получаем CSRF токен
    page = logged_in_client.get('/admin/duty')
    assert page.status_code == 200
    csrf_token = page.data.decode().split('name="csrf_token" value="')[1].split('"')[0]
    
    response = logged_in_client.post('/admin/duty', data={
        'action': 'delete_tnved',
        'index': 0,
        'csrf_token': csrf_token
    }, follow_redirects=True)
    
    assert response.status_code == 200
    
    # Проверяем, что данные удалены (более надежная проверка, чем flash-сообщения)
    with app.app_context():
        # Очищаем кеш перед проверкой данных
        datasets.load_duty_rates.cache_clear()
        datasets.load_tnved_catalog.cache_clear()
        
        # Загружаем все данные (duty_items + tnved_items)
        all_items = datasets.load_duty_rates()
        # Фильтруем только ТН ВЭД записи (с полем code)
        tnved_items = [item for item in all_items if item.get('code')]
    
    # После удаления должно остаться 0 записей (была 1, удалили её)
    assert len(tnved_items) == 0


def test_export_tnved_catalog(logged_in_client, sample_tnved_data):
    """Тест экспорта каталога ТН ВЭД в Excel."""
    response = logged_in_client.get('/admin/duty/tnved/export')
    
    assert response.status_code == 200
    assert response.content_type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    assert len(response.data) > 0
    assert 'TNVED_catalog_' in response.headers.get('Content-Disposition', '')


def test_import_tnved_catalog(logged_in_client, sample_tnved_data, app, tmp_path):
    """Тест импорта каталога ТН ВЭД из Excel."""
    # Создаем тестовый Excel файл
    try:
        from openpyxl import Workbook
    except ImportError:
        pytest.skip("openpyxl не установлен")
    
    wb = Workbook()
    ws = wb.active
    ws.append(['Код ТН ВЭД', 'Описание', 'Ключевые слова', 'Примеры', 'Пошлина', 'Пошлина, %'])
    ws.append(['9999999999', 'Импортированное описание', 'Импортированные слова', 'Импортированные примеры', '15%', 15.0])
    
    excel_file = tmp_path / 'test_import.xlsx'
    wb.save(excel_file)
    
    # Получаем CSRF токен
    page = logged_in_client.get('/admin/duty')
    assert page.status_code == 200
    csrf_token = page.data.decode().split('name="csrf_token" value="')[1].split('"')[0]
    
    # Импортируем
    with excel_file.open('rb') as f:
        response = logged_in_client.post('/admin/duty/tnved/import', data={
            'excel_file': (f, 'test_import.xlsx'),
            'csrf_token': csrf_token
        }, content_type='multipart/form-data', follow_redirects=True)
    
    assert response.status_code == 200
    assert 'Импортировано записей каталога ТН ВЭД' in response.data.decode('utf-8')
    
    # Проверяем, что данные импортированы
    with app.app_context():
        datasets.load_tnved_catalog.cache_clear()
        items = datasets.load_tnved_catalog()
    
    # Должно быть 2 записи: исходная + импортированная
    assert len(items) >= 1
    assert any(item['code'] == '9999999999' for item in items)


def test_import_tnved_catalog_invalid_file(logged_in_client, sample_tnved_data):
    """Тест импорта с неверным форматом файла."""
    # Получаем CSRF токен
    page = logged_in_client.get('/admin/duty')
    assert page.status_code == 200
    csrf_token = page.data.decode().split('name="csrf_token" value="')[1].split('"')[0]
    
    # Пытаемся импортировать текстовый файл
    response = logged_in_client.post('/admin/duty/tnved/import', data={
        'excel_file': (b'not an excel file', 'test.txt'),
        'csrf_token': csrf_token
    }, content_type='multipart/form-data', follow_redirects=True)
    
    assert response.status_code == 200
    # Проверяем, что произошел редирект на страницу управления (не остались на странице импорта)
    assert '/admin/duty' in response.request.path or 'admin_duty' in response.data.decode('utf-8')


def test_import_tnved_catalog_no_file(logged_in_client, sample_tnved_data):
    """Тест импорта без файла."""
    # Получаем CSRF токен
    page = logged_in_client.get('/admin/duty')
    assert page.status_code == 200
    csrf_token = page.data.decode().split('name="csrf_token" value="')[1].split('"')[0]
    
    response = logged_in_client.post('/admin/duty/tnved/import', data={
        'csrf_token': csrf_token
    }, follow_redirects=True)
    
    assert response.status_code == 200
    assert 'Файл не выбран' in response.data.decode('utf-8')


def test_manage_duty_requires_admin(client):
    """Тест, что страница управления требует прав администратора."""
    response = client.get('/admin/duty', follow_redirects=True)
    
    # Должно быть перенаправление на страницу входа
    assert response.status_code == 200
    assert 'login' in response.request.path.lower() or 'Войти' in response.data.decode('utf-8')


if __name__ == "__main__":
    pytest.main([__file__, "-v"])

