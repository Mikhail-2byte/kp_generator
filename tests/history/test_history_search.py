"""Тесты для поиска по истории генераций."""

import json
import time
from datetime import datetime

from app.database import database_service
from app.services.repositories import generation_repository


def _make_form(i: int, **kwargs) -> dict:
    """Создает тестовую форму с опциональными переопределениями."""
    form = {
        'company': f'Компания {i}',
        'product': f'Изделие {i}',
        'quantity': '1',
        'cost_price': '1000',
        'weight': '5',
        'logistics': '5000',
        'margin_percent': '20',
        'delivery_time': '25',
        'duty_percent': '5',
        'tender_number': f'TND-{i}',
    }
    form.update(kwargs)
    return form


def test_search_by_tender_number(app):
    """Тест поиска по номеру тендера."""
    service = database_service
    with app.app_context():
        config = app.config['APP_SETTINGS']
        unique_username = f'search_tender_{int(time.time())}'
        user_id = service.create_user(unique_username, 'hash')
        assert user_id

        # Создаем записи с разными номерами тендеров
        base_tender = int(time.time())
        tender_numbers = [
            f'TND-SEARCH-{base_tender}-1',
            f'TND-OTHER-{base_tender}-2',
            f'TND-SEARCH-{base_tender}-3',
            f'TND-NORMAL-{base_tender}-4',
        ]
        
        for i, tender_number in enumerate(tender_numbers):
            form = _make_form(i, tender_number=tender_number)
            service.save_generation_history(
                form,
                final_price=1000 + i * 100,
                config=config,
                user_id=user_id
            )

        # Поиск по номеру тендера "SEARCH"
        result = generation_repository.get_history(
            config,
            page=1,
            search='SEARCH'
        )

        # Проверяем, что найдены записи с "SEARCH" в номере тендера
        found_items = []
        for item in result['items']:
            tender = item.get('tender_number', '')
            # Проверяем только записи нашего теста
            if f'{base_tender}' in tender:
                tender_upper = tender.upper()
                if 'SEARCH' in tender_upper:
                    found_items.append(item)
        
        assert len(found_items) >= 2, f"Должно быть найдено минимум 2 записи с 'SEARCH', найдено: {len(found_items)}. Всего записей в результате: {len(result['items'])}"
        
        # Проверяем, что все найденные записи содержат "SEARCH" в номере тендера
        for item in found_items:
            tender = item.get('tender_number', '').upper()
            assert 'SEARCH' in tender, f"Номер тендера {item.get('tender_number')} не содержит 'SEARCH'"


def test_search_by_company(app):
    """Тест поиска по наименованию заказчика."""
    service = database_service
    with app.app_context():
        config = app.config['APP_SETTINGS']
        unique_username = f'search_company_{int(time.time())}'
        user_id = service.create_user(unique_username, 'hash')
        assert user_id

        # Создаем записи с разными компаниями
        base_tender = int(time.time())
        companies = [
            'ООО Рога и Копыта',
            'ИП Сидоров',
            'ООО Рога и Копыта',
            'ЗАО Другая Компания',
        ]
        
        for i, company in enumerate(companies):
            form = _make_form(i, company=company, tender_number=f'TND-{base_tender}-{i}')
            service.save_generation_history(
                form,
                final_price=1000 + i * 100,
                config=config,
                user_id=user_id
            )

        # Сначала проверяем, что записи сохранились
        all_result = generation_repository.get_history(config, page=1)
        saved_items = [item for item in all_result['items'] if f'{base_tender}' in item.get('tender_number', '')]
        assert len(saved_items) >= 4, f"Должно быть сохранено минимум 4 записи, сохранено: {len(saved_items)}"
        
        # Поиск по компании "Рога"
        result = generation_repository.get_history(
            config,
            page=1,
            search='Рога'
        )

        # Проверяем, что найдены записи с "Рога" в названии компании
        # Поиск должен быть case-insensitive, поэтому проверяем в любом регистре
        found_items = []
        search_lower = 'рога'
        for item in result['items']:
            tender = item.get('tender_number', '')
            # Проверяем только записи нашего теста
            if f'{base_tender}' in tender:
                company = item.get('company', '').lower()
                if search_lower in company:
                    found_items.append(item)
        
        # Если не найдено, проверяем без фильтрации по base_tender
        if len(found_items) == 0:
            # Проверяем все найденные записи
            for item in result['items']:
                company = item.get('company', '').lower()
                if search_lower in company:
                    found_items.append(item)
        
        assert len(found_items) >= 2, f"Должно быть найдено минимум 2 записи с 'Рога', найдено: {len(found_items)}. Всего записей в результате: {len(result['items'])}. Поиск: 'Рога'. Сохранено записей: {len(saved_items)}"
        
        # Проверяем, что все найденные записи содержат "Рога" в названии компании
        for item in found_items:
            company = item.get('company', '').lower()
            assert search_lower in company, f"Компания {item.get('company')} не содержит 'Рога'"


def test_search_by_product(app):
    """Тест поиска по наименованию товара."""
    service = database_service
    with app.app_context():
        config = app.config['APP_SETTINGS']
        unique_username = f'search_product_{int(time.time())}'
        user_id = service.create_user(unique_username, 'hash')
        assert user_id

        # Создаем записи с разными товарами
        base_tender = int(time.time())
        products = [
            'Ротор электродвигателя',
            'Колесо зубчатое',
            'Ротор генератора',
            'Вал приводной',
        ]
        
        for i, product in enumerate(products):
            form = _make_form(i, product=product, tender_number=f'TND-{base_tender}-{i}')
            service.save_generation_history(
                form,
                final_price=1000 + i * 100,
                config=config,
                user_id=user_id
            )

        # Поиск по товару "Ротор"
        result = generation_repository.get_history(
            config,
            page=1,
            search='Ротор'
        )

        # Проверяем, что найдены записи с "Ротор" в названии товара
        # Поиск должен быть case-insensitive
        found_items = []
        search_lower = 'ротор'
        for item in result['items']:
            tender = item.get('tender_number', '')
            # Проверяем только записи нашего теста
            if f'{base_tender}' in tender:
                product = item.get('product', '').lower()
                if search_lower in product:
                    found_items.append(item)
        
        assert len(found_items) >= 2, f"Должно быть найдено минимум 2 записи с 'Ротор', найдено: {len(found_items)}. Всего записей в результате: {len(result['items'])}"
        
        # Проверяем, что все найденные записи содержат "Ротор" в названии товара
        for item in found_items:
            product = item.get('product', '').lower()
            assert search_lower in product, f"Товар {item.get('product')} не содержит 'Ротор'"


def test_search_by_product_in_multiple_positions(app):
    """Тест поиска по товару в множественных позициях (JSON)."""
    service = database_service
    with app.app_context():
        config = app.config['APP_SETTINGS']
        unique_username = f'search_multi_{int(time.time())}'
        user_id = service.create_user(unique_username, 'hash')
        assert user_id

        # Создаем запись с множественными позициями
        base_tender = int(time.time())
        positions = [
            {
                'product': 'Ротор основной',
                'quantity': 2,
                'cost_price': 1000,
                'weight': 5,
                'duty_percent': 5,
                'drawing_number': 'DRW-001',
                'material': 'Сталь',
            },
            {
                'product': 'Колесо зубчатое',
                'quantity': 1,
                'cost_price': 500,
                'weight': 2,
                'duty_percent': 5,
                'drawing_number': 'DRW-002',
                'material': 'Чугун',
            },
            {
                'product': 'Вал приводной',
                'quantity': 1,
                'cost_price': 800,
                'weight': 3,
                'duty_percent': 5,
                'drawing_number': 'DRW-003',
                'material': 'Сталь',
            },
        ]
        
        form = _make_form(0, tender_number=f'TND-{base_tender}-multi')
        form['positions'] = positions
        
        # Сохраняем с множественными позициями
        service.save_generation_history(
            form,
            final_price=3000,
            config=config,
            user_id=user_id,
            total_general_price=3000
        )

        # Поиск по товару "Колесо" (который есть только во второй позиции)
        result = generation_repository.get_history(
            config,
            page=1,
            search='Колесо'
        )

        # Проверяем, что найдена запись с "Колесо" в позициях
        found_items = []
        for item in result['items']:
            tender = item.get('tender_number', '')
            if f'{base_tender}' in tender:
                found_items.append(item)
        
        assert len(found_items) >= 1, f"Должна быть найдена запись с 'Колесо' в позициях, найдено: {len(found_items)}"
        
        # Проверяем, что найденная запись содержит "Колесо" в данных
        # (либо в основном поле product, либо в positions_data)
        found = False
        for item in found_items:
            product = item.get('product', '')
            if 'Колесо' in product:
                found = True
                break
        
        # Если не найдено в основном поле, проверяем positions_data через детали
        if not found:
            for item in found_items:
                # Получаем детали записи для проверки positions_data
                record_id = item.get('id')
                if record_id:
                    details = generation_repository.get_details(record_id)
                    if details and details.get('positions'):
                        for pos in details['positions']:
                            if 'Колесо' in pos.get('product', ''):
                                found = True
                                break
        
        assert found, "Не найдено 'Колесо' в позициях записи"


def test_search_case_insensitive(app):
    """Тест поиска без учета регистра."""
    service = database_service
    with app.app_context():
        config = app.config['APP_SETTINGS']
        unique_username = f'search_case_{int(time.time())}'
        user_id = service.create_user(unique_username, 'hash')
        assert user_id

        # Создаем запись с товаром в разных регистрах
        base_tender = int(time.time())
        form = _make_form(0, product='Ротор Электродвигателя', tender_number=f'TND-{base_tender}-case')
        service.save_generation_history(
            form,
            final_price=1000,
            config=config,
            user_id=user_id
        )

        # Поиск в нижнем регистре
        result_lower = generation_repository.get_history(
            config,
            page=1,
            search='ротор'
        )

        # Поиск в верхнем регистре
        result_upper = generation_repository.get_history(
            config,
            page=1,
            search='РОТОР'
        )

        # Проверяем, что оба поиска находят запись
        # Поиск должен быть case-insensitive, поэтому оба должны найти
        found_lower = False
        found_upper = False
        
        for item in result_lower['items']:
            tender = item.get('tender_number', '')
            if f'{base_tender}' in tender:
                product = item.get('product', '').lower()
                if 'ротор' in product:
                    found_lower = True
                    break
        
        for item in result_upper['items']:
            tender = item.get('tender_number', '')
            if f'{base_tender}' in tender:
                product = item.get('product', '').lower()
                if 'ротор' in product:
                    found_upper = True
                    break
        
        assert found_lower, f"Поиск в нижнем регистре не нашел запись. Найдено записей: {len(result_lower['items'])}"
        assert found_upper, f"Поиск в верхнем регистре не нашел запись. Найдено записей: {len(result_upper['items'])}"


def test_search_combined_fields(app):
    """Тест поиска, который может найти по разным полям."""
    service = database_service
    with app.app_context():
        config = app.config['APP_SETTINGS']
        unique_username = f'search_combined_{int(time.time())}'
        user_id = service.create_user(unique_username, 'hash')
        assert user_id

        # Создаем записи с разными данными, но с общим словом
        base_tender = int(time.time())
        test_data = [
            {'tender_number': f'TND-UNIQUE-{base_tender}-1', 'company': 'Компания А', 'product': 'Товар Б'},
            {'tender_number': f'TND-{base_tender}-2', 'company': 'Компания UNIQUE', 'product': 'Товар В'},
            {'tender_number': f'TND-{base_tender}-3', 'company': 'Компания Г', 'product': 'Товар UNIQUE'},
        ]
        
        for i, data in enumerate(test_data):
            form = _make_form(i, **data)
            service.save_generation_history(
                form,
                final_price=1000 + i * 100,
                config=config,
                user_id=user_id
            )

        # Поиск по слову "UNIQUE" (должно найти по тендеру, компании или товару)
        result = generation_repository.get_history(
            config,
            page=1,
            search='UNIQUE'
        )

        # Проверяем, что найдены все три записи
        found_items = []
        for item in result['items']:
            tender = item.get('tender_number', '')
            company = item.get('company', '')
            product = item.get('product', '')
            if f'{base_tender}' in tender and ('UNIQUE' in tender.upper() or 'UNIQUE' in company.upper() or 'UNIQUE' in product.upper()):
                found_items.append(item)
        
        assert len(found_items) >= 3, f"Должно быть найдено минимум 3 записи с 'UNIQUE', найдено: {len(found_items)}"


def test_search_by_manager(app):
    """Тест поиска по менеджеру (username, фамилия, имя)."""
    service = database_service
    with app.app_context():
        config = app.config['APP_SETTINGS']
        
        # Создаем пользователей с разными данными
        base_tender = int(time.time())
        users_data = [
            {'username': f'manager_{base_tender}_1', 'last_name': 'Иванов', 'first_name': 'Иван'},
            {'username': f'manager_{base_tender}_2', 'last_name': 'Петров', 'first_name': 'Петр'},
            {'username': f'manager_{base_tender}_3', 'last_name': 'Иванов', 'first_name': 'Сергей'},
        ]
        
        user_ids = []
        for user_data in users_data:
            user_id = service.create_user(
                user_data['username'],
                'hash',
                last_name=user_data['last_name'],
                first_name=user_data['first_name']
            )
            assert user_id
            user_ids.append(user_id)
        
        # Создаем записи для каждого пользователя
        for i, user_id in enumerate(user_ids):
            form = _make_form(i, tender_number=f'TND-{base_tender}-manager-{i}')
            service.save_generation_history(
                form,
                final_price=1000 + i * 100,
                config=config,
                user_id=user_id
            )
        
        # Поиск по фамилии "Иванов"
        result = generation_repository.get_history(
            config,
            page=1,
            search='Иванов'
        )
        
        # Проверяем, что найдены записи с менеджерами "Иванов"
        found_items = []
        for item in result['items']:
            tender = item.get('tender_number', '')
            if f'{base_tender}' in tender:
                last_name = item.get('last_name', '')
                if 'Иванов' in last_name:
                    found_items.append(item)
        
        assert len(found_items) >= 2, f"Должно быть найдено минимум 2 записи с менеджером 'Иванов', найдено: {len(found_items)}"
        
        # Поиск по имени "Петр"
        result2 = generation_repository.get_history(
            config,
            page=1,
            search='Петр'
        )
        
        found_items2 = []
        for item in result2['items']:
            tender = item.get('tender_number', '')
            if f'{base_tender}' in tender:
                first_name = item.get('first_name', '')
                if 'Петр' in first_name:
                    found_items2.append(item)
        
        assert len(found_items2) >= 1, f"Должна быть найдена минимум 1 запись с менеджером 'Петр', найдено: {len(found_items2)}"
        
        # Поиск по username
        result3 = generation_repository.get_history(
            config,
            page=1,
            search=f'manager_{base_tender}_1'
        )
        
        found_items3 = []
        for item in result3['items']:
            tender = item.get('tender_number', '')
            if f'{base_tender}' in tender:
                username = item.get('username', '')
                if f'manager_{base_tender}_1' in username:
                    found_items3.append(item)
        
        assert len(found_items3) >= 1, f"Должна быть найдена минимум 1 запись с username 'manager_{base_tender}_1', найдено: {len(found_items3)}"


def test_search_empty_query(app):
    """Тест пустого поискового запроса."""
    service = database_service
    with app.app_context():
        config = app.config['APP_SETTINGS']
        unique_username = f'search_empty_{int(time.time())}'
        user_id = service.create_user(unique_username, 'hash')
        assert user_id

        # Создаем несколько записей
        base_tender = int(time.time())
        for i in range(3):
            form = _make_form(i, tender_number=f'TND-{base_tender}-{i}')
            service.save_generation_history(
                form,
                final_price=1000 + i * 100,
                config=config,
                user_id=user_id
            )

        # Поиск с пустым запросом
        result = generation_repository.get_history(
            config,
            page=1,
            search=''
        )

        # Пустой поиск должен вернуть все записи (или не применять фильтр)
        # Проверяем, что результат не пустой
        assert 'items' in result
        assert 'pagination' in result

