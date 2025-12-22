"""Тесты для серверной фильтрации истории генераций."""

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


def test_filter_by_price_range(app):
    """Тест фильтрации по диапазону цен."""
    service = database_service
    with app.app_context():
        config = app.config['APP_SETTINGS']
        unique_username = f'price_filter_{int(time.time())}'
        user_id = service.create_user(unique_username, 'hash')
        assert user_id

        # Создаем записи с разными ценами
        base_tender = int(time.time())
        prices = [1000, 5000, 10000, 20000, 50000]
        for i, price in enumerate(prices):
            form = _make_form(i, tender_number=f'TND-{base_tender}-{i}')
            service.save_generation_history(
                form,
                final_price=price,
                config=config,
                user_id=user_id,
                total_general_price=price * 2  # Общая цена = цена * количество
            )

        # Фильтр: цена от 5000 до 20000
        result = generation_repository.get_history(
            config,
            page=1,
            price_from=5000,
            price_to=20000
        )

        # Проверяем, что все записи в диапазоне
        for item in result['items']:
            price = item.get('total_general_price', 0)
            assert 5000 <= price <= 20000, f"Цена {price} не в диапазоне [5000, 20000]"


def test_filter_by_margin_range(app):
    """Тест фильтрации по диапазону маржи."""
    service = database_service
    with app.app_context():
        config = app.config['APP_SETTINGS']
        unique_username = f'margin_filter_{int(time.time())}'
        user_id = service.create_user(unique_username, 'hash')
        assert user_id

        # Создаем записи с разной маржой
        base_tender = int(time.time())
        margins = [10, 15, 20, 25, 30]
        for i, margin in enumerate(margins):
            form = _make_form(i, margin_percent=str(margin), tender_number=f'TND-{base_tender}-{i}')
            service.save_generation_history(
                form,
                final_price=1000 + i * 100,
                config=config,
                user_id=user_id
            )

        # Фильтр: маржа от 15 до 25
        result = generation_repository.get_history(
            config,
            page=1,
            margin_from=15,
            margin_to=25
        )

        # Проверяем, что все записи в диапазоне маржи
        for item in result['items']:
            margin = item.get('margin_percent', 0)
            # Проверяем только записи нашего пользователя
            if item.get('tender_number', '').startswith(f'TND-{base_tender}'):
                assert 15 <= margin <= 25, f"Маржа {margin} не в диапазоне [15, 25]"


def test_filter_by_companies(app):
    """Тест фильтрации по компаниям."""
    service = database_service
    with app.app_context():
        config = app.config['APP_SETTINGS']
        unique_username = f'company_filter_{int(time.time())}'
        user_id = service.create_user(unique_username, 'hash')
        assert user_id

        # Создаем записи для разных компаний
        base_tender = int(time.time())
        companies = ['Компания А', 'Компания Б', 'Компания В', 'Компания Г']
        for i, company in enumerate(companies):
            form = _make_form(i, company=company, tender_number=f'TND-{base_tender}-{i}')
            service.save_generation_history(
                form,
                final_price=1000 + i * 100,
                config=config,
                user_id=user_id
            )

        # Фильтр: только Компания А и Компания В
        result = generation_repository.get_history(
            config,
            page=1,
            companies=['Компания А', 'Компания В']
        )

        # Проверяем, что все записи принадлежат выбранным компаниям
        for item in result['items']:
            company = item.get('company', '')
            if item.get('tender_number', '').startswith(f'TND-{base_tender}'):
                assert company in ['Компания А', 'Компания В'], f"Компания {company} не в списке фильтра"


def test_filter_by_search(app):
    """Тест текстового поиска."""
    service = database_service
    with app.app_context():
        config = app.config['APP_SETTINGS']
        unique_username = f'search_filter_{int(time.time())}'
        user_id = service.create_user(unique_username, 'hash')
        assert user_id

        # Создаем записи с разными номерами тендеров и товарами
        # Используем английские слова для надежности работы с SQLite
        base_tender = int(time.time())
        test_data = [
            (f'TND-SEARCH-{base_tender}-1', 'Search Product One'),
            (f'TND-NORMAL-{base_tender}-2', 'Normal Product Two'),
            (f'TND-SEARCH-{base_tender}-3', 'Search Product Three'),
        ]
        
        for i, (tender, product) in enumerate(test_data):
            form = _make_form(i, tender_number=tender, product=product)
            service.save_generation_history(
                form,
                final_price=1000 + i * 100,
                config=config,
                user_id=user_id
            )

        # Поиск по слову "search" (английское слово для надежности)
        result = generation_repository.get_history(
            config,
            page=1,
            search='search'
        )

        # Проверяем, что все найденные записи содержат "search"
        found_search = False
        for item in result['items']:
            tender = item.get('tender_number', '')
            product = item.get('product', '')
            # Проверяем только записи нашего теста
            if f'{base_tender}' in tender:
                tender_lower = tender.lower()
                product_lower = product.lower()
                if 'search' in tender_lower or 'search' in product_lower:
                    found_search = True
        
        assert found_search, f"Не найдено записей с 'search'. Найдено записей: {len(result['items'])}, tender: {base_tender}"


def test_filter_by_date_range(app):
    """Тест фильтрации по диапазону дат."""
    service = database_service
    with app.app_context():
        config = app.config['APP_SETTINGS']
        unique_username = f'date_filter_{int(time.time())}'
        user_id = service.create_user(unique_username, 'hash')
        assert user_id

        # Создаем записи (они будут с текущей датой)
        base_tender = int(time.time())
        for i in range(5):
            form = _make_form(i, tender_number=f'TND-{base_tender}-{i}')
            service.save_generation_history(
                form,
                final_price=1000 + i * 100,
                config=config,
                user_id=user_id
            )

        # Фильтр: сегодняшняя дата
        today = datetime.now().strftime('%Y-%m-%d')
        result = generation_repository.get_history(
            config,
            page=1,
            date_from=today,
            date_to=today
        )

        # Проверяем, что записи есть (они должны быть с сегодняшней датой)
        assert len(result['items']) > 0, "Не найдено записей за сегодня"


def test_filter_combinations(app):
    """Тест комбинации нескольких фильтров."""
    service = database_service
    with app.app_context():
        config = app.config['APP_SETTINGS']
        unique_username = f'combo_filter_{int(time.time())}'
        user_id = service.create_user(unique_username, 'hash')
        assert user_id

        # Создаем записи с разными параметрами
        base_tender = int(time.time())
        test_cases = [
            {'company': 'Компания X', 'margin_percent': '20', 'price': 10000},
            {'company': 'Компания X', 'margin_percent': '25', 'price': 15000},
            {'company': 'Компания Y', 'margin_percent': '20', 'price': 10000},
            {'company': 'Компания X', 'margin_percent': '15', 'price': 5000},
        ]
        
        for i, case in enumerate(test_cases):
            form = _make_form(
                i,
                company=case['company'],
                margin_percent=case['margin_percent'],
                tender_number=f'TND-{base_tender}-{i}'
            )
            service.save_generation_history(
                form,
                final_price=case['price'],
                config=config,
                user_id=user_id,
                total_general_price=case['price']
            )

        # Комбинированный фильтр: Компания X, маржа >= 20, цена >= 10000
        result = generation_repository.get_history(
            config,
            page=1,
            companies=['Компания X'],
            margin_from=20,
            price_from=10000
        )

        # Проверяем, что все записи соответствуют всем критериям
        for item in result['items']:
            if item.get('tender_number', '').startswith(f'TND-{base_tender}'):
                company = item.get('company', '')
                margin = item.get('margin_percent', 0)
                price = item.get('total_general_price', 0)
                
                assert company == 'Компания X', f"Компания {company} не соответствует фильтру"
                assert margin >= 20, f"Маржа {margin} меньше 20"
                assert price >= 10000, f"Цена {price} меньше 10000"


def test_sorting_by_price(app):
    """Тест сортировки по цене."""
    service = database_service
    with app.app_context():
        config = app.config['APP_SETTINGS']
        unique_username = f'sort_price_{int(time.time())}'
        user_id = service.create_user(unique_username, 'hash')
        assert user_id

        # Создаем записи с разными ценами
        base_tender = int(time.time())
        prices = [5000, 2000, 10000, 1000, 8000]
        for i, price in enumerate(prices):
            form = _make_form(i, tender_number=f'TND-{base_tender}-{i}')
            service.save_generation_history(
                form,
                final_price=price,
                config=config,
                user_id=user_id,
                total_general_price=price
            )

        # Сортировка по цене по убыванию
        result = generation_repository.get_history(
            config,
            page=1,
            sort_by='price_sale',
            sort_order='desc'
        )

        # Проверяем, что записи отсортированы по убыванию цены
        prices_found = []
        for item in result['items']:
            if item.get('tender_number', '').startswith(f'TND-{base_tender}'):
                prices_found.append(item.get('total_general_price', 0))
        
        if len(prices_found) > 1:
            assert prices_found == sorted(prices_found, reverse=True), "Цены не отсортированы по убыванию"


def test_sorting_by_margin(app):
    """Тест сортировки по марже."""
    service = database_service
    with app.app_context():
        config = app.config['APP_SETTINGS']
        unique_username = f'sort_margin_{int(time.time())}'
        user_id = service.create_user(unique_username, 'hash')
        assert user_id

        # Создаем записи с разной маржой
        base_tender = int(time.time())
        margins = [30, 15, 25, 10, 20]
        for i, margin in enumerate(margins):
            form = _make_form(i, margin_percent=str(margin), tender_number=f'TND-{base_tender}-{i}')
            service.save_generation_history(
                form,
                final_price=1000 + i * 100,
                config=config,
                user_id=user_id
            )

        # Сортировка по марже по возрастанию
        result = generation_repository.get_history(
            config,
            page=1,
            sort_by='margin_percent',
            sort_order='asc'
        )

        # Проверяем, что записи отсортированы по возрастанию маржи
        margins_found = []
        for item in result['items']:
            if item.get('tender_number', '').startswith(f'TND-{base_tender}'):
                margins_found.append(item.get('margin_percent', 0))
        
        if len(margins_found) > 1:
            assert margins_found == sorted(margins_found), "Маржа не отсортирована по возрастанию"


def test_filter_with_pagination(app):
    """Тест фильтрации с пагинацией."""
    service = database_service
    with app.app_context():
        config = app.config['APP_SETTINGS']
        unique_username = f'pagination_filter_{int(time.time())}'
        user_id = service.create_user(unique_username, 'hash')
        assert user_id

        # Создаем 15 записей с маржой 20%
        base_tender = int(time.time())
        for i in range(15):
            form = _make_form(i, margin_percent='20', tender_number=f'TND-{base_tender}-{i}')
            service.save_generation_history(
                form,
                final_price=1000 + i * 100,
                config=config,
                user_id=user_id
            )

        # Фильтр по марже 20% с пагинацией
        result_page1 = generation_repository.get_history(
            config,
            page=1,
            per_page=5,
            margin_from=20,
            margin_to=20
        )

        result_page2 = generation_repository.get_history(
            config,
            page=2,
            per_page=5,
            margin_from=20,
            margin_to=20
        )

        # Проверяем, что пагинация работает с фильтрами
        assert result_page1['pagination']['page'] == 1
        assert result_page2['pagination']['page'] == 2
        assert len(result_page1['items']) <= 5
        assert len(result_page2['items']) <= 5
        
        # Проверяем, что все записи соответствуют фильтру
        for item in result_page1['items'] + result_page2['items']:
            if item.get('tender_number', '').startswith(f'TND-{base_tender}'):
                assert item.get('margin_percent', 0) == 20, "Запись не соответствует фильтру маржи"

