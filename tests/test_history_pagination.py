from app.database import database_service
from app.services.repositories import generation_repository


def _make_form(i: int) -> dict:
    return {
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


def test_history_pagination(app):
    service = database_service
    with app.app_context():
        config = app.config['APP_SETTINGS']
        user_id = service.create_user('pager', 'hash')
        assert user_id

        # создаём 30 записей истории
        for i in range(30):
            form = _make_form(i)
            service.save_generation_history(form, final_price=1500 + i, config=config, user_id=user_id)

        result = generation_repository.get_history(config, page=2, per_page=10)

        assert result['pagination']['page'] == 2
        assert result['pagination']['per_page'] == 10
        assert result['pagination']['total'] == 30
        assert len(result['items']) == 10
        assert result['pagination']['has_prev'] is True
        assert result['pagination']['has_next'] is True
        assert all(item.get('version_count', 1) == 1 for item in result['items'])


def test_history_tender_endpoint(app, client):
    service = database_service
    with app.app_context():
        config = app.config['APP_SETTINGS']
        user_id = service.create_user('tender-user', 'hash')
        assert user_id

        form = _make_form(101)
        service.save_generation_history(form, final_price=2000, config=config, user_id=user_id)
        service.save_generation_history(form, final_price=2200, config=config, user_id=user_id)

    response = client.get('/history/tender/TND-101')
    assert response.status_code == 200
    payload = response.get_json()
    assert len(payload['records']) >= 2

