import importlib

import pytest
from flask import Flask


@pytest.fixture
def orm_environment(monkeypatch, tmp_path):
    db_path = tmp_path / 'test.sqlite'
    db_url = f'sqlite:///{db_path}'
    monkeypatch.setenv('DATABASE_URL', db_url)

    import app.extensions as extensions
    import app.models as models
    import app.database as database

    extensions = importlib.reload(extensions)
    models = importlib.reload(models)
    database = importlib.reload(database)

    app = Flask(__name__)
    engine = extensions.init_db_engine(app)
    models.Base.metadata.create_all(bind=engine)

    yield database, extensions, models

    models.Base.metadata.drop_all(bind=engine)
    extensions.SessionLocal.remove()


def test_create_and_fetch_user(orm_environment):
    database, extensions, models = orm_environment

    user_id = database.create_user('tester', 'hash', 'Doe', 'John', role='Admin')
    assert user_id is not None

    by_id = database.get_user_by_id(user_id)
    assert by_id[1] == 'tester'
    assert by_id[7] is None
    assert by_id[8] == 'admin'

    by_username = database.get_user_by_username('tester')
    assert by_username[0] == user_id

    assert database.update_last_login(user_id)

    updated = database.get_user_by_id(user_id)
    assert updated[4] is not None

    contact_line = 'John Doe | +7 999 111-22-33 | john@example.com'
    assert database.update_user_profile(user_id, 'tester', 'Doe', 'John', contact_line)

    refreshed = database.get_user_by_id(user_id)
    assert refreshed[7] == contact_line


def test_generation_history_flow(orm_environment):
    database, extensions, models = orm_environment

    user_id = database.create_user('author', 'hash')
    assert user_id is not None

    config = {'margin_percent': 25, 'default_duty_percent': 5}
    payload = {
        'tender_number': 'TN-001',
        'company': 'Acme',
        'product': 'Widget',
        'quantity': 10,
        'cost_price': 100,
        'weight': 2.5,
        'logistics': 300,
        'margin_percent': 30,
        'drawing_number': 'DR-1',
        'material': 'Steel',
        'delivery_address': 'Earth',
        'duty_percent': 7,
        'delivery_time': 14,
        'comment': 'test',
    }

    assert database.save_generation_history(payload, final_price=150.5, config=config, user_id=user_id)

    history = database.get_generation_history({'max_history_items': 5})
    assert len(history) == 1
    assert history[0]['company'] == 'Acme'
    assert history[0]['username'] == 'author'

    details = database.get_generation_details(history[0]['id'])
    assert details['product'] == 'Widget'
    assert details['user_id'] == user_id

    loaded = database.load_generation_data(history[0]['id'])
    assert loaded['quantity'] == 10


def test_delete_user_detaches_history(orm_environment):
    database, extensions, models = orm_environment

    user_id = database.create_user('erasable', 'hash')
    database.save_generation_history(
        {
            'company': 'Corp',
            'product': 'Thing',
            'quantity': 1,
            'cost_price': 10,
            'weight': 1,
            'logistics': 5,
            'margin_percent': 10,
            'delivery_time': 3,
        },
        final_price=12.3,
        config={},
        user_id=user_id,
    )

    assert database.delete_user(user_id)

    history = database.get_generation_history({'max_history_items': 5})
    assert history[0]['username'] is None

    assert database.get_user_by_id(user_id) is None
