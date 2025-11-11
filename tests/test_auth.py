from werkzeug.security import generate_password_hash

from app.core.extensions import SessionLocal
from app.models.models import UserRecord


def _create_user(username: str, password: str) -> None:
    session = SessionLocal()
    try:
        user = UserRecord(
            username=username,
            password_hash=generate_password_hash(password),
            role='user',
        )
        session.add(user)
        session.commit()
    finally:
        SessionLocal.remove()


def test_login_flow(client, app):
    _create_user('test@example.com', 'Secret123!')

    response = client.post(
        '/profile',
        data={
            'login-username': 'test@example.com',
            'login-password': 'Secret123!',
            'login-submit_login': 'Войти',
        },
        follow_redirects=True,
    )

    body = response.get_data(as_text=True)
    assert response.status_code == 200
    assert 'Вы успешно вошли в систему.' in body

