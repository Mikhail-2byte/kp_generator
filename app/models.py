from datetime import datetime
from flask_login import UserMixin


class User(UserMixin):
    def __init__(
        self,
        user_id,
        username,
        password_hash,
        created_at=None,
        last_login=None,
        last_name=None,
        first_name=None,
        role='user'
    ):
        self.id = str(user_id)
        self.username = username
        self.password_hash = password_hash
        self.created_at = created_at
        self.last_login = last_login
        self.last_name = last_name
        self.first_name = first_name
        self.role = (role or 'user').lower()

    @classmethod
    def from_row(cls, row):
        if not row:
            return None
        last_name = row[5] if len(row) > 5 else None
        first_name = row[6] if len(row) > 6 else None
        role = row[7] if len(row) > 7 else 'user'
        return cls(
            user_id=row[0],
            username=row[1],
            password_hash=row[2],
            created_at=row[3],
            last_login=row[4],
            last_name=last_name,
            first_name=first_name,
            role=role
        )

    @property
    def is_admin(self) -> bool:
        return self.role == 'admin'

    def set_last_login_now(self):
        self.last_login = datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')
