"""Bridge missing revision to align database history

Revision ID: 20260101_0009
Revises: 20260101_0008
Create Date: 2026-01-23 10:45:00.000000

Эта миграция не изменяет схему БД и служит только для
выравнивания истории Alembic в случаях, когда в таблице
alembic_version уже присутствует ревизия 20260101_0009,
но соответствующий файл миграции отсутствовал в репозитории.

Эта ситуация возникла после отката к коммиту f2b5ee4,
когда база данных содержала более новую ревизию, чем
доступные файлы миграций.
"""

from alembic import op  # noqa: F401
import sqlalchemy as sa  # noqa: F401


# revision identifiers, used by Alembic.
revision = '20260101_0009'
down_revision = '20260101_0008'
branch_labels = None
depends_on = None


def upgrade() -> None:
    """No-op upgrade: schema remains unchanged."""
    # Эта миграция намеренно оставлена пустой.
    pass


def downgrade() -> None:
    """No-op downgrade: schema remains unchanged."""
    # Откат также не вносит изменений в схему.
    pass

