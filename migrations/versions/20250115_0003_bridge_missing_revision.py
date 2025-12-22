"""Bridge missing revision to align database history

Revision ID: 20250115_0003
Revises: 20250115_0002
Create Date: 2025-01-15 13:00:00.000000

Эта миграция не изменяет схему БД и служит только для
выравнивания истории Alembic в случаях, когда в таблице
alembic_version уже присутствует ревизия 20250115_0003,
но соответствующий файл миграции отсутствовал в репозитории.
"""

from alembic import op  # noqa: F401
import sqlalchemy as sa  # noqa: F401


# revision identifiers, used by Alembic.
revision = '20250115_0003'
down_revision = '20250115_0002'
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


