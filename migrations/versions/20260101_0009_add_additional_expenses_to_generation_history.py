"""add additional_expenses to generation_history

Revision ID: 20260101_0009
Revises: 20260101_0008
Create Date: 2026-01-01 12:00:00.000000
"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = '20260101_0009'
down_revision = '20260101_0008'
branch_labels = None
depends_on = None


def upgrade() -> None:
    op.add_column('generation_history', sa.Column('additional_expenses', sa.Text(), nullable=True))


def downgrade() -> None:
    op.drop_column('generation_history', 'additional_expenses')

