"""add warranty_period to generation_history

Revision ID: 20251212_0007
Revises: 20251212_0006
Create Date: 2025-12-12 20:20:00.000000
"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = '20251212_0007'
down_revision = '20251212_0006'
branch_labels = None
depends_on = None


def upgrade() -> None:
    op.add_column('generation_history', sa.Column('warranty_period', sa.Text(), nullable=True))


def downgrade() -> None:
    op.drop_column('generation_history', 'warranty_period')

