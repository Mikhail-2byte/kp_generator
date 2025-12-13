"""add payment_terms to generation_history

Revision ID: 20251212_0005
Revises: 20241222_0004
Create Date: 2025-12-12 20:00:00.000000
"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = '20251212_0005'
down_revision = '20241222_0004'
branch_labels = None
depends_on = None


def upgrade() -> None:
    op.add_column('generation_history', sa.Column('payment_terms', sa.Text(), nullable=True))


def downgrade() -> None:
    op.drop_column('generation_history', 'payment_terms')


