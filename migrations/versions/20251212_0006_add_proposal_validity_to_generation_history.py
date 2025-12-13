"""add proposal_validity to generation_history

Revision ID: 20251212_0006
Revises: 20251212_0005
Create Date: 2025-12-12 20:10:00.000000
"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = '20251212_0006'
down_revision = '20251212_0005'
branch_labels = None
depends_on = None


def upgrade() -> None:
    op.add_column('generation_history', sa.Column('proposal_validity', sa.Text(), nullable=True))


def downgrade() -> None:
    op.drop_column('generation_history', 'proposal_validity')


