"""add indexes for generation history lookups

Revision ID: 20241222_0004
Revises: 20241220_0003
Create Date: 2025-11-20 13:05:00.000000
"""

from alembic import op


# revision identifiers, used by Alembic.
revision = '20241222_0004'
down_revision = '20241220_0003'
branch_labels = None
depends_on = None


def upgrade() -> None:
    op.create_index('ix_generation_history_timestamp', 'generation_history', ['timestamp'])
    op.create_index('ix_generation_history_drawing_number', 'generation_history', ['drawing_number'])
    op.create_index('ix_generation_history_user_id', 'generation_history', ['user_id'])


def downgrade() -> None:
    op.drop_index('ix_generation_history_user_id', table_name='generation_history')
    op.drop_index('ix_generation_history_drawing_number', table_name='generation_history')
    op.drop_index('ix_generation_history_timestamp', table_name='generation_history')

