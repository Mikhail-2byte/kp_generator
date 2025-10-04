"""Create initial tables

Revision ID: 20241004_0001
Revises: None
Create Date: 2024-10-04 00:00:00

"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = '20241004_0001'
down_revision: Union[str, None] = None
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.create_table(
        'users',
        sa.Column('id', sa.Integer(), primary_key=True, autoincrement=True),
        sa.Column('username', sa.Text(), nullable=False, unique=True),
        sa.Column('password_hash', sa.Text(), nullable=False),
        sa.Column('created_at', sa.DateTime(), server_default=sa.text('CURRENT_TIMESTAMP')),
        sa.Column('last_login', sa.DateTime()),
        sa.Column('last_name', sa.Text()),
        sa.Column('first_name', sa.Text()),
        sa.Column('role', sa.Text(), nullable=False, server_default=sa.text("'user'")),
    )

    op.create_table(
        'generation_history',
        sa.Column('id', sa.Integer(), primary_key=True, autoincrement=True),
        sa.Column('timestamp', sa.DateTime(), server_default=sa.text('CURRENT_TIMESTAMP')),
        sa.Column('tender_number', sa.Text()),
        sa.Column('company', sa.Text(), nullable=False),
        sa.Column('product', sa.Text(), nullable=False),
        sa.Column('quantity', sa.Integer(), nullable=False),
        sa.Column('cost_price', sa.Float(), nullable=False),
        sa.Column('weight', sa.Float(), nullable=False),
        sa.Column('logistics', sa.Float(), nullable=False),
        sa.Column('margin_percent', sa.Float(), nullable=False),
        sa.Column('final_price', sa.Float(), nullable=False),
        sa.Column('drawing_number', sa.Text()),
        sa.Column('material', sa.Text()),
        sa.Column('delivery_address', sa.Text()),
        sa.Column('duty_percent', sa.Float(), server_default=sa.text('0')), 
        sa.Column('delivery_time', sa.Integer(), server_default=sa.text('0')),
        sa.Column('comment', sa.Text()),
        sa.Column('user_id', sa.Integer(), sa.ForeignKey('users.id')),
    )

    op.create_index('idx_generation_history_timestamp', 'generation_history', ['timestamp'])
    op.create_index('idx_generation_history_company', 'generation_history', ['company'])
    op.create_index('idx_generation_history_user', 'generation_history', ['user_id'])


def downgrade() -> None:
    op.drop_index('idx_generation_history_user', table_name='generation_history')
    op.drop_index('idx_generation_history_company', table_name='generation_history')
    op.drop_index('idx_generation_history_timestamp', table_name='generation_history')
    op.drop_table('generation_history')
    op.drop_table('users')
