"""add audit_logs table

Revision ID: 20251214_161415
Revises: 20251212_0007
Create Date: 2025-12-14 16:14:15.000000
"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = '20251214_161415'
down_revision = '20251212_0007'
branch_labels = None
depends_on = None


def upgrade() -> None:
    op.create_table(
        'audit_logs',
        sa.Column('id', sa.Integer(), primary_key=True, autoincrement=True),
        sa.Column('user_id', sa.Integer(), sa.ForeignKey('users.id'), nullable=True),
        sa.Column('username', sa.Text(), nullable=False),
        sa.Column('action_type', sa.Text(), nullable=False),
        sa.Column('resource_type', sa.Text(), nullable=True),
        sa.Column('resource_id', sa.Text(), nullable=True),
        sa.Column('description', sa.Text(), nullable=False),
        sa.Column('ip_address', sa.Text(), nullable=True),
        sa.Column('user_agent', sa.Text(), nullable=True),
        sa.Column('changes_before', sa.Text(), nullable=True),
        sa.Column('changes_after', sa.Text(), nullable=True),
        sa.Column('created_at', sa.DateTime(), server_default=sa.text('CURRENT_TIMESTAMP'), nullable=False),
    )
    
    # Индексы для быстрого поиска
    op.create_index('idx_audit_logs_user_id', 'audit_logs', ['user_id'])
    op.create_index('idx_audit_logs_action_type', 'audit_logs', ['action_type'])
    op.create_index('idx_audit_logs_resource_type', 'audit_logs', ['resource_type'])
    op.create_index('idx_audit_logs_created_at', 'audit_logs', ['created_at'])


def downgrade() -> None:
    op.drop_index('idx_audit_logs_created_at', table_name='audit_logs')
    op.drop_index('idx_audit_logs_resource_type', table_name='audit_logs')
    op.drop_index('idx_audit_logs_action_type', table_name='audit_logs')
    op.drop_index('idx_audit_logs_user_id', table_name='audit_logs')
    op.drop_table('audit_logs')

