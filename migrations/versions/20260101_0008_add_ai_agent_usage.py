"""Add ai_agent_usage table for monitoring API usage

Revision ID: 20260101_0008
Revises: 20250115_0003
Create Date: 2026-01-01 20:00:00.000000

"""
from alembic import op
import sqlalchemy as sa
from sqlalchemy.sql import func


# revision identifiers, used by Alembic.
revision = '20260101_0008'
down_revision = '20250115_0003'
branch_labels = None
depends_on = None


def upgrade():
    """Create ai_agent_usage table for tracking API usage and performance."""
    op.create_table(
        'ai_agent_usage',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('timestamp', sa.DateTime(), server_default=func.now(), nullable=False),
        sa.Column('user_id', sa.Integer(), nullable=True),
        sa.Column('request_tokens', sa.Integer(), nullable=True),
        sa.Column('response_tokens', sa.Integer(), nullable=True),
        sa.Column('total_cost', sa.Float(), nullable=True),
        sa.Column('response_time_ms', sa.Integer(), nullable=True),
        sa.Column('model_name', sa.String(length=100), nullable=True),
        sa.Column('error_type', sa.String(length=50), nullable=True),
        sa.Column('user_message', sa.Text(), nullable=True),
        sa.Column('cache_hit', sa.Boolean(), default=False, nullable=False),
        sa.PrimaryKeyConstraint('id'),
        sa.ForeignKeyConstraint(['user_id'], ['users.id'], ondelete='SET NULL')
    )
    
    # Создаем индексы для оптимизации запросов
    op.create_index('idx_ai_agent_usage_timestamp', 'ai_agent_usage', ['timestamp'])
    op.create_index('idx_ai_agent_usage_user_id', 'ai_agent_usage', ['user_id'])
    op.create_index('idx_ai_agent_usage_error_type', 'ai_agent_usage', ['error_type'])
    op.create_index('idx_ai_agent_usage_model_name', 'ai_agent_usage', ['model_name'])


def downgrade():
    """Drop ai_agent_usage table and its indexes."""
    op.drop_index('idx_ai_agent_usage_model_name', table_name='ai_agent_usage')
    op.drop_index('idx_ai_agent_usage_error_type', table_name='ai_agent_usage')
    op.drop_index('idx_ai_agent_usage_user_id', table_name='ai_agent_usage')
    op.drop_index('idx_ai_agent_usage_timestamp', table_name='ai_agent_usage')
    op.drop_table('ai_agent_usage')

