"""Optimize database indexes for better query performance

Revision ID: 20250115_0002
Revises: 20250115_0001
Create Date: 2025-01-15 12:30:00.000000

"""
from alembic import op
import sqlalchemy as sa
from sqlalchemy import inspect


# revision identifiers, used by Alembic.
revision = '20250115_0002'
down_revision = '20250115_0001'
branch_labels = None
depends_on = None


def _index_exists(table_name: str, index_name: str) -> bool:
    """Проверяет существование индекса в таблице."""
    bind = op.get_bind()
    inspector = inspect(bind)
    indexes = inspector.get_indexes(table_name)
    return any(idx['name'] == index_name for idx in indexes)


def upgrade() -> None:
    # Индексы для generation_history (дополнительные для оптимизации запросов)
    # Проверяем, не создан ли уже индекс с другим именем
    if not _index_exists('generation_history', 'ix_generation_history_company') and \
       not _index_exists('generation_history', 'idx_generation_history_company'):
        op.create_index('ix_generation_history_company', 'generation_history', ['company'], unique=False)
    
    if not _index_exists('generation_history', 'ix_generation_history_tender_number'):
        op.create_index('ix_generation_history_tender_number', 'generation_history', ['tender_number'], unique=False)
    
    # Составной индекс для частых запросов по user_id и timestamp
    if not _index_exists('generation_history', 'ix_generation_history_user_timestamp'):
        op.create_index('ix_generation_history_user_timestamp', 'generation_history', ['user_id', 'timestamp'], unique=False)
    
    # Индексы для users (для поиска и фильтрации)
    if not _index_exists('users', 'ix_users_username'):
        op.create_index('ix_users_username', 'users', ['username'], unique=True)
    
    if not _index_exists('users', 'ix_users_role'):
        op.create_index('ix_users_role', 'users', ['role'], unique=False)
    
    # Индексы для customer_contacts уже созданы в миграции 20250115_0001
    
    # Индексы для audit_logs (дополнительные составные индексы)
    if not _index_exists('audit_logs', 'ix_audit_logs_user_created'):
        op.create_index('ix_audit_logs_user_created', 'audit_logs', ['user_id', 'created_at'], unique=False)
    
    if not _index_exists('audit_logs', 'ix_audit_logs_resource'):
        op.create_index('ix_audit_logs_resource', 'audit_logs', ['resource_type', 'resource_id'], unique=False)


def downgrade() -> None:
    op.drop_index('ix_audit_logs_resource', table_name='audit_logs')
    op.drop_index('ix_audit_logs_user_created', table_name='audit_logs')
    # Индекс ix_customer_contacts_company_name удаляется в миграции 20250115_0001
    op.drop_index('ix_users_role', table_name='users')
    op.drop_index('ix_users_username', table_name='users')
    op.drop_index('ix_generation_history_user_timestamp', table_name='generation_history')
    op.drop_index('ix_generation_history_tender_number', table_name='generation_history')
    op.drop_index('ix_generation_history_company', table_name='generation_history')

