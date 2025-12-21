"""Add customer_contacts table

Revision ID: 20250115_0001
Revises: 20251214_161415
Create Date: 2025-01-15 12:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = '20250115_0001'
down_revision = '20251214_161415'
branch_labels = None
depends_on = None


def upgrade():
    op.create_table(
        'customer_contacts',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('company_name', sa.Text(), nullable=False),
        sa.Column('contact_person', sa.Text(), nullable=True),
        sa.Column('phone', sa.Text(), nullable=True),
        sa.Column('email', sa.Text(), nullable=True),
        sa.Column('address', sa.Text(), nullable=True),
        sa.Column('notes', sa.Text(), nullable=True),
        sa.Column('created_at', sa.DateTime(), server_default=sa.text('CURRENT_TIMESTAMP'), nullable=True),
        sa.Column('updated_at', sa.DateTime(), server_default=sa.text('CURRENT_TIMESTAMP'), nullable=True),
        sa.PrimaryKeyConstraint('id')
    )
    op.create_index('ix_customer_contacts_company_name', 'customer_contacts', ['company_name'], unique=False)


def downgrade():
    op.drop_index('ix_customer_contacts_company_name', table_name='customer_contacts')
    op.drop_table('customer_contacts')

