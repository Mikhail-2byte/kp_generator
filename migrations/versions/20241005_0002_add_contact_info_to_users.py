"""Add contact info to users

Revision ID: 20241005_0002
Revises: 20241004_0001
Create Date: 2024-10-05 00:00:00

"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = '20241005_0002'
down_revision: Union[str, None] = '20241004_0001'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.add_column('users', sa.Column('contact_info', sa.Text(), nullable=True))


def downgrade() -> None:
    op.drop_column('users', 'contact_info')
