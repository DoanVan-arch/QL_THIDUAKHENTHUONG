"""add_lop_to_quan_nhan

Revision ID: d5e6f7a8b9c0
Revises: c4d5e6f7a8b9
Create Date: 2026-05-07
"""
from alembic import op
import sqlalchemy as sa

revision = 'd5e6f7a8b9c0'
down_revision = 'c4d5e6f7a8b9'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column('quan_nhan', sa.Column('lop', sa.String(100), nullable=True))


def downgrade():
    op.drop_column('quan_nhan', 'lop')
