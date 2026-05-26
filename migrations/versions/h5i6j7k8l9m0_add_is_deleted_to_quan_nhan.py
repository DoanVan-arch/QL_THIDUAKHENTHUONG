"""add is_deleted fields to quan_nhan

Revision ID: h5i6j7k8l9m0
Revises: g4h5i6j7k8l9
Create Date: 2026-05-26
"""
from alembic import op
import sqlalchemy as sa

revision = 'h5i6j7k8l9m0'
down_revision = 'g4h5i6j7k8l9'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column('quan_nhan', sa.Column('is_deleted', sa.Boolean(), nullable=False, server_default='0'))
    op.add_column('quan_nhan', sa.Column('deleted_at', sa.DateTime(), nullable=True))
    op.add_column('quan_nhan', sa.Column('deleted_by_id', sa.Integer(), nullable=True))


def downgrade():
    op.drop_column('quan_nhan', 'deleted_by_id')
    op.drop_column('quan_nhan', 'deleted_at')
    op.drop_column('quan_nhan', 'is_deleted')
