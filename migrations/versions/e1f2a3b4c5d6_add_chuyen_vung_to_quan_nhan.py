"""add_chuyen_vung_to_quan_nhan

Revision ID: e1f2a3b4c5d6
Revises: d5e6f7a8b9c0
Create Date: 2026-05-10

"""
from alembic import op
import sqlalchemy as sa

revision = 'e1f2a3b4c5d6'
down_revision = 'd5e6f7a8b9c0'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column('quan_nhan', sa.Column('is_chuyen_vung', sa.Boolean(), nullable=False, server_default='0'))
    op.add_column('quan_nhan', sa.Column('ngay_chuyen_vung', sa.DateTime(), nullable=True))


def downgrade():
    op.drop_column('quan_nhan', 'ngay_chuyen_vung')
    op.drop_column('quan_nhan', 'is_chuyen_vung')
