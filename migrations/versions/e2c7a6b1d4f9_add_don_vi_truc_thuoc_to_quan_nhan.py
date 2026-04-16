"""add don_vi_truc_thuoc to quan_nhan

Revision ID: e2c7a6b1d4f9
Revises: b4f1a5d6c2e9
Create Date: 2026-04-15 00:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


revision = 'e2c7a6b1d4f9'
down_revision = 'b4f1a5d6c2e9'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column('quan_nhan', sa.Column('don_vi_truc_thuoc', sa.String(length=150), nullable=True))


def downgrade():
    op.drop_column('quan_nhan', 'don_vi_truc_thuoc')
