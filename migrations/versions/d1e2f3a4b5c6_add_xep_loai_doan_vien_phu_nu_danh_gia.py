"""add xep_loai_doan_vien phu_nu to danh_gia_hang_nam

Revision ID: d1e2f3a4b5c6
Revises: c2d3e4f5a6b7
Create Date: 2026-05-16
"""
from alembic import op
import sqlalchemy as sa

revision = 'd1e2f3a4b5c6'
down_revision = 'c2d3e4f5a6b7'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column('danh_gia_hang_nam',
        sa.Column('xep_loai_doan_vien', sa.String(100), nullable=True))
    op.add_column('danh_gia_hang_nam',
        sa.Column('xep_loai_phu_nu', sa.String(100), nullable=True))


def downgrade():
    op.drop_column('danh_gia_hang_nam', 'xep_loai_phu_nu')
    op.drop_column('danh_gia_hang_nam', 'xep_loai_doan_vien')
