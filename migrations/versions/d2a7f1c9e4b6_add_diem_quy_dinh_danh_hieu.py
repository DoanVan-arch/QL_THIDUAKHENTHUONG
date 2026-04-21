"""add diem quy dinh danh hieu

Revision ID: d2a7f1c9e4b6
Revises: c1a9b7d4e2f3
Create Date: 2026-04-21 14:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


revision = 'd2a7f1c9e4b6'
down_revision = 'c1a9b7d4e2f3'
branch_labels = None
depends_on = None


def upgrade():
    op.create_table(
        'diem_quy_dinh_danh_hieu',
        sa.Column('id', sa.Integer(), autoincrement=True, nullable=False),
        sa.Column('loai_danh_hieu', sa.String(length=100), nullable=False),
        sa.Column('tieu_chi_field', sa.String(length=100), nullable=False),
        sa.Column('min_diem', sa.String(length=20), nullable=False),
        sa.Column('is_active', sa.Boolean(), nullable=True),
        sa.Column('created_at', sa.DateTime(), nullable=True),
        sa.Column('updated_at', sa.DateTime(), nullable=True),
        sa.PrimaryKeyConstraint('id'),
        sa.UniqueConstraint('loai_danh_hieu', 'tieu_chi_field', name='uq_diem_rule_danh_hieu_field'),
    )
    op.create_index(op.f('ix_diem_quy_dinh_danh_hieu_loai_danh_hieu'), 'diem_quy_dinh_danh_hieu', ['loai_danh_hieu'], unique=False)
    op.create_index(op.f('ix_diem_quy_dinh_danh_hieu_tieu_chi_field'), 'diem_quy_dinh_danh_hieu', ['tieu_chi_field'], unique=False)


def downgrade():
    op.drop_index(op.f('ix_diem_quy_dinh_danh_hieu_tieu_chi_field'), table_name='diem_quy_dinh_danh_hieu')
    op.drop_index(op.f('ix_diem_quy_dinh_danh_hieu_loai_danh_hieu'), table_name='diem_quy_dinh_danh_hieu')
    op.drop_table('diem_quy_dinh_danh_hieu')
