"""add doan_vien and phu_nu fields to de_xuat_chi_tiet

Revision ID: c2d3e4f5a6b7
Revises: b1c2d3e4f5a6
Create Date: 2026-05-15

"""
from alembic import op
import sqlalchemy as sa

revision = 'c2d3e4f5a6b7'
down_revision = 'b1c2d3e4f5a6'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column('de_xuat_chi_tiet', sa.Column('xep_loai_doan_vien', sa.String(50), nullable=True))
    op.add_column('de_xuat_chi_tiet', sa.Column('ket_qua_phu_nu', sa.String(255), nullable=True))
    op.add_column('de_xuat_chi_tiet', sa.Column('hinh_thuc_khen_thuong_qc', sa.String(255), nullable=True))
    op.add_column('de_xuat_chi_tiet', sa.Column('hinh_thuc_khen_thuong_pn', sa.String(255), nullable=True))


def downgrade():
    op.drop_column('de_xuat_chi_tiet', 'hinh_thuc_khen_thuong_pn')
    op.drop_column('de_xuat_chi_tiet', 'hinh_thuc_khen_thuong_qc')
    op.drop_column('de_xuat_chi_tiet', 'ket_qua_phu_nu')
    op.drop_column('de_xuat_chi_tiet', 'xep_loai_doan_vien')
