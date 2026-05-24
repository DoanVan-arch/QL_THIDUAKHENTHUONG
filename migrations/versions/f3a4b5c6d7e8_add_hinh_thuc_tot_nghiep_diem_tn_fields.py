"""add hinh_thuc_tot_nghiep and diem_tn fields to de_xuat_chi_tiet

Revision ID: f3a4b5c6d7e8
Revises: e2f3a4b5c6d7
Create Date: 2026-05-24
"""
from alembic import op
import sqlalchemy as sa

revision = 'f3a4b5c6d7e8'
down_revision = 'e2f3a4b5c6d7'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column('de_xuat_chi_tiet', sa.Column('hinh_thuc_tot_nghiep', sa.String(100), nullable=True))
    op.add_column('de_xuat_chi_tiet', sa.Column('diem_tn_ctd', sa.String(50), nullable=True))
    op.add_column('de_xuat_chi_tiet', sa.Column('diem_tn_ct', sa.String(50), nullable=True))
    op.add_column('de_xuat_chi_tiet', sa.Column('diem_tn_ta', sa.String(50), nullable=True))
    op.add_column('de_xuat_chi_tiet', sa.Column('diem_tn_mon4', sa.String(50), nullable=True))
    op.add_column('de_xuat_chi_tiet', sa.Column('diem_tn_chuyennganh', sa.String(50), nullable=True))
    op.add_column('de_xuat_chi_tiet', sa.Column('diem_tn_baove', sa.String(50), nullable=True))


def downgrade():
    op.drop_column('de_xuat_chi_tiet', 'diem_tn_baove')
    op.drop_column('de_xuat_chi_tiet', 'diem_tn_chuyennganh')
    op.drop_column('de_xuat_chi_tiet', 'diem_tn_mon4')
    op.drop_column('de_xuat_chi_tiet', 'diem_tn_ta')
    op.drop_column('de_xuat_chi_tiet', 'diem_tn_ct')
    op.drop_column('de_xuat_chi_tiet', 'diem_tn_ctd')
    op.drop_column('de_xuat_chi_tiet', 'hinh_thuc_tot_nghiep')
