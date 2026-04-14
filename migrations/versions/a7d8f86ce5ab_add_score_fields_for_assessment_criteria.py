"""add score fields for assessment criteria

Revision ID: a7d8f86ce5ab
Revises: f4e2dcb8c3aa
Create Date: 2026-04-14 00:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


revision = 'a7d8f86ce5ab'
down_revision = 'f4e2dcb8c3aa'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column('de_xuat_chi_tiet', sa.Column('diem_kiem_tra_tin_hoc', sa.String(length=20), nullable=True))
    op.add_column('de_xuat_chi_tiet', sa.Column('diem_kiem_tra_dieu_lenh', sa.String(length=20), nullable=True))
    op.add_column('de_xuat_chi_tiet', sa.Column('diem_dia_ly_quan_su', sa.String(length=20), nullable=True))
    op.add_column('de_xuat_chi_tiet', sa.Column('diem_ban_sung', sa.String(length=20), nullable=True))
    op.add_column('de_xuat_chi_tiet', sa.Column('diem_the_luc', sa.String(length=20), nullable=True))
    op.add_column('de_xuat_chi_tiet', sa.Column('diem_kiem_tra_chinh_tri', sa.String(length=20), nullable=True))


def downgrade():
    op.drop_column('de_xuat_chi_tiet', 'diem_kiem_tra_chinh_tri')
    op.drop_column('de_xuat_chi_tiet', 'diem_the_luc')
    op.drop_column('de_xuat_chi_tiet', 'diem_ban_sung')
    op.drop_column('de_xuat_chi_tiet', 'diem_dia_ly_quan_su')
    op.drop_column('de_xuat_chi_tiet', 'diem_kiem_tra_dieu_lenh')
    op.drop_column('de_xuat_chi_tiet', 'diem_kiem_tra_tin_hoc')
