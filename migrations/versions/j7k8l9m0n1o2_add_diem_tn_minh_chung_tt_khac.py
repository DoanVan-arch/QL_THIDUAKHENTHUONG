"""add diem_tot_nghiep and minh_chung_thanh_tich_khac to de_xuat_chi_tiet

Revision ID: j7k8l9m0n1o2
Revises: i6j7k8l9m0n1
Create Date: 2026-06-02
"""
from alembic import op
import sqlalchemy as sa

revision = 'j7k8l9m0n1o2'
down_revision = 'i6j7k8l9m0n1'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column('de_xuat_chi_tiet', sa.Column('diem_tot_nghiep', sa.Float(), nullable=True))
    op.add_column('de_xuat_chi_tiet', sa.Column('minh_chung_thanh_tich_khac', sa.String(length=255), nullable=True))


def downgrade():
    op.drop_column('de_xuat_chi_tiet', 'minh_chung_thanh_tich_khac')
    op.drop_column('de_xuat_chi_tiet', 'diem_tot_nghiep')
