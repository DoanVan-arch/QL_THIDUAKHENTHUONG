"""add bi_loai soft-remove fields to de_xuat_chi_tiet

Revision ID: k8l9m0n1o2p3
Revises: j7k8l9m0n1o2
Create Date: 2026-06-03
"""
from alembic import op
import sqlalchemy as sa

revision = 'k8l9m0n1o2p3'
down_revision = 'j7k8l9m0n1o2'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column('de_xuat_chi_tiet', sa.Column('bi_loai', sa.Boolean(), nullable=False, server_default='0'))
    op.add_column('de_xuat_chi_tiet', sa.Column('ly_do_loai', sa.Text(), nullable=True))
    op.add_column('de_xuat_chi_tiet', sa.Column('phong_loai', sa.String(length=100), nullable=True))
    op.add_column('de_xuat_chi_tiet', sa.Column('ngay_loai', sa.DateTime(), nullable=True))


def downgrade():
    op.drop_column('de_xuat_chi_tiet', 'ngay_loai')
    op.drop_column('de_xuat_chi_tiet', 'phong_loai')
    op.drop_column('de_xuat_chi_tiet', 'ly_do_loai')
    op.drop_column('de_xuat_chi_tiet', 'bi_loai')
