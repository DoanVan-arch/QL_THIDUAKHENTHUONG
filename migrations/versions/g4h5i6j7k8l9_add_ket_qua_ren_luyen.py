"""add ket_qua_ren_luyen to de_xuat_chi_tiet

Revision ID: g4h5i6j7k8l9
Revises: f3a4b5c6d7e8
Create Date: 2026-05-24
"""
from alembic import op
import sqlalchemy as sa

revision = 'g4h5i6j7k8l9'
down_revision = 'f3a4b5c6d7e8'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column('de_xuat_chi_tiet', sa.Column('ket_qua_ren_luyen', sa.String(100), nullable=True))


def downgrade():
    op.drop_column('de_xuat_chi_tiet', 'ket_qua_ren_luyen')
