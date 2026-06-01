"""add mo_ta_khoa_hoc to de_xuat_chi_tiet

Revision ID: i6j7k8l9m0n1
Revises: h5i6j7k8l9m0
Create Date: 2026-06-01
"""
from alembic import op
import sqlalchemy as sa

revision = 'i6j7k8l9m0n1'
down_revision = 'h5i6j7k8l9m0'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column('de_xuat_chi_tiet', sa.Column('mo_ta_khoa_hoc', sa.Text(), nullable=True))


def downgrade():
    op.drop_column('de_xuat_chi_tiet', 'mo_ta_khoa_hoc')
