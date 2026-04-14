"""add xep_loai fields to de_xuat_chi_tiet

Revision ID: f2a6e5d4b0c7
Revises: c9f2e2a1b1cd
Create Date: 2026-04-14 00:00:00.000000

"""
from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = 'f2a6e5d4b0c7'
down_revision = 'c9f2e2a1b1cd'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column('de_xuat_chi_tiet', sa.Column('xep_loai', sa.String(length=50), nullable=True))
    op.add_column('de_xuat_chi_tiet', sa.Column('xep_loai_dang_vien', sa.String(length=50), nullable=True))


def downgrade():
    op.drop_column('de_xuat_chi_tiet', 'xep_loai_dang_vien')
    op.drop_column('de_xuat_chi_tiet', 'xep_loai')
