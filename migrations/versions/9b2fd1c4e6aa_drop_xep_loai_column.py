"""drop xep_loai column

Revision ID: 9b2fd1c4e6aa
Revises: c3a1e2d9f7b4
Create Date: 2026-04-14 00:00:00.000000

"""
from alembic import op


revision = '9b2fd1c4e6aa'
down_revision = 'c3a1e2d9f7b4'
branch_labels = None
depends_on = None


def upgrade():
    op.execute("ALTER TABLE de_xuat_chi_tiet DROP COLUMN xep_loai")
    op.execute("DELETE FROM tieu_chi WHERE ma_truong = 'xep_loai'")


def downgrade():
    op.execute("ALTER TABLE de_xuat_chi_tiet ADD COLUMN xep_loai VARCHAR(50) NULL")
