"""add admin_approved to de_xuat_chi_tiet

Revision ID: b2c3d4e5f6a1
Revises: a1b2c3d4e5f6
Create Date: 2026-04-28

"""
from alembic import op
import sqlalchemy as sa

revision = 'b2c3d4e5f6a1'
down_revision = 'a1b2c3d4e5f6'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column(
        'de_xuat_chi_tiet',
        sa.Column('admin_approved', sa.Boolean(), nullable=False, server_default='0')
    )


def downgrade():
    op.drop_column('de_xuat_chi_tiet', 'admin_approved')
