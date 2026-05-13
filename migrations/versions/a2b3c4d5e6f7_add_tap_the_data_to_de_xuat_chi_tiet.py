"""add_tap_the_data_to_de_xuat_chi_tiet

Revision ID: a2b3c4d5e6f7
Revises: f1a2b3c4d5e6
Create Date: 2026-05-13

"""
from alembic import op
import sqlalchemy as sa

revision = 'a2b3c4d5e6f7'
down_revision = 'f1a2b3c4d5e6'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column('de_xuat_chi_tiet',
        sa.Column('tap_the_data', sa.Text(), nullable=True)
    )


def downgrade():
    op.drop_column('de_xuat_chi_tiet', 'tap_the_data')
