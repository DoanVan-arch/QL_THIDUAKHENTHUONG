"""widen trang_thai to varchar(50)

Revision ID: b1c2d3e4f5a6
Revises: f4e2dcb8c3aa
Create Date: 2026-05-14
"""
from alembic import op
import sqlalchemy as sa

revision = 'b1c2d3e4f5a6'
down_revision = 'a2b3c4d5e6f7'
branch_labels = None
depends_on = None


def upgrade():
    op.alter_column('de_xuat', 'trang_thai',
                    existing_type=sa.String(30),
                    type_=sa.String(50),
                    existing_nullable=True)


def downgrade():
    op.alter_column('de_xuat', 'trang_thai',
                    existing_type=sa.String(50),
                    type_=sa.String(30),
                    existing_nullable=True)
