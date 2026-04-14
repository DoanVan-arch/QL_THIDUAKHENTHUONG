"""add can_cuoc_cong_dan to quan_nhan

Revision ID: c9f2e2a1b1cd
Revises: 74ea7da7de03
Create Date: 2026-04-14 00:00:00.000000

"""
from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = 'c9f2e2a1b1cd'
down_revision = '74ea7da7de03'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column('quan_nhan', sa.Column('can_cuoc_cong_dan', sa.String(length=20), nullable=True))


def downgrade():
    op.drop_column('quan_nhan', 'can_cuoc_cong_dan')
