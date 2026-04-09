"""add_danh_hieu_table

Revision ID: d952c1a0e20e
Revises: 478e22fbb445
Create Date: 2026-04-08 16:05:26.104961

"""
from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = 'd952c1a0e20e'
down_revision = '478e22fbb445'
branch_labels = None
depends_on = None


def upgrade():
    op.create_table('danh_hieu',
    sa.Column('id', sa.Integer(), autoincrement=True, nullable=False),
    sa.Column('ten_danh_hieu', sa.String(length=100), nullable=False),
    sa.Column('ma_danh_hieu', sa.String(length=20), nullable=False),
    sa.Column('pham_vi', sa.String(length=20), nullable=False),
    sa.Column('tieu_chi', sa.Text(), nullable=True),
    sa.Column('thu_tu', sa.Integer(), nullable=True),
    sa.Column('is_active', sa.Boolean(), nullable=True),
    sa.Column('created_at', sa.DateTime(), nullable=True),
    sa.Column('updated_at', sa.DateTime(), nullable=True),
    sa.PrimaryKeyConstraint('id'),
    sa.UniqueConstraint('ma_danh_hieu'),
    sa.UniqueConstraint('ten_danh_hieu')
    )


def downgrade():
    op.drop_table('danh_hieu')
