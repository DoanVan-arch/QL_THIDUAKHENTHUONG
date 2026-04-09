"""add_tieu_chi_table

Revision ID: 74ea7da7de03
Revises: d952c1a0e20e
Create Date: 2026-04-08 16:38:59.812658

"""
from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = '74ea7da7de03'
down_revision = 'd952c1a0e20e'
branch_labels = None
depends_on = None


def upgrade():
    op.create_table('tieu_chi',
    sa.Column('id', sa.Integer(), autoincrement=True, nullable=False),
    sa.Column('ma_truong', sa.String(length=50), nullable=False),
    sa.Column('ten', sa.String(length=150), nullable=False),
    sa.Column('huong_dan', sa.Text(), nullable=True),
    sa.Column('nhom', sa.String(length=50), nullable=False),
    sa.Column('co_minh_chung', sa.Boolean(), nullable=True),
    sa.Column('phong_duyet', sa.Text(), nullable=True),
    sa.Column('thu_tu', sa.Integer(), nullable=True),
    sa.Column('is_active', sa.Boolean(), nullable=True),
    sa.Column('created_at', sa.DateTime(), nullable=True),
    sa.Column('updated_at', sa.DateTime(), nullable=True),
    sa.PrimaryKeyConstraint('id'),
    sa.UniqueConstraint('ma_truong')
    )


def downgrade():
    op.drop_table('tieu_chi')
