"""add doi tuong option table

Revision ID: c1a9b7d4e2f3
Revises: 8c2d1a4f6b77
Create Date: 2026-04-21 11:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


revision = 'c1a9b7d4e2f3'
down_revision = '8c2d1a4f6b77'
branch_labels = None
depends_on = None


def upgrade():
    op.create_table(
        'doi_tuong_option',
        sa.Column('id', sa.Integer(), autoincrement=True, nullable=False),
        sa.Column('ten', sa.String(length=120), nullable=False),
        sa.Column('thu_tu', sa.Integer(), nullable=True),
        sa.Column('is_active', sa.Boolean(), nullable=True),
        sa.Column('created_at', sa.DateTime(), nullable=True),
        sa.Column('updated_at', sa.DateTime(), nullable=True),
        sa.PrimaryKeyConstraint('id'),
        sa.UniqueConstraint('ten'),
    )


def downgrade():
    op.drop_table('doi_tuong_option')
