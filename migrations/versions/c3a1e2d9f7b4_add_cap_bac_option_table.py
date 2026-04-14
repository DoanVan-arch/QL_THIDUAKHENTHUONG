"""add cap_bac_option table

Revision ID: c3a1e2d9f7b4
Revises: a7d8f86ce5ab
Create Date: 2026-04-14 00:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


revision = 'c3a1e2d9f7b4'
down_revision = 'a7d8f86ce5ab'
branch_labels = None
depends_on = None


def upgrade():
    op.create_table(
        'cap_bac_option',
        sa.Column('id', sa.Integer(), autoincrement=True, nullable=False),
        sa.Column('ten', sa.String(length=120), nullable=False),
        sa.Column('thu_tu', sa.Integer(), nullable=True),
        sa.Column('is_active', sa.Boolean(), nullable=True),
        sa.Column('created_at', sa.DateTime(), nullable=True),
        sa.Column('updated_at', sa.DateTime(), nullable=True),
        sa.PrimaryKeyConstraint('id'),
        sa.UniqueConstraint('ten')
    )

    default_values = [
        'Binh nhì', 'Binh nhất', 'Hạ sĩ', 'Trung sĩ', 'Thượng sĩ',
        'Thiếu úy', 'Trung úy', 'Thượng úy', 'Đại úy',
        'Thiếu tá', 'Trung tá', 'Thượng tá', 'Đại tá'
    ]
    for i, ten in enumerate(default_values, start=1):
        op.execute(
            sa.text(
                "INSERT INTO cap_bac_option (ten, thu_tu, is_active, created_at, updated_at) "
                "VALUES (:ten, :thu_tu, 1, NOW(), NOW())"
            ).bindparams(ten=ten, thu_tu=i)
        )


def downgrade():
    op.drop_table('cap_bac_option')
