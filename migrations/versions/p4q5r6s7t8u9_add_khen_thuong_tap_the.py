"""add khen_thuong_tap_the table

Revision ID: p4q5r6s7t8u9
Revises: 0874451bc2b5
Create Date: 2026-09-10
"""
from alembic import op
import sqlalchemy as sa

revision = 'p4q5r6s7t8u9'
down_revision = '0874451bc2b5'
branch_labels = None
depends_on = None


def upgrade():
    op.create_table(
        'khen_thuong_tap_the',
        sa.Column('id', sa.Integer(), autoincrement=True, nullable=False),
        sa.Column('don_vi_id', sa.Integer(), nullable=False),
        sa.Column('khen_thuong_id', sa.Integer(), nullable=True),
        sa.Column('nguon', sa.String(length=20), nullable=False, server_default='thu_cong'),
        sa.Column('loai', sa.String(length=50), nullable=True),
        sa.Column('ten', sa.String(length=255), nullable=False),
        sa.Column('so_quyet_dinh', sa.String(length=100), nullable=True),
        sa.Column('ngay_cap', sa.Date(), nullable=True),
        sa.Column('nam_hoc', sa.String(length=20), nullable=True),
        sa.Column('cap_quyet_dinh', sa.String(length=255), nullable=True),
        sa.Column('anh_minh_chung', sa.String(length=255), nullable=True),
        sa.Column('created_by_id', sa.Integer(), nullable=True),
        sa.Column('created_at', sa.DateTime(), nullable=True),
        sa.Column('updated_at', sa.DateTime(), nullable=True),
        sa.ForeignKeyConstraint(['don_vi_id'], ['don_vi.id']),
        sa.ForeignKeyConstraint(['khen_thuong_id'], ['khen_thuong.id']),
        sa.ForeignKeyConstraint(['created_by_id'], ['users.id']),
        sa.PrimaryKeyConstraint('id'),
    )
    op.create_index(op.f('ix_khen_thuong_tap_the_don_vi_id'), 'khen_thuong_tap_the', ['don_vi_id'], unique=False)
    op.create_index(op.f('ix_khen_thuong_tap_the_khen_thuong_id'), 'khen_thuong_tap_the', ['khen_thuong_id'], unique=False)


def downgrade():
    op.drop_index(op.f('ix_khen_thuong_tap_the_khen_thuong_id'), table_name='khen_thuong_tap_the')
    op.drop_index(op.f('ix_khen_thuong_tap_the_don_vi_id'), table_name='khen_thuong_tap_the')
    op.drop_table('khen_thuong_tap_the')
