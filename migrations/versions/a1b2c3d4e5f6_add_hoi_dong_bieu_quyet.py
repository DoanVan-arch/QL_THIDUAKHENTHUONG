"""add hoi_dong_bieu_quyet table

Revision ID: a1b2c3d4e5f6
Revises: 562238791222
Create Date: 2026-04-28

"""
from alembic import op
import sqlalchemy as sa

revision = 'a1b2c3d4e5f6'
down_revision = '562238791222'
branch_labels = None
depends_on = None


def upgrade():
    op.create_table(
        'hoi_dong_bieu_quyet',
        sa.Column('id', sa.Integer(), nullable=False, autoincrement=True),
        sa.Column('de_xuat_id', sa.Integer(), nullable=False),
        sa.Column('chi_tiet_id', sa.Integer(), nullable=False),
        sa.Column('nguoi_bieu_quyet_id', sa.Integer(), nullable=False),
        sa.Column('vai_tro', sa.String(50), nullable=False),
        sa.Column('ket_qua', sa.String(30), nullable=False),
        sa.Column('ghi_chu', sa.Text(), nullable=True),
        sa.Column('created_at', sa.DateTime(), nullable=True),
        sa.Column('updated_at', sa.DateTime(), nullable=True),
        sa.ForeignKeyConstraint(['de_xuat_id'], ['de_xuat.id']),
        sa.ForeignKeyConstraint(['chi_tiet_id'], ['de_xuat_chi_tiet.id']),
        sa.ForeignKeyConstraint(['nguoi_bieu_quyet_id'], ['users.id']),
        sa.PrimaryKeyConstraint('id'),
        sa.UniqueConstraint('chi_tiet_id', 'vai_tro', name='uq_chitiet_vaitro'),
    )
    op.create_index('ix_hoi_dong_bieu_quyet_de_xuat_id', 'hoi_dong_bieu_quyet', ['de_xuat_id'])
    op.create_index('ix_hoi_dong_bieu_quyet_chi_tiet_id', 'hoi_dong_bieu_quyet', ['chi_tiet_id'])


def downgrade():
    op.drop_index('ix_hoi_dong_bieu_quyet_chi_tiet_id', table_name='hoi_dong_bieu_quyet')
    op.drop_index('ix_hoi_dong_bieu_quyet_de_xuat_id', table_name='hoi_dong_bieu_quyet')
    op.drop_table('hoi_dong_bieu_quyet')
