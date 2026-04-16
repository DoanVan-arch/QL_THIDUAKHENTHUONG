"""add nhom tieu chi and danh gia hang nam tables

Revision ID: 5f4a8c2b9d11
Revises: e2c7a6b1d4f9
Create Date: 2026-04-16 10:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


revision = '5f4a8c2b9d11'
down_revision = 'e2c7a6b1d4f9'
branch_labels = None
depends_on = None


def upgrade():
    op.create_table(
        'nhom_tieu_chi',
        sa.Column('id', sa.Integer(), autoincrement=True, nullable=False),
        sa.Column('ma_nhom', sa.String(length=50), nullable=False),
        sa.Column('ten_nhom', sa.String(length=150), nullable=False),
        sa.Column('mo_ta', sa.Text(), nullable=True),
        sa.Column('doi_tuong_ap_dung', sa.Text(), nullable=True),
        sa.Column('thu_tu', sa.Integer(), nullable=True),
        sa.Column('is_active', sa.Boolean(), nullable=True),
        sa.Column('created_at', sa.DateTime(), nullable=True),
        sa.Column('updated_at', sa.DateTime(), nullable=True),
        sa.PrimaryKeyConstraint('id'),
        sa.UniqueConstraint('ma_nhom'),
    )

    op.create_table(
        'danh_gia_hang_nam',
        sa.Column('id', sa.Integer(), autoincrement=True, nullable=False),
        sa.Column('quan_nhan_id', sa.Integer(), nullable=False),
        sa.Column('don_vi_id', sa.Integer(), nullable=False),
        sa.Column('nam_hoc', sa.String(length=20), nullable=False),
        sa.Column('xep_loai_dang_vien', sa.String(length=100), nullable=False),
        sa.Column('xep_loai_can_bo', sa.String(length=100), nullable=False),
        sa.Column('nguoi_cap_nhat_id', sa.Integer(), nullable=True),
        sa.Column('created_at', sa.DateTime(), nullable=True),
        sa.Column('updated_at', sa.DateTime(), nullable=True),
        sa.ForeignKeyConstraint(['don_vi_id'], ['don_vi.id']),
        sa.ForeignKeyConstraint(['nguoi_cap_nhat_id'], ['users.id']),
        sa.ForeignKeyConstraint(['quan_nhan_id'], ['quan_nhan.id']),
        sa.PrimaryKeyConstraint('id'),
        sa.UniqueConstraint('quan_nhan_id', 'nam_hoc', name='uq_dghn_quan_nhan_nam_hoc'),
    )

    op.create_index(op.f('ix_danh_gia_hang_nam_don_vi_id'), 'danh_gia_hang_nam', ['don_vi_id'], unique=False)
    op.create_index(op.f('ix_danh_gia_hang_nam_nam_hoc'), 'danh_gia_hang_nam', ['nam_hoc'], unique=False)
    op.create_index(op.f('ix_danh_gia_hang_nam_quan_nhan_id'), 'danh_gia_hang_nam', ['quan_nhan_id'], unique=False)


def downgrade():
    op.drop_index(op.f('ix_danh_gia_hang_nam_quan_nhan_id'), table_name='danh_gia_hang_nam')
    op.drop_index(op.f('ix_danh_gia_hang_nam_nam_hoc'), table_name='danh_gia_hang_nam')
    op.drop_index(op.f('ix_danh_gia_hang_nam_don_vi_id'), table_name='danh_gia_hang_nam')
    op.drop_table('danh_gia_hang_nam')
    op.drop_table('nhom_tieu_chi')
