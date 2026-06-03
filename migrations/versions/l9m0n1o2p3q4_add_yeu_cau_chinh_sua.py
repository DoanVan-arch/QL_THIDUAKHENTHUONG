"""add yeu_cau_chinh_sua table

Revision ID: l9m0n1o2p3q4
Revises: k8l9m0n1o2p3
Create Date: 2026-06-03
"""
from alembic import op
import sqlalchemy as sa

revision = 'l9m0n1o2p3q4'
down_revision = 'k8l9m0n1o2p3'
branch_labels = None
depends_on = None


def upgrade():
    op.create_table(
        'yeu_cau_chinh_sua',
        sa.Column('id', sa.Integer(), autoincrement=True, nullable=False),
        sa.Column('de_xuat_id', sa.Integer(), nullable=False),
        sa.Column('chi_tiet_id', sa.Integer(), nullable=False),
        sa.Column('phong_yeu_cau', sa.String(length=100), nullable=False),
        sa.Column('nguoi_yeu_cau_id', sa.Integer(), nullable=True),
        sa.Column('cac_truong', sa.Text(), nullable=True),
        sa.Column('ly_do', sa.Text(), nullable=True),
        sa.Column('trang_thai', sa.String(length=20), nullable=False, server_default='cho_sua'),
        sa.Column('created_at', sa.DateTime(), nullable=True),
        sa.Column('updated_at', sa.DateTime(), nullable=True),
        sa.Column('ngay_sua', sa.DateTime(), nullable=True),
        sa.ForeignKeyConstraint(['de_xuat_id'], ['de_xuat.id']),
        sa.ForeignKeyConstraint(['chi_tiet_id'], ['de_xuat_chi_tiet.id']),
        sa.ForeignKeyConstraint(['nguoi_yeu_cau_id'], ['users.id']),
        sa.PrimaryKeyConstraint('id'),
    )
    op.create_index('ix_yeu_cau_chinh_sua_de_xuat_id', 'yeu_cau_chinh_sua', ['de_xuat_id'])
    op.create_index('ix_yeu_cau_chinh_sua_chi_tiet_id', 'yeu_cau_chinh_sua', ['chi_tiet_id'])


def downgrade():
    op.drop_index('ix_yeu_cau_chinh_sua_chi_tiet_id', table_name='yeu_cau_chinh_sua')
    op.drop_index('ix_yeu_cau_chinh_sua_de_xuat_id', table_name='yeu_cau_chinh_sua')
    op.drop_table('yeu_cau_chinh_sua')
