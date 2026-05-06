"""add_chuyen_don_vi_table

Revision ID: c4d5e6f7a8b9
Revises: b2c3d4e5f6a1
Create Date: 2026-05-06
"""
from alembic import op
import sqlalchemy as sa

revision = 'c4d5e6f7a8b9'
down_revision = 'b2c3d4e5f6a1'
branch_labels = None
depends_on = None


def upgrade():
    op.create_table(
        'chuyen_don_vi',
        sa.Column('id',                sa.Integer(),     nullable=False, autoincrement=True),
        sa.Column('quan_nhan_id',      sa.Integer(),     nullable=False),
        sa.Column('don_vi_nguon_id',   sa.Integer(),     nullable=False),
        sa.Column('don_vi_dich_id',    sa.Integer(),     nullable=False),
        sa.Column('nguoi_tao_id',      sa.Integer(),     nullable=True),
        sa.Column('nguoi_xac_nhan_id', sa.Integer(),     nullable=True),
        sa.Column('trang_thai',        sa.String(20),    nullable=False, server_default='PENDING'),
        sa.Column('ly_do',             sa.Text(),        nullable=True),
        sa.Column('ghi_chu',           sa.Text(),        nullable=True),
        sa.Column('ngay_tao',          sa.DateTime(),    nullable=False, server_default=sa.text('NOW()')),
        sa.Column('ngay_xu_ly',        sa.DateTime(),    nullable=True),
        sa.PrimaryKeyConstraint('id'),
        sa.ForeignKeyConstraint(['quan_nhan_id'],      ['quan_nhan.id']),
        sa.ForeignKeyConstraint(['don_vi_nguon_id'],   ['don_vi.id']),
        sa.ForeignKeyConstraint(['don_vi_dich_id'],    ['don_vi.id']),
        sa.ForeignKeyConstraint(['nguoi_tao_id'],      ['users.id']),
        sa.ForeignKeyConstraint(['nguoi_xac_nhan_id'], ['users.id']),
    )
    op.create_index('ix_chuyen_don_vi_quan_nhan_id',    'chuyen_don_vi', ['quan_nhan_id'])
    op.create_index('ix_chuyen_don_vi_don_vi_dich_id',  'chuyen_don_vi', ['don_vi_dich_id'])
    op.create_index('ix_chuyen_don_vi_trang_thai',      'chuyen_don_vi', ['trang_thai'])


def downgrade():
    op.drop_table('chuyen_don_vi')
