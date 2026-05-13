"""add_new_roles_haucankythuat_saudaihoc

Revision ID: f1a2b3c4d5e6
Revises: e1f2a3b4c5d6
Create Date: 2026-05-13

"""
from alembic import op

revision = 'f1a2b3c4d5e6'
down_revision = 'e1f2a3b4c5d6'
branch_labels = None
depends_on = None


def upgrade():
    op.execute("""
        ALTER TABLE users MODIFY COLUMN `role` ENUM(
            'unit_user',
            'phong_chinhTri',
            'phong_thamMuu',
            'phong_khoaHoc',
            'phong_daoTao',
            'thu_truongPhongChinhTri',
            'thu_truongPhongTMHC',
            'ban_canBo',
            'ban_toChuc',
            'ban_tuyenHuan',
            'ban_congTacQuanChung',
            'ban_congNgheThongTin',
            'ban_tacHuan',
            'ban_khaoThi',
            'ban_baoVeAnNinh',
            'ban_keHoachTongHop',
            'uy_banKiemTra',
            'ban_quanLuc',
            'phong_hauCanKyThuat',
            'ban_sauDaiHoc',
            'admin'
        ) NOT NULL
    """)


def downgrade():
    op.execute("""
        ALTER TABLE users MODIFY COLUMN `role` ENUM(
            'unit_user',
            'phong_chinhTri',
            'phong_thamMuu',
            'phong_khoaHoc',
            'phong_daoTao',
            'thu_truongPhongChinhTri',
            'thu_truongPhongTMHC',
            'ban_canBo',
            'ban_toChuc',
            'ban_tuyenHuan',
            'ban_congTacQuanChung',
            'ban_congNgheThongTin',
            'ban_tacHuan',
            'ban_khaoThi',
            'ban_baoVeAnNinh',
            'ban_keHoachTongHop',
            'uy_banKiemTra',
            'ban_quanLuc',
            'admin'
        ) NOT NULL
    """)
