"""add ban_khaothi role

Revision ID: f4e2dcb8c3aa
Revises: e7a9c40f2f8f
Create Date: 2026-04-14 00:00:00.000000

"""
from alembic import op


revision = 'f4e2dcb8c3aa'
down_revision = 'e7a9c40f2f8f'
branch_labels = None
depends_on = None


def upgrade():
    op.execute(
        "ALTER TABLE users MODIFY COLUMN role "
        "ENUM('unit_user','phong_chinhTri','phong_thamMuu','phong_khoaHoc','phong_daoTao',"
        "'thu_truongPhongChinhTri','thu_truongPhongTMHC',"
        "'ban_canBo','ban_toChuc','ban_tuyenHuan','ban_congTacQuanChung','ban_congNgheThongTin','ban_tacHuan','ban_khaoThi',"
        "'ban_baoVeAnNinh','ban_keHoachTongHop','uy_banKiemTra','ban_quanLuc','admin') NOT NULL"
    )


def downgrade():
    op.execute(
        "ALTER TABLE users MODIFY COLUMN role "
        "ENUM('unit_user','phong_chinhTri','phong_thamMuu','phong_khoaHoc','phong_daoTao',"
        "'thu_truongPhongChinhTri','thu_truongPhongTMHC',"
        "'ban_canBo','ban_toChuc','ban_tuyenHuan','ban_congTacQuanChung','ban_congNgheThongTin','ban_tacHuan',"
        "'ban_baoVeAnNinh','ban_keHoachTongHop','uy_banKiemTra','ban_quanLuc','admin') NOT NULL"
    )
