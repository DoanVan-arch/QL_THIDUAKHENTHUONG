"""extend role enum for thu truong and viewers

Revision ID: e7a9c40f2f8f
Revises: d1f35b5ad8ce
Create Date: 2026-04-14 00:00:00.000000

"""
from alembic import op


revision = 'e7a9c40f2f8f'
down_revision = 'd1f35b5ad8ce'
branch_labels = None
depends_on = None


def upgrade():
    op.execute(
        "ALTER TABLE users MODIFY COLUMN role "
        "ENUM('unit_user','phong_chinhTri','phong_thamMuu','phong_khoaHoc','phong_daoTao',"
        "'thu_truongPhongChinhTri','thu_truongPhongTMHC',"
        "'ban_canBo','ban_toChuc','ban_tuyenHuan','ban_congTacQuanChung','ban_congNgheThongTin','ban_tacHuan',"
        "'ban_baoVeAnNinh','ban_keHoachTongHop','uy_banKiemTra','ban_quanLuc','admin') NOT NULL"
    )


def downgrade():
    op.execute(
        "ALTER TABLE users MODIFY COLUMN role "
        "ENUM('unit_user','phong_chinhTri','phong_thamMuu','phong_khoaHoc','phong_daoTao',"
        "'ban_canBo','ban_toChuc','ban_tuyenHuan','ban_congTacQuanChung','ban_congNgheThongTin','ban_tacHuan',"
        "'ban_quanLuc','admin') NOT NULL"
    )
