"""add more performance indexes (phong_duyet, don_vi thu_tu, bi_loai/phong_loai/trang_thai, users don_vi_id)

Revision ID: 0874451bc2b5
Revises: b2c3d4e5f6a8
Create Date: 2026-07-01 00:00:00.000000

"""
from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = '0874451bc2b5'
down_revision = 'b2c3d4e5f6a8'
branch_labels = None
depends_on = None


def upgrade():
    # ── PheDuyet.phong_duyet ────────────────────────────────────────────────
    # Trước đây phong_duyet chỉ được cover bởi UNIQUE(de_xuat_id, phong_duyet)
    # (de_xuat_id là cột dẫn đầu) — không giúp ích cho các query rất phổ biến
    # kiểu `PheDuyet.query.filter_by(phong_duyet=phong_name)` (lọc RIÊNG theo
    # phong_duyet, không kèm de_xuat_id) dùng trong pending_list()/export_word().
    with op.batch_alter_table('phe_duyet', schema=None) as batch_op:
        batch_op.create_index(
            'ix_phe_duyet_phong_duyet_ket_qua', ['phong_duyet', 'ket_qua'], unique=False
        )

    # ── DonVi (is_active, thu_tu) ────────────────────────────────────────────
    # thu_tu dùng để ORDER BY ở hầu hết các trang danh sách/tracking, thường kèm
    # điều kiện is_active = true.
    with op.batch_alter_table('don_vi', schema=None) as batch_op:
        batch_op.create_index(
            'ix_don_vi_active_thu_tu', ['is_active', 'thu_tu'], unique=False
        )

    # ── DeXuatChiTiet (bi_loai, phong_loai, trang_thai) ──────────────────────
    # Hỗ trợ trực tiếp điều kiện lọc "ẩn khi bị từ chối dứt điểm bởi Tuyên huấn":
    # bi_loai=1 AND phong_loai='Tuyên huấn' AND trang_thai='tu_choi'.
    with op.batch_alter_table('de_xuat_chi_tiet', schema=None) as batch_op:
        batch_op.create_index(
            'ix_dxct_bi_loai_phong_loai_trang_thai',
            ['bi_loai', 'phong_loai', 'trang_thai'], unique=False
        )

    # ── users.don_vi_id ───────────────────────────────────────────────────────
    # FK chưa có index — dùng để tìm user quản lý 1 đơn vị (User.query.filter_by(don_vi_id=...)).
    with op.batch_alter_table('users', schema=None) as batch_op:
        batch_op.create_index('ix_users_don_vi_id', ['don_vi_id'], unique=False)


def downgrade():
    with op.batch_alter_table('users', schema=None) as batch_op:
        batch_op.drop_index('ix_users_don_vi_id')

    with op.batch_alter_table('de_xuat_chi_tiet', schema=None) as batch_op:
        batch_op.drop_index('ix_dxct_bi_loai_phong_loai_trang_thai')

    with op.batch_alter_table('don_vi', schema=None) as batch_op:
        batch_op.drop_index('ix_don_vi_active_thu_tu')

    with op.batch_alter_table('phe_duyet', schema=None) as batch_op:
        batch_op.drop_index('ix_phe_duyet_phong_duyet_ket_qua')
