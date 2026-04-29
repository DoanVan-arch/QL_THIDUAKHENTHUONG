from sqlalchemy import Column, Integer, String, DateTime, ForeignKey, Text, func, UniqueConstraint
from sqlalchemy.orm import relationship
from app.extensions import db


# 6 Hội đồng roles allowed to cast votes
HOI_DONG_VAI_TRO = [
    'ban_tuyenHuan',
    'ban_canBo',
    'ban_congTacQuanChung',
    'ban_baoVeAnNinh',
    'ban_toChuc',
    'ban_keHoachTongHop',
]

HOI_DONG_VAI_TRO_DISPLAY = {
    'ban_tuyenHuan': 'Ban Tuyên huấn',
    'ban_canBo': 'Ban Cán bộ',
    'ban_congTacQuanChung': 'Ban Công tác quần chúng',
    'ban_baoVeAnNinh': 'Ban Bảo vệ an ninh',
    'ban_toChuc': 'Ban Tổ chức',
    'ban_keHoachTongHop': 'Ban Kế hoạch tổng hợp',
}

KET_QUA_DONG_Y = 'Đồng ý'
KET_QUA_KHONG_DONG_Y = 'Không đồng ý'


class HoiDongBieuQuyet(db.Model):
    """Vote cast by a Hội đồng member for one individual nomination item."""
    __tablename__ = 'hoi_dong_bieu_quyet'

    id = Column(Integer, primary_key=True, autoincrement=True)
    de_xuat_id = Column(Integer, ForeignKey('de_xuat.id'), nullable=False, index=True)
    chi_tiet_id = Column(Integer, ForeignKey('de_xuat_chi_tiet.id'), nullable=False, index=True)
    nguoi_bieu_quyet_id = Column(Integer, ForeignKey('users.id'), nullable=False)
    vai_tro = Column(String(50), nullable=False)          # one of HOI_DONG_VAI_TRO
    ket_qua = Column(String(30), nullable=False)          # KET_QUA_DONG_Y or KET_QUA_KHONG_DONG_Y
    ghi_chu = Column(Text, nullable=True)
    created_at = Column(DateTime, default=func.now())
    updated_at = Column(DateTime, default=func.now(), onupdate=func.now())

    __table_args__ = (
        UniqueConstraint('chi_tiet_id', 'vai_tro', name='uq_chitiet_vaitro'),
    )

    de_xuat = relationship('DeXuat')
    chi_tiet = relationship('DeXuatChiTiet')
    nguoi_bieu_quyet = relationship('User')
