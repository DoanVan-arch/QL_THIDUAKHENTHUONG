from sqlalchemy import Column, Integer, String, Boolean, DateTime, ForeignKey, Text, func
from sqlalchemy.orm import relationship
from app.extensions import db


class ThongBao(db.Model):
    """Notification for unit accounts when individuals are rejected."""
    __tablename__ = 'thong_bao'

    id = Column(Integer, primary_key=True, autoincrement=True)
    user_id = Column(Integer, ForeignKey('users.id'), nullable=False, index=True)
    de_xuat_id = Column(Integer, ForeignKey('de_xuat.id'), nullable=True, index=True)
    chi_tiet_id = Column(Integer, ForeignKey('de_xuat_chi_tiet.id'), nullable=True)
    loai = Column(String(50), nullable=False, default='tu_choi')  # tu_choi, thong_tin
    tieu_de = Column(String(255), nullable=False)
    noi_dung = Column(Text, nullable=True)
    da_doc = Column(Boolean, default=False)
    created_at = Column(DateTime, default=func.now())

    user = relationship('User')
    de_xuat = relationship('DeXuat')
    chi_tiet = relationship('DeXuatChiTiet')
