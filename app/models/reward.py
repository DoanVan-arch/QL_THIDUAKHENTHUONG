from sqlalchemy import Column, Integer, String, DateTime, ForeignKey, Text, func
from sqlalchemy.orm import relationship
from app.extensions import db


class KhenThuong(db.Model):
    """Record of awarded individuals after final approval."""
    __tablename__ = 'khen_thuong'

    id = Column(Integer, primary_key=True, autoincrement=True)
    de_xuat_id = Column(Integer, ForeignKey('de_xuat.id'), nullable=False, index=True)
    chi_tiet_id = Column(Integer, ForeignKey('de_xuat_chi_tiet.id'), nullable=False, index=True)
    quan_nhan_id = Column(Integer, ForeignKey('quan_nhan.id'), nullable=True, index=True)
    don_vi_id = Column(Integer, ForeignKey('don_vi.id'), nullable=False, index=True)

    ho_ten = Column(String(100), nullable=False)
    cap_bac = Column(String(50), nullable=True)
    chuc_vu = Column(String(100), nullable=True)
    doi_tuong = Column(String(50), nullable=True)
    loai_danh_hieu = Column(String(50), nullable=False)
    nam_hoc = Column(String(20), nullable=False)

    nguoi_duyet_id = Column(Integer, ForeignKey('users.id'), nullable=True)
    ngay_duyet = Column(DateTime, nullable=True)
    ghi_chu = Column(Text, nullable=True)

    created_at = Column(DateTime, default=func.now())

    de_xuat = relationship('DeXuat')
    chi_tiet = relationship('DeXuatChiTiet')
    quan_nhan = relationship('QuanNhan')
    don_vi = relationship('DonVi')
    nguoi_duyet = relationship('User')
