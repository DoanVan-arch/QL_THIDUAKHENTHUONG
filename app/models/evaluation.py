import json
from sqlalchemy import Column, Integer, String, Boolean, DateTime, ForeignKey, Text, UniqueConstraint, func
from sqlalchemy.orm import relationship
from app.extensions import db


class NhomTieuChi(db.Model):
    __tablename__ = 'nhom_tieu_chi'

    id = Column(Integer, primary_key=True, autoincrement=True)
    ma_nhom = Column(String(50), nullable=False, unique=True)
    ten_nhom = Column(String(150), nullable=False)
    mo_ta = Column(Text, nullable=True)
    _doi_tuong_ap_dung = Column('doi_tuong_ap_dung', Text, nullable=True)
    thu_tu = Column(Integer, default=0)
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=func.now())
    updated_at = Column(DateTime, default=func.now(), onupdate=func.now())

    @property
    def doi_tuong_ap_dung(self):
        if self._doi_tuong_ap_dung:
            try:
                return json.loads(self._doi_tuong_ap_dung)
            except (json.JSONDecodeError, TypeError):
                return []
        return []

    @doi_tuong_ap_dung.setter
    def doi_tuong_ap_dung(self, value):
        if isinstance(value, list):
            self._doi_tuong_ap_dung = json.dumps(value, ensure_ascii=False)
        else:
            self._doi_tuong_ap_dung = value


class DanhGiaHangNam(db.Model):
    __tablename__ = 'danh_gia_hang_nam'
    __table_args__ = (
        UniqueConstraint('quan_nhan_id', 'nam_hoc', name='uq_dghn_quan_nhan_nam_hoc'),
    )

    XEP_LOAI_DANG_VIEN_CHOICES = [
        'Hoàn thành xuất sắc nhiệm vụ',
        'Hoàn thành tốt nhiệm vụ',
        'Hoàn thành nhiệm vụ',
        'Không hoàn thành nhiệm vụ',
    ]

    XEP_LOAI_CAN_BO_CHOICES = [
        'Hoàn thành xuất sắc nhiệm vụ',
        'Hoàn thành tốt nhiệm vụ',
        'Hoàn thành nhiệm vụ',
        'Không hoàn thành nhiệm vụ',
    ]

    id = Column(Integer, primary_key=True, autoincrement=True)
    quan_nhan_id = Column(Integer, ForeignKey('quan_nhan.id'), nullable=False, index=True)
    don_vi_id = Column(Integer, ForeignKey('don_vi.id'), nullable=False, index=True)
    nam_hoc = Column(String(20), nullable=False, index=True)
    xep_loai_dang_vien = Column(String(100), nullable=False)
    xep_loai_can_bo = Column(String(100), nullable=False)
    nguoi_cap_nhat_id = Column(Integer, ForeignKey('users.id'), nullable=True)
    created_at = Column(DateTime, default=func.now())
    updated_at = Column(DateTime, default=func.now(), onupdate=func.now())

    quan_nhan = relationship('QuanNhan')
    don_vi = relationship('DonVi')
    nguoi_cap_nhat = relationship('User')
