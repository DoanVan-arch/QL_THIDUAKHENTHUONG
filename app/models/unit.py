import enum
from sqlalchemy import Column, Integer, String, Enum, Boolean
from sqlalchemy.orm import relationship
from app.extensions import db


class LoaiDonVi(enum.Enum):
    KHOI_COQUAN = 'Khối cơ quan'
    KHOA = 'Các khoa'
    HE = 'Hệ'
    TIEU_DOAN = 'Tiểu đoàn'


class DonVi(db.Model):
    __tablename__ = 'don_vi'

    id = Column(Integer, primary_key=True, autoincrement=True)
    ma_don_vi = Column(String(20), unique=True, nullable=False)
    ten_don_vi = Column(String(200), nullable=False)
    loai_don_vi = Column(Enum(LoaiDonVi), nullable=False)
    thu_tu = Column(Integer, default=0)
    is_active = Column(Boolean, default=True)

    user_account = relationship('User', back_populates='don_vi', uselist=False)
    quan_nhans = relationship('QuanNhan', back_populates='don_vi', lazy='dynamic')
    de_xuats = relationship('DeXuat', back_populates='don_vi', lazy='dynamic')
