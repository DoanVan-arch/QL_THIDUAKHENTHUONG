import enum
from sqlalchemy import Column, Integer, String, Enum, Date, DateTime, ForeignKey, Text, func
from sqlalchemy.orm import relationship
from app.extensions import db


class LoaiChungChi(enum.Enum):
    BANG_KHEN = 'Bằng khen'
    GIAY_KHEN = 'Giấy khen'
    THANH_TICH_KHAC = 'Thành tích khác'


class ChungChi(db.Model):
    __tablename__ = 'chung_chi'

    id = Column(Integer, primary_key=True, autoincrement=True)
    quan_nhan_id = Column(Integer, ForeignKey('quan_nhan.id'), nullable=False, index=True)
    loai = Column(String(50), nullable=False)
    ten_chung_chi = Column(String(255), nullable=False)
    so_hieu = Column(String(100), nullable=True)
    ngay_cap = Column(Date, nullable=True)
    co_quan_cap = Column(String(255), nullable=True)
    duong_dan_anh = Column(String(255), nullable=True)
    ghi_chu = Column(Text, nullable=True)
    created_at = Column(DateTime, default=func.now())

    quan_nhan = relationship('QuanNhan', back_populates='chung_chis')
