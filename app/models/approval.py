import enum
from sqlalchemy import Column, Integer, String, DateTime, ForeignKey, Text, func, UniqueConstraint
from sqlalchemy.orm import relationship
from app.extensions import db


class PhongDuyet(enum.Enum):
    PHONG_CHINHTRI = 'Phòng Chính trị'
    PHONG_THAMMUU = 'Phòng Tham mưu'
    PHONG_KHOAHOC = 'Phòng Khoa học'
    PHONG_DAOTAO = 'Phòng Đào tạo'
    THU_TRUONG_PHONG_CHINHTRI = 'Thủ trưởng Phòng Chính trị'
    THU_TRUONG_PHONG_TMHC = 'Thủ trưởng Phòng TM-HC'
    BAN_CANBO = 'Ban Cán bộ'
    BAN_TOCHUC = 'Ban Tổ chức'
    BAN_TUYENHUAN = 'Ban Tuyên huấn'
    BAN_CTCQ = 'Ban Công tác quần chúng'
    BAN_CNTT = 'Ban Công nghệ thông tin'
    BAN_TAC_HUAN = 'Ban Tác huấn'
    BAN_KHAOTHI = 'Ban Khảo thí'
    UY_BAN_KIEMTRA = 'Ủy ban Kiểm tra'
    BAN_QUANLUC = 'Ban Quân lực'
    ADMIN_TUYENHUAN = 'Tuyên huấn'


class KetQuaDuyet(enum.Enum):
    CHO_DUYET = 'Chờ duyệt'
    DONG_Y = 'Đồng ý'
    TU_CHOI = 'Từ chối'


class PheDuyet(db.Model):
    __tablename__ = 'phe_duyet'

    id = Column(Integer, primary_key=True, autoincrement=True)
    de_xuat_id = Column(Integer, ForeignKey('de_xuat.id'), nullable=False, index=True)
    phong_duyet = Column(String(50), nullable=False)
    ket_qua = Column(String(30), default=KetQuaDuyet.CHO_DUYET.value)
    nguoi_duyet_id = Column(Integer, ForeignKey('users.id'), nullable=True)
    ngay_duyet = Column(DateTime, nullable=True)
    ly_do = Column(Text, nullable=True)
    ghi_chu = Column(Text, nullable=True)
    created_at = Column(DateTime, default=func.now())
    updated_at = Column(DateTime, default=func.now(), onupdate=func.now())

    __table_args__ = (
        UniqueConstraint('de_xuat_id', 'phong_duyet', name='uq_de_xuat_phong'),
    )

    de_xuat = relationship('DeXuat', back_populates='phe_duyets')
    nguoi_duyet = relationship('User')
    chi_tiet_duyet = relationship('KetQuaDuyetChiTiet', back_populates='phe_duyet', cascade='all, delete-orphan')


class KetQuaDuyetChiTiet(db.Model):
    """Per-item approval results within a department's review."""
    __tablename__ = 'ket_qua_duyet_chi_tiet'

    id = Column(Integer, primary_key=True, autoincrement=True)
    phe_duyet_id = Column(Integer, ForeignKey('phe_duyet.id'), nullable=False, index=True)
    chi_tiet_id = Column(Integer, ForeignKey('de_xuat_chi_tiet.id'), nullable=False, index=True)
    ket_qua = Column(String(30), default=KetQuaDuyet.CHO_DUYET.value)
    ly_do = Column(Text, nullable=True)
    created_at = Column(DateTime, default=func.now())
    updated_at = Column(DateTime, default=func.now(), onupdate=func.now())

    __table_args__ = (
        UniqueConstraint('phe_duyet_id', 'chi_tiet_id', name='uq_pheduyet_chitiet'),
    )

    phe_duyet = relationship('PheDuyet', back_populates='chi_tiet_duyet')
    chi_tiet = relationship('DeXuatChiTiet')
