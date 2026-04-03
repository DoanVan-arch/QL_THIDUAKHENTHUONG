import enum
from sqlalchemy import Column, Integer, String, Enum, Boolean, Date, DateTime, ForeignKey, Text, func
from sqlalchemy.orm import relationship
from app.extensions import db


class CapBac(enum.Enum):
    BINH_NHI = 'Binh nhì'
    BINH_NHAT = 'Binh nhất'
    HA_SI = 'Hạ sĩ'
    TRUNG_SI = 'Trung sĩ'
    THUONG_SI = 'Thượng sĩ'
    THIEU_UY = 'Thiếu úy'
    TRUNG_UY = 'Trung úy'
    THUONG_UY = 'Thượng úy'
    DAI_UY = 'Đại úy'
    THIEU_TA = 'Thiếu tá'
    TRUNG_TA = 'Trung tá'
    THUONG_TA = 'Thượng tá'
    DAI_TA = 'Đại tá'


class HocHam(enum.Enum):
    NONE = 'Không'
    PGS = 'Phó Giáo sư'
    GS = 'Giáo sư'


class HocVi(enum.Enum):
    NONE = 'Không'
    THAC_SI = 'Thạc sĩ'
    TIEN_SI = 'Tiến sĩ'


class DoiTuong(enum.Enum):
    GV = 'Giảng viên'
    CB = 'Cán bộ'
    CCQP = 'Công chức quốc phòng'
    QNCN = 'Quân nhân chuyên nghiệp'
    CNV = 'Công nhân viên'
    SV_NAM1 = 'Học viên năm thứ I'
    SV_NAM2 = 'Học viên năm thứ II'
    SV_NAM3 = 'Học viên năm thứ III'
    SV_NAM4 = 'Học viên năm thứ IV'
    SAU_DAI_HOC = 'Học viên sau đại học'
    VB2 = 'Học viên VB2'
    HVTS = 'Học viên tiến sĩ'
    HVQT = 'Học viên quốc tế'


class MucDoHoanThanh(enum.Enum):
    HTXSNV = 'Hoàn thành xuất sắc nhiệm vụ'
    HTTNV = 'Hoàn thành tốt nhiệm vụ'
    HTNV = 'Hoàn thành nhiệm vụ'
    KHTNV = 'Không hoàn thành nhiệm vụ'


DOI_TUONG_IS_STUDENT = {
    DoiTuong.SV_NAM1, DoiTuong.SV_NAM2, DoiTuong.SV_NAM3,
    DoiTuong.SV_NAM4, DoiTuong.SAU_DAI_HOC, DoiTuong.VB2,
    DoiTuong.HVTS, DoiTuong.HVQT,
}

DOI_TUONG_IS_LECTURER = {DoiTuong.GV}

DOI_TUONG_IS_CADRE = {DoiTuong.CB, DoiTuong.CCQP, DoiTuong.QNCN, DoiTuong.CNV}


class QuanNhan(db.Model):
    __tablename__ = 'quan_nhan'

    id = Column(Integer, primary_key=True, autoincrement=True)
    don_vi_id = Column(Integer, ForeignKey('don_vi.id'), nullable=False, index=True)
    ho_ten = Column(String(100), nullable=False)
    cap_bac = Column(String(50), nullable=True)
    chuc_danh = Column(String(100), nullable=True)
    chuc_vu = Column(String(100), nullable=True)
    ngay_sinh = Column(Date, nullable=True)
    ngay_nhap_ngu = Column(String(20), nullable=True)
    doi_tuong = Column(String(50), nullable=True)
    hoc_ham = Column(String(50), default='Không')
    hoc_vi = Column(String(50), default='Không')
    trinh_do_hoc_van = Column(String(50), nullable=True)
    ngoai_ngu = Column(String(100), nullable=True)
    la_chi_huy = Column(Boolean, default=False)
    la_bi_thu = Column(Boolean, default=False)
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=func.now())
    updated_at = Column(DateTime, default=func.now(), onupdate=func.now())

    don_vi = relationship('DonVi', back_populates='quan_nhans')
    chung_chis = relationship('ChungChi', back_populates='quan_nhan', cascade='all, delete-orphan')
    de_xuat_items = relationship('DeXuatChiTiet', back_populates='quan_nhan')

    @property
    def doi_tuong_enum(self):
        try:
            return DoiTuong(self.doi_tuong)
        except (ValueError, KeyError):
            return None

    @property
    def is_student(self):
        dt = self.doi_tuong_enum
        return dt in DOI_TUONG_IS_STUDENT if dt else False

    @property
    def is_lecturer(self):
        dt = self.doi_tuong_enum
        return dt in DOI_TUONG_IS_LECTURER if dt else False

    @property
    def has_tien_si(self):
        return self.hoc_vi == HocVi.TIEN_SI.value
