import enum
import json
from sqlalchemy import Column, Integer, String, Enum, Boolean, Float, Date, DateTime, ForeignKey, Text, func
from sqlalchemy.orm import relationship
from app.extensions import db


class TieuChi(db.Model):
    """Criteria definition with metadata, tooltip, evidence flag, and department assignments."""
    __tablename__ = 'tieu_chi'

    id = Column(Integer, primary_key=True, autoincrement=True)
    ma_truong = Column(String(50), nullable=False, unique=True)   # matches DeXuatChiTiet column name
    ten = Column(String(150), nullable=False)                      # display label
    huong_dan = Column(Text, nullable=True)                        # tooltip / guidance text
    nhom = Column(String(50), nullable=False, default='chung')     # category: chung, giang_vien, hoc_vien, nckh, khac
    co_minh_chung = Column(Boolean, default=False)                 # requires evidence upload?
    _phong_duyet = Column('phong_duyet', Text, nullable=True)      # JSON list of department role strings
    thu_tu = Column(Integer, default=0)
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=func.now())
    updated_at = Column(DateTime, default=func.now(), onupdate=func.now())

    @property
    def phong_duyet(self):
        if self._phong_duyet:
            try:
                return json.loads(self._phong_duyet)
            except (json.JSONDecodeError, TypeError):
                return []
        return []

    @phong_duyet.setter
    def phong_duyet(self, value):
        if isinstance(value, list):
            self._phong_duyet = json.dumps(value, ensure_ascii=False)
        else:
            self._phong_duyet = value

    # Convenience: map of nhom display names
    NHOM_CHOICES = {
        'chung': 'Tiêu chí chung',
        'giang_vien': 'Tiêu chí giảng viên',
        'hoc_vien': 'Tiêu chí học viên',
        'nckh': 'Tiêu chí NCKH',
        'khac': 'Khác',
    }


class LoaiDanhHieu(enum.Enum):
    CHIEN_SI_THI_DUA = 'Chiến sĩ thi đua'
    CHIEN_SI_TIEN_TIEN = 'Chiến sĩ tiên tiến'
    DON_VI_QUYET_THANG = 'Đơn vị quyết thắng'
    DON_VI_TIEN_TIEN = 'Đơn vị tiên tiến'


class DanhHieu(db.Model):
    """Award title with associated criteria fields (stored as JSON list of field names)."""
    __tablename__ = 'danh_hieu'

    id = Column(Integer, primary_key=True, autoincrement=True)
    ten_danh_hieu = Column(String(100), nullable=False, unique=True)
    ma_danh_hieu = Column(String(20), nullable=False, unique=True)
    pham_vi = Column(String(20), nullable=False, default='Cá nhân')  # 'Cá nhân' or 'Đơn vị'
    _tieu_chi = Column('tieu_chi', Text, nullable=True)  # JSON list of field names
    thu_tu = Column(Integer, default=0)
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=func.now())
    updated_at = Column(DateTime, default=func.now(), onupdate=func.now())

    @property
    def tieu_chi(self):
        if self._tieu_chi:
            try:
                return json.loads(self._tieu_chi)
            except (json.JSONDecodeError, TypeError):
                return []
        return []

    @tieu_chi.setter
    def tieu_chi(self, value):
        if isinstance(value, list):
            self._tieu_chi = json.dumps(value, ensure_ascii=False)
        else:
            self._tieu_chi = value


class TrangThaiDeXuat(enum.Enum):
    NHAP = 'Nháp'
    CHO_DUYET = 'Chờ duyệt'
    DANG_DUYET = 'Đang duyệt'
    DA_DUYET = 'Đã duyệt'
    PHE_DUYET_CUOI = 'Phê duyệt cuối'
    TU_CHOI = 'Từ chối'


class DeXuat(db.Model):
    __tablename__ = 'de_xuat'

    id = Column(Integer, primary_key=True, autoincrement=True)
    don_vi_id = Column(Integer, ForeignKey('don_vi.id'), nullable=False, index=True)
    nam_hoc = Column(String(20), nullable=False)
    trang_thai = Column(String(30), default=TrangThaiDeXuat.NHAP.value)
    ngay_tao = Column(DateTime, default=func.now())
    ngay_gui = Column(DateTime, nullable=True)
    nguoi_tao_id = Column(Integer, ForeignKey('users.id'), nullable=False)
    ghi_chu = Column(Text, nullable=True)
    updated_at = Column(DateTime, default=func.now(), onupdate=func.now())

    don_vi = relationship('DonVi', back_populates='de_xuats')
    nguoi_tao = relationship('User')
    chi_tiets = relationship('DeXuatChiTiet', back_populates='de_xuat', cascade='all, delete-orphan')
    phe_duyets = relationship('PheDuyet', back_populates='de_xuat', cascade='all, delete-orphan')

    @property
    def trang_thai_enum(self):
        try:
            return TrangThaiDeXuat(self.trang_thai)
        except (ValueError, KeyError):
            return None

    @property
    def is_editable(self):
        return self.trang_thai in (TrangThaiDeXuat.NHAP.value, TrangThaiDeXuat.TU_CHOI.value)

    @property
    def approval_progress(self):
        total = len([p for p in self.phe_duyets if p.phong_duyet != 'Tuyên huấn'])
        approved = len([p for p in self.phe_duyets if p.phong_duyet != 'Tuyên huấn' and p.ket_qua == 'Đồng ý'])
        return approved, total


class DeXuatChiTiet(db.Model):
    __tablename__ = 'de_xuat_chi_tiet'

    id = Column(Integer, primary_key=True, autoincrement=True)
    de_xuat_id = Column(Integer, ForeignKey('de_xuat.id'), nullable=False, index=True)
    quan_nhan_id = Column(Integer, ForeignKey('quan_nhan.id'), nullable=True, index=True)

    loai_danh_hieu = Column(String(50), nullable=False)
    doi_tuong = Column(String(50), nullable=True)
    nam_hoc = Column(String(20), nullable=True)

    # Common fields
    muc_do_hoan_thanh = Column(String(100), nullable=True)
    kiem_tra_tin_hoc = Column(String(50), nullable=True)
    diem_kiem_tra_tin_hoc = Column(String(20), nullable=True)
    kiem_tra_dieu_lenh = Column(String(50), nullable=True)
    diem_kiem_tra_dieu_lenh = Column(String(20), nullable=True)
    dia_ly_quan_su = Column(String(50), nullable=True)
    diem_dia_ly_quan_su = Column(String(20), nullable=True)
    ban_sung = Column(String(50), nullable=True)
    diem_ban_sung = Column(String(20), nullable=True)
    the_luc = Column(String(50), nullable=True)
    diem_the_luc = Column(String(20), nullable=True)
    kiem_tra_chinh_tri = Column(String(50), nullable=True)
    diem_kiem_tra_chinh_tri = Column(String(20), nullable=True)
    phieu_tin_nhiem = Column(String(50), nullable=True)
    xep_loai_dang_vien = Column(String(50), nullable=True)
    ket_qua_doan_the = Column(String(255), nullable=True)
    chu_tri_don_vi_danh_hieu = Column(String(255), nullable=True)

    # Lecturer-specific
    danh_hieu_gv_gioi = Column(String(100), nullable=True)
    tien_do_pgs = Column(String(255), nullable=True)
    dinh_muc_giang_day = Column(String(100), nullable=True)
    thoi_gian_lao_dong_kh = Column(String(100), nullable=True)
    ket_qua_kiem_tra_giang = Column(String(100), nullable=True)

    # Student-specific
    danh_hieu_hv_gioi = Column(String(100), nullable=True)
    diem_tong_ket = Column(String(50), nullable=True)
    ket_qua_thuc_hanh = Column(String(100), nullable=True)

    # Scientific research
    diem_nckh = Column(Float, nullable=True)
    nckh_noi_dung = Column(Text, nullable=True)
    nckh_minh_chung = Column(String(255), nullable=True)

    thanh_tich_ca_nhan_khac = Column(Text, nullable=True)

    ghi_chu = Column(Text, nullable=True)
    created_at = Column(DateTime, default=func.now())
    updated_at = Column(DateTime, default=func.now(), onupdate=func.now())

    de_xuat = relationship('DeXuat', back_populates='chi_tiets')
    quan_nhan = relationship('QuanNhan', back_populates='de_xuat_items')
    minh_chungs = relationship('MinhChung', back_populates='chi_tiet', cascade='all, delete-orphan')


class MinhChung(db.Model):
    __tablename__ = 'minh_chung'

    id = Column(Integer, primary_key=True, autoincrement=True)
    chi_tiet_id = Column(Integer, ForeignKey('de_xuat_chi_tiet.id'), nullable=False, index=True)
    loai_minh_chung = Column(String(100), nullable=False)
    duong_dan = Column(String(255), nullable=False)
    ten_file_goc = Column(String(255), nullable=True)
    mo_ta = Column(String(255), nullable=True)
    created_at = Column(DateTime, default=func.now())

    chi_tiet = relationship('DeXuatChiTiet', back_populates='minh_chungs')
