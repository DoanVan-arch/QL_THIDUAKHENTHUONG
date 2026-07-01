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
    loai_input = Column(String(20), nullable=False, default='textbox')  # 'textbox' | 'combobox'
    _gia_tri_chon = Column('gia_tri_chon', Text, nullable=True)    # JSON list of options for combobox
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

    @property
    def gia_tri_chon(self):
        if self._gia_tri_chon:
            try:
                return json.loads(self._gia_tri_chon)
            except (json.JSONDecodeError, TypeError):
                return []
        return []

    @gia_tri_chon.setter
    def gia_tri_chon(self, value):
        if isinstance(value, list):
            self._gia_tri_chon = json.dumps(value, ensure_ascii=False)
        elif isinstance(value, str) and value.strip():
            # Accept newline-separated string from textarea
            items = [v.strip() for v in value.splitlines() if v.strip()]
            self._gia_tri_chon = json.dumps(items, ensure_ascii=False)
        else:
            self._gia_tri_chon = None

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
    HOI_DONG = 'Hội đồng'
    PHE_DUYET_CUOI = 'Phê duyệt cuối'
    TU_CHOI = 'Từ chối'


class TrangThaiChiTiet(enum.Enum):
    """Per-individual/unit item approval stage."""
    NHAP          = 'nhap'           # de_xuat still draft
    DANG_DUYET    = 'dang_duyet'     # in departmental review (Bảng 1 – pending)
    DA_DUYET      = 'da_duyet'       # all depts approved → Bảng 1 admin view (HOI_DONG)
    HOI_DONG      = 'hoi_dong'       # admin_approved=True → Bảng 2 (hội đồng voting)
    PHE_DUYET_CUOI= 'phe_duyet_cuoi' # KhenThuong created → Bảng 3 (reward list)
    TU_CHOI       = 'tu_choi'        # bi_loai=True (rejected / removed)


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
    def chi_tiets_active(self):
        """Chi_tiets still in the approval process.
        Hide only items definitively rejected by the admin/Tuyên huấn board:
        bi_loai=True AND phong_loai='Tuyên huấn' AND trang_thai='tu_choi'.
        Items rejected by a single department (phong_loai = other dept name) remain visible
        since the overall đề xuất may still be in progress elsewhere.
        """
        return [
            c for c in self.chi_tiets
            if not (c.bi_loai and c.phong_loai == 'Tuyên huấn' and c.trang_thai == TrangThaiChiTiet.TU_CHOI.value)
        ]
   
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
    xep_loai_doan_vien = Column(String(50), nullable=True)
    hinh_thuc_khen_thuong_qc = Column(String(255), nullable=True)
    ket_qua_phu_nu = Column(String(255), nullable=True)
    hinh_thuc_khen_thuong_pn = Column(String(255), nullable=True)
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
    xep_loai_tong_ket = Column(String(50), nullable=True)
    ket_qua_thuc_hanh = Column(String(100), nullable=True)
    ket_qua_ren_luyen = Column(String(100), nullable=True)

    # Graduation exam (hoc_cuoi)
    hinh_thuc_tot_nghiep = Column(String(100), nullable=True)
    diem_tn_ctd = Column(String(50), nullable=True)
    diem_tn_ct = Column(String(50), nullable=True)
    diem_tn_ta = Column(String(50), nullable=True)
    diem_tn_mon4 = Column(String(50), nullable=True)
    diem_tn_chuyennganh = Column(String(50), nullable=True)
    diem_tn_baove = Column(String(50), nullable=True)

    # Scientific research
    diem_nckh = Column(Float, nullable=True)
    nckh_noi_dung = Column(Text, nullable=True)
    nckh_minh_chung = Column(String(255), nullable=True)
    mo_ta_khoa_hoc = Column(Text, nullable=True)

    # Graduation overall score (average) + other-achievement evidence file path
    diem_tot_nghiep = Column(Float, nullable=True)
    minh_chung_thanh_tich_khac = Column(String(255), nullable=True)

    thanh_tich_ca_nhan_khac = Column(Text, nullable=True)

    # For collective (tap_the/don_vi) nominations — unit name being proposed
    ten_don_vi_de_xuat = Column(String(255), nullable=True)

    # JSON blob for collective-specific criteria values (DVQT / DVTT)
    tap_the_data = Column(Text, nullable=True)

    ghi_chu = Column(Text, nullable=True)

    # Admin pre-approval flag (Bảng 1 → Bảng 2 transition)
    admin_approved = Column(db.Boolean, default=False, nullable=False, server_default='0')

    # Soft-remove: when an approving department rejects this single cá nhân/tập thể,
    # the item is removed from the active approval process (hidden from pending/tracking)
    # while the rest of the đề xuất continues. Kept in DB for audit / possible restore.
    bi_loai = Column(db.Boolean, default=False, nullable=False, server_default='0')
    ly_do_loai = Column(Text, nullable=True)
    phong_loai = Column(String(100), nullable=True)
    ngay_loai = Column(DateTime, nullable=True)

    # Per-item approval stage (tracked separately from parent de_xuat.trang_thai)
    trang_thai = Column(String(30), nullable=False, default=TrangThaiChiTiet.NHAP.value,
                        server_default=TrangThaiChiTiet.NHAP.value)

    created_at = Column(DateTime, default=func.now())
    updated_at = Column(DateTime, default=func.now(), onupdate=func.now())

    @property
    def tap_the_dict(self):
        if self.tap_the_data:
            try:
                return json.loads(self.tap_the_data)
            except (json.JSONDecodeError, TypeError):
                return {}
        return {}

    @tap_the_dict.setter
    def tap_the_dict(self, value):
        if isinstance(value, dict):
            self.tap_the_data = json.dumps(value, ensure_ascii=False)
        else:
            self.tap_the_data = None

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
