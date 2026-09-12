from sqlalchemy import Column, Integer, String, DateTime, Date, ForeignKey, Text, func
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


class LoaiKhenThuongTapThe:
    """Loại khen thưởng tập thể (tự nhập)."""
    BANG_KHEN = 'Bằng khen'
    GIAY_KHEN = 'Giấy khen'

    @classmethod
    def choices(cls):
        return [cls.BANG_KHEN, cls.GIAY_KHEN]


class KhenThuongTapThe(db.Model):
    """★ Danh sách khen thưởng tập thể do đơn vị tự quản lý (nomination.reward_list_tap_the).

    Mỗi bản ghi có thể được:
    - Tự nhập thủ công (nguon='thu_cong'): điền đầy đủ các trường.
    - Lấy từ Bảng 3 (nguon='bang_3'): chỉ điền `ten` và `nam_hoc` lấy từ
      KhenThuong (loai_danh_hieu ĐVQT/ĐVTT), các trường khác để trống cho đơn
      vị tự bổ sung sau (loai, so_quyet_dinh, ngay_cap, cap_quyet_dinh,
      anh_minh_chung).
    """
    __tablename__ = 'khen_thuong_tap_the'

    id = Column(Integer, primary_key=True, autoincrement=True)
    don_vi_id = Column(Integer, ForeignKey('don_vi.id'), nullable=False, index=True)
    # Liên kết tới bản ghi Bảng 3 (KhenThuong) nếu được lấy từ đó — tránh import trùng
    khen_thuong_id = Column(Integer, ForeignKey('khen_thuong.id'), nullable=True, index=True)

    nguon = Column(String(20), nullable=False, default='thu_cong')  # 'thu_cong' | 'bang_3'

    loai = Column(String(50), nullable=True)          # 'Bằng khen' / 'Giấy khen'
    ten = Column(String(255), nullable=False)
    so_quyet_dinh = Column(String(100), nullable=True)
    ngay_cap = Column(Date, nullable=True)
    nam_hoc = Column(String(20), nullable=True)
    cap_quyet_dinh = Column(String(255), nullable=True)
    anh_minh_chung = Column(String(255), nullable=True)

    created_by_id = Column(Integer, ForeignKey('users.id'), nullable=True)
    created_at = Column(DateTime, default=func.now())
    updated_at = Column(DateTime, default=func.now(), onupdate=func.now())

    don_vi = relationship('DonVi')
    khen_thuong = relationship('KhenThuong')
    created_by = relationship('User')

