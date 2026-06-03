import enum
import json
from sqlalchemy import Column, Integer, String, DateTime, ForeignKey, Text, func
from sqlalchemy.orm import relationship
from app.extensions import db


class TrangThaiYeuCauSua(enum.Enum):
    CHO_SUA = 'cho_sua'      # waiting for the unit to edit
    DA_SUA = 'da_sua'        # unit has edited and resubmitted
    HUY = 'huy'              # cancelled (e.g. item removed)


class YeuCauChinhSua(db.Model):
    """A request from an approving department asking the unit to fix one or more
    criteria of a single cá nhân/tập thể. The unit may edit ONLY the flagged
    criteria; once resubmitted, only the requesting department must re-review."""
    __tablename__ = 'yeu_cau_chinh_sua'

    id = Column(Integer, primary_key=True, autoincrement=True)
    de_xuat_id = Column(Integer, ForeignKey('de_xuat.id'), nullable=False, index=True)
    chi_tiet_id = Column(Integer, ForeignKey('de_xuat_chi_tiet.id'), nullable=False, index=True)
    phong_yeu_cau = Column(String(100), nullable=False)  # PhongDuyet value of requester
    nguoi_yeu_cau_id = Column(Integer, ForeignKey('users.id'), nullable=True)
    _cac_truong = Column('cac_truong', Text, nullable=True)  # JSON list of ma_truong (field names)
    ly_do = Column(Text, nullable=True)
    trang_thai = Column(String(20), nullable=False, default=TrangThaiYeuCauSua.CHO_SUA.value)
    created_at = Column(DateTime, default=func.now())
    updated_at = Column(DateTime, default=func.now(), onupdate=func.now())
    ngay_sua = Column(DateTime, nullable=True)

    de_xuat = relationship('DeXuat')
    chi_tiet = relationship('DeXuatChiTiet')
    nguoi_yeu_cau = relationship('User')

    @property
    def cac_truong(self):
        if self._cac_truong:
            try:
                return json.loads(self._cac_truong)
            except (json.JSONDecodeError, TypeError):
                return []
        return []

    @cac_truong.setter
    def cac_truong(self, value):
        if isinstance(value, list):
            self._cac_truong = json.dumps(value, ensure_ascii=False)
        else:
            self._cac_truong = value
