"""
ChuyenDonVi — personnel transfer request between units.

Flow:
  source unit  → tạo yêu cầu (PENDING)
  target unit  → xác nhận (CONFIRMED) hoặc từ chối (REJECTED)
  On CONFIRMED: QuanNhan.don_vi_id updated to don_vi_dich_id
"""
from datetime import datetime
from app.extensions import db


class TrangThaiChuyen:
    PENDING   = 'PENDING'     # Chờ đơn vị tiếp nhận xác nhận
    CONFIRMED = 'CONFIRMED'   # Đã xác nhận — quân nhân đã chuyển
    REJECTED  = 'REJECTED'    # Đơn vị tiếp nhận từ chối


class ChuyenDonVi(db.Model):
    __tablename__ = 'chuyen_don_vi'

    id               = db.Column(db.Integer, primary_key=True)
    quan_nhan_id     = db.Column(db.Integer, db.ForeignKey('quan_nhan.id'), nullable=False, index=True)
    don_vi_nguon_id  = db.Column(db.Integer, db.ForeignKey('don_vi.id'),    nullable=False)
    don_vi_dich_id   = db.Column(db.Integer, db.ForeignKey('don_vi.id'),    nullable=False)
    nguoi_tao_id     = db.Column(db.Integer, db.ForeignKey('users.id'),     nullable=True)
    nguoi_xac_nhan_id = db.Column(db.Integer, db.ForeignKey('users.id'),    nullable=True)
    trang_thai       = db.Column(db.String(20), nullable=False, default=TrangThaiChuyen.PENDING)
    ly_do            = db.Column(db.Text, nullable=True)          # reason from source
    ghi_chu          = db.Column(db.Text, nullable=True)          # note from target on confirm/reject
    ngay_tao         = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    ngay_xu_ly       = db.Column(db.DateTime, nullable=True)      # confirm or reject datetime

    # Relationships
    quan_nhan      = db.relationship('QuanNhan', foreign_keys=[quan_nhan_id], backref='chuyen_don_vis')
    don_vi_nguon   = db.relationship('DonVi', foreign_keys=[don_vi_nguon_id])
    don_vi_dich    = db.relationship('DonVi', foreign_keys=[don_vi_dich_id])
    nguoi_tao      = db.relationship('User', foreign_keys=[nguoi_tao_id])
    nguoi_xac_nhan = db.relationship('User', foreign_keys=[nguoi_xac_nhan_id])

    @property
    def is_pending(self):
        return self.trang_thai == TrangThaiChuyen.PENDING

    @property
    def is_confirmed(self):
        return self.trang_thai == TrangThaiChuyen.CONFIRMED

    @property
    def is_rejected(self):
        return self.trang_thai == TrangThaiChuyen.REJECTED
