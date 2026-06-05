from sqlalchemy import Column, Integer, String, DateTime, ForeignKey, Text, func
from sqlalchemy.orm import relationship
from app.extensions import db


class ActivityLog(db.Model):
    """Audit trail — every significant user action is recorded here."""
    __tablename__ = 'activity_log'

    id = Column(Integer, primary_key=True, autoincrement=True)

    # Who performed the action (nullable so system/anonymous actions can be stored)
    user_id = Column(Integer, ForeignKey('users.id', ondelete='SET NULL'), nullable=True, index=True)
    username = Column(String(80), nullable=True)    # denormalised — kept even after user deletion
    ho_ten = Column(String(100), nullable=True)
    role = Column(String(50), nullable=True)

    # What happened
    action = Column(String(100), nullable=False, index=True)   # e.g. 'login', 'submit_nomination'
    resource_type = Column(String(50), nullable=True)          # e.g. 'de_xuat', 'chi_tiet', 'user'
    resource_id = Column(Integer, nullable=True)               # PK of the affected row
    detail = Column(Text, nullable=True)                       # free-text description

    # Where / when
    ip_address = Column(String(45), nullable=True)
    created_at = Column(DateTime, default=func.now(), index=True)

    user = relationship('User', foreign_keys=[user_id], passive_deletes=True)

    # ── Human-readable action labels ──────────────────────────────────────────
    ACTION_LABELS = {
        'login':                     'Đăng nhập',
        'logout':                    'Đăng xuất',
        'login_failed':              'Đăng nhập thất bại',
        'create_nomination':         'Tạo đề xuất',
        'submit_nomination':         'Gửi đề xuất duyệt',
        'edit_nomination':           'Chỉnh sửa đề xuất',
        'add_chi_tiet':              'Thêm cá nhân/tập thể vào đề xuất',
        'delete_chi_tiet':           'Xóa cá nhân/tập thể khỏi đề xuất',
        'dept_approve':              'Cơ quan duyệt đồng ý',
        'dept_approve_item':         'Cơ quan duyệt đồng ý (cá nhân)',
        'dept_reject':               'Cơ quan từ chối',
        'dept_reject_item':          'Cơ quan từ chối (cá nhân)',
        'dept_submit_review':        'Cơ quan hoàn thành phê duyệt',
        'admin_pre_approve':         'Admin chuyển sang Hội đồng (Bảng 2)',
        'admin_final_approve':       'Admin phê duyệt cuối (cá nhân)',
        'batch_final_approve':       'Admin phê duyệt cuối (hàng loạt)',
        'admin_reject':              'Admin từ chối đề xuất',
        'admin_reject_individual':   'Admin từ chối cá nhân',
        'admin_reject_nomination':   'Admin từ chối đề xuất',
        'hoi_dong_vote':             'Hội đồng biểu quyết',
        'revoke_final_approval':     'Thu hồi phê duyệt cuối',
        'create_user':               'Tạo tài khoản',
        'edit_user':                 'Chỉnh sửa tài khoản',
        'delete_user':               'Xóa tài khoản',
        'create_personnel':          'Thêm quân nhân',
        'edit_personnel':            'Chỉnh sửa quân nhân',
        'delete_personnel':          'Xóa quân nhân',
        'import_personnel':          'Nhập quân nhân từ Excel',
    }

    def action_label(self):
        return self.ACTION_LABELS.get(self.action, self.action)
