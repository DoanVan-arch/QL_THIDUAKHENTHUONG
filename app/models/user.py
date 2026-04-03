import enum
from werkzeug.security import generate_password_hash, check_password_hash
from flask_login import UserMixin
from sqlalchemy import Column, Integer, String, Enum, Boolean, DateTime, ForeignKey, func
from sqlalchemy.orm import relationship
from app.extensions import db, login_manager


class Role(enum.Enum):
    UNIT_USER = 'unit_user'
    PHONG_CHINHTRI = 'phong_chinhTri'
    PHONG_THAMMUU = 'phong_thamMuu'
    PHONG_KHOAHOC = 'phong_khoaHoc'
    PHONG_DAOTAO = 'phong_daoTao'
    BAN_CANBO = 'ban_canBo'
    BAN_QUANLUC = 'ban_quanLuc'
    ADMIN = 'admin'


ROLE_DISPLAY = {
    Role.UNIT_USER: 'Đơn vị',
    Role.PHONG_CHINHTRI: 'Phòng Chính trị',
    Role.PHONG_THAMMUU: 'Phòng Tham mưu - Hành chính',
    Role.PHONG_KHOAHOC: 'Phòng Khoa học quân sự',
    Role.PHONG_DAOTAO: 'Phòng Đào tạo',
    Role.BAN_CANBO: 'Ban Cán bộ',
    Role.BAN_QUANLUC: 'Ban Quân lực',
    Role.ADMIN: 'Cơ quan Tuyên huấn (Admin)',
}


class User(db.Model, UserMixin):
    __tablename__ = 'users'

    id = Column(Integer, primary_key=True, autoincrement=True)
    username = Column(String(80), unique=True, nullable=False, index=True)
    password_hash = Column(String(256), nullable=False)
    ho_ten = Column(String(100), nullable=False)
    role = Column(Enum(Role), nullable=False)
    don_vi_id = Column(Integer, ForeignKey('don_vi.id'), nullable=True)
    is_active_account = Column(Boolean, default=True)
    created_at = Column(DateTime, default=func.now())
    updated_at = Column(DateTime, default=func.now(), onupdate=func.now())

    don_vi = relationship('DonVi', back_populates='user_account')

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    @property
    def is_active(self):
        return self.is_active_account

    @property
    def role_display(self):
        return ROLE_DISPLAY.get(self.role, self.role.value)

    @property
    def is_department(self):
        return self.role in (
            Role.PHONG_CHINHTRI, Role.PHONG_THAMMUU,
            Role.PHONG_KHOAHOC, Role.PHONG_DAOTAO,
            Role.BAN_CANBO, Role.BAN_QUANLUC,
        )

    @property
    def is_admin(self):
        return self.role == Role.ADMIN

    @property
    def is_unit_user(self):
        return self.role == Role.UNIT_USER


@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))
