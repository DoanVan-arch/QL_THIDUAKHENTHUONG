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
    THU_TRUONG_PHONG_CHINHTRI = 'thu_truongPhongChinhTri'
    THU_TRUONG_PHONG_TMHC = 'thu_truongPhongTMHC'
    BAN_CANBO = 'ban_canBo'
    BAN_TOCHUC = 'ban_toChuc'
    BAN_TUYENHUAN = 'ban_tuyenHuan'
    BAN_CTCQ = 'ban_congTacQuanChung'
    BAN_CNTT = 'ban_congNgheThongTin'
    BAN_TAC_HUAN = 'ban_tacHuan'
    BAN_KHAOTHI = 'ban_khaoThi'
    BAN_BAOVE_ANNINH = 'ban_baoVeAnNinh'
    BAN_KEHOACH_TONGHOP = 'ban_keHoachTongHop'
    UY_BAN_KIEMTRA = 'uy_banKiemTra'
    BAN_QUANLUC = 'ban_quanLuc'
    ADMIN = 'admin'


ROLE_DISPLAY = {
    Role.UNIT_USER: 'Đơn vị',
    Role.PHONG_CHINHTRI: 'Phòng Chính trị',
    Role.PHONG_THAMMUU: 'Phòng Tham mưu - Hành chính',
    Role.PHONG_KHOAHOC: 'Phòng Khoa học quân sự',
    Role.PHONG_DAOTAO: 'Phòng Đào tạo',
    Role.THU_TRUONG_PHONG_CHINHTRI: 'Thủ trưởng Phòng Chính trị',
    Role.THU_TRUONG_PHONG_TMHC: 'Thủ trưởng Phòng TM-HC',
    Role.BAN_CANBO: 'Ban Cán bộ',
    Role.BAN_TOCHUC: 'Ban Tổ chức',
    Role.BAN_TUYENHUAN: 'Ban Tuyên huấn',
    Role.BAN_CTCQ: 'Ban Công tác quần chúng',
    Role.BAN_CNTT: 'Ban Công nghệ thông tin',
    Role.BAN_TAC_HUAN: 'Ban Tác huấn',
    Role.BAN_KHAOTHI: 'Ban Khảo thí',
    Role.BAN_BAOVE_ANNINH: 'Ban Bảo vệ an ninh',
    Role.BAN_KEHOACH_TONGHOP: 'Ban Kế hoạch tổng hợp',
    Role.UY_BAN_KIEMTRA: 'Ủy ban Kiểm tra',
    Role.BAN_QUANLUC: 'Ban Quân lực',
    Role.ADMIN: 'Ban thư ký Hội đồng thi đua khen thưởng',
}


class User(db.Model, UserMixin):
    __tablename__ = 'users'

    id = Column(Integer, primary_key=True, autoincrement=True)
    username = Column(String(80), unique=True, nullable=False, index=True)
    password_hash = Column(String(256), nullable=False)
    ho_ten = Column(String(100), nullable=False)
    role = Column(Enum(Role, values_callable=lambda x: [e.value for e in x], native_enum=False), nullable=False)
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
            Role.PHONG_KHOAHOC, Role.PHONG_DAOTAO,
            Role.THU_TRUONG_PHONG_CHINHTRI, Role.THU_TRUONG_PHONG_TMHC,
            Role.BAN_CANBO, Role.BAN_TOCHUC, Role.BAN_TUYENHUAN,
            Role.BAN_CTCQ, Role.BAN_CNTT, Role.BAN_TAC_HUAN,
            Role.BAN_KHAOTHI,
            Role.UY_BAN_KIEMTRA,
            Role.BAN_QUANLUC,
        )

    @property
    def is_admin(self):
        return self.role == Role.ADMIN

    @property
    def is_unit_user(self):
        return self.role == Role.UNIT_USER

    @property
    def is_reward_viewer(self):
        return self.role in (
            Role.BAN_TUYENHUAN,
            Role.BAN_CANBO,
            Role.BAN_TOCHUC,
            Role.BAN_BAOVE_ANNINH,
            Role.BAN_CTCQ,
            Role.BAN_KEHOACH_TONGHOP,
            Role.UY_BAN_KIEMTRA,
        )


@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))
