from functools import wraps
from flask import abort, redirect, url_for
from flask_login import current_user
from app.models.user import Role


def role_required(*roles):
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if not current_user.is_authenticated:
                return redirect(url_for('auth.login'))
            if current_user.role not in roles:
                abort(403)
            return f(*args, **kwargs)
        return decorated_function
    return decorator


def unit_user_required(f):
    return role_required(Role.UNIT_USER)(f)


def department_required(f):
    return role_required(
        Role.PHONG_KHOAHOC, Role.PHONG_DAOTAO,
        Role.THU_TRUONG_PHONG_CHINHTRI, Role.THU_TRUONG_PHONG_TMHC,
        Role.BAN_CANBO, Role.BAN_TOCHUC, Role.BAN_TUYENHUAN,
        Role.BAN_CTCQ, Role.BAN_CNTT, Role.BAN_TAC_HUAN,
        Role.BAN_KHAOTHI, Role.UY_BAN_KIEMTRA,
        Role.BAN_QUANLUC,
    )(f)


def admin_required(f):
    return role_required(Role.ADMIN)(f)


def admin_or_department_required(f):
    return role_required(
        Role.ADMIN, Role.PHONG_KHOAHOC, Role.PHONG_DAOTAO,
        Role.THU_TRUONG_PHONG_CHINHTRI, Role.THU_TRUONG_PHONG_TMHC,
        Role.BAN_CANBO, Role.BAN_TOCHUC, Role.BAN_TUYENHUAN,
        Role.BAN_CTCQ, Role.BAN_CNTT, Role.BAN_TAC_HUAN,
        Role.BAN_KHAOTHI, Role.UY_BAN_KIEMTRA,
        Role.BAN_QUANLUC,
    )(f)


def admin_or_reward_viewer_required(f):
    return role_required(
        Role.ADMIN,
        Role.BAN_TUYENHUAN,
        Role.BAN_CANBO,
        Role.BAN_TOCHUC,
        Role.BAN_BAOVE_ANNINH,
        Role.BAN_CTCQ,
        Role.BAN_KEHOACH_TONGHOP,
        Role.UY_BAN_KIEMTRA,
    )(f)
