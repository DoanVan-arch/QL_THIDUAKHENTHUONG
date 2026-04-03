from flask import Blueprint, render_template
from flask_login import login_required, current_user
from app.models.user import Role
from app.models.unit import DonVi
from app.models.personnel import QuanNhan
from app.models.nomination import DeXuat, TrangThaiDeXuat
from app.models.approval import PheDuyet, KetQuaDuyet

dashboard_bp = Blueprint('dashboard', __name__)


@dashboard_bp.route('/')
@login_required
def index():
    if current_user.is_unit_user:
        return unit_dashboard()
    elif current_user.is_department:
        return department_dashboard()
    elif current_user.is_admin:
        return admin_dashboard()
    return render_template('dashboard/unit_dashboard.html')


@dashboard_bp.route('/huong-dan')
@login_required
def user_guide():
    """Hướng dẫn sử dụng hệ thống - accessible by all authenticated users."""
    return render_template('dashboard/user_guide.html')


def unit_dashboard():
    don_vi = current_user.don_vi
    if not don_vi:
        return render_template('dashboard/unit_dashboard.html', don_vi=None,
                               total_personnel=0, total_nominations=0,
                               pending_nominations=0, approved_nominations=0)

    total_personnel = QuanNhan.query.filter_by(don_vi_id=don_vi.id, is_active=True).count()
    total_nominations = DeXuat.query.filter_by(don_vi_id=don_vi.id).count()
    pending_nominations = DeXuat.query.filter_by(
        don_vi_id=don_vi.id, trang_thai=TrangThaiDeXuat.CHO_DUYET.value
    ).count()
    approved_nominations = DeXuat.query.filter_by(
        don_vi_id=don_vi.id, trang_thai=TrangThaiDeXuat.PHE_DUYET_CUOI.value
    ).count()

    recent_nominations = DeXuat.query.filter_by(don_vi_id=don_vi.id)\
        .order_by(DeXuat.ngay_tao.desc()).limit(5).all()

    return render_template('dashboard/unit_dashboard.html',
                           don_vi=don_vi,
                           total_personnel=total_personnel,
                           total_nominations=total_nominations,
                           pending_nominations=pending_nominations,
                           approved_nominations=approved_nominations,
                           recent_nominations=recent_nominations)


def department_dashboard():
    role_phong_map = {
        Role.PHONG_CHINHTRI: 'Phòng Chính trị',
        Role.PHONG_THAMMUU: 'Phòng Tham mưu',
        Role.PHONG_KHOAHOC: 'Phòng Khoa học',
        Role.PHONG_DAOTAO: 'Phòng Đào tạo',
        Role.BAN_CANBO: 'Ban Cán bộ',
        Role.BAN_QUANLUC: 'Ban Quân lực',
    }
    phong_name = role_phong_map.get(current_user.role, '')

    pending_count = PheDuyet.query.filter_by(
        phong_duyet=phong_name,
        ket_qua=KetQuaDuyet.CHO_DUYET.value
    ).count()

    approved_count = PheDuyet.query.filter_by(
        phong_duyet=phong_name,
        ket_qua=KetQuaDuyet.DONG_Y.value
    ).count()

    rejected_count = PheDuyet.query.filter_by(
        phong_duyet=phong_name,
        ket_qua=KetQuaDuyet.TU_CHOI.value
    ).count()

    recent_pending = PheDuyet.query.filter_by(
        phong_duyet=phong_name,
        ket_qua=KetQuaDuyet.CHO_DUYET.value
    ).order_by(PheDuyet.created_at.desc()).limit(10).all()

    return render_template('dashboard/department_dashboard.html',
                           phong_name=phong_name,
                           pending_count=pending_count,
                           approved_count=approved_count,
                           rejected_count=rejected_count,
                           recent_pending=recent_pending)


def admin_dashboard():
    total_units = DonVi.query.filter_by(is_active=True).count()
    total_personnel = QuanNhan.query.filter_by(is_active=True).count()
    total_nominations = DeXuat.query.count()

    awaiting_final = DeXuat.query.filter_by(
        trang_thai=TrangThaiDeXuat.DA_DUYET.value
    ).count()

    final_approved = DeXuat.query.filter_by(
        trang_thai=TrangThaiDeXuat.PHE_DUYET_CUOI.value
    ).count()

    recent_nominations = DeXuat.query.order_by(DeXuat.ngay_tao.desc()).limit(10).all()

    return render_template('dashboard/admin_dashboard.html',
                           total_units=total_units,
                           total_personnel=total_personnel,
                           total_nominations=total_nominations,
                           awaiting_final=awaiting_final,
                           final_approved=final_approved,
                           recent_nominations=recent_nominations)
