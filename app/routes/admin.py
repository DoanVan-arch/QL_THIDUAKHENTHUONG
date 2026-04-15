from io import BytesIO
from types import SimpleNamespace
from flask import Blueprint, render_template, redirect, url_for, flash, request, jsonify, send_file
from flask_login import login_required, current_user
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from app.extensions import db
from app.models.user import User, Role, ROLE_DISPLAY
from app.models.unit import DonVi, LoaiDonVi
from app.models.personnel import QuanNhan, CapBac, HocHam, HocVi, DoiTuong
from app.models.certificate import ChungChi, LoaiChungChi
from app.models.nomination import DeXuat, DeXuatChiTiet, TrangThaiDeXuat, LoaiDanhHieu, DanhHieu, TieuChi
from app.models.approval import PheDuyet, PhongDuyet, KetQuaDuyet, KetQuaDuyetChiTiet
from app.models.reward import KhenThuong
from app.models.catalog import ChucVuOption, CapBacOption
from app.utils.decorators import admin_required, admin_or_reward_viewer_required
from app.utils.file_upload import save_upload, delete_upload
from datetime import datetime

admin_bp = Blueprint('admin', __name__)

# The six reviewing departments (excluding admin)
DEPT_NAMES = [
    'Phòng Khoa học', 'Phòng Đào tạo',
    'Thủ trưởng Phòng Chính trị', 'Thủ trưởng Phòng TM-HC',
    'Ban Cán bộ', 'Ban Tổ chức', 'Ban Tuyên huấn', 'Ban Công tác quần chúng',
    'Ban Công nghệ thông tin', 'Ban Tác huấn', 'Ban Khảo thí', 'Ủy ban Kiểm tra', 'Ban Quân lực'
]

# Display order for approval columns in tracking screens
TRACKING_DEPT_COLUMNS = [
    {'key': 'Ban Cán bộ', 'label': 'Ban Cán bộ'},
    {'key': 'Ban Tổ chức', 'label': 'Ban Tổ chức'},
    {'key': 'Ban Tuyên huấn', 'label': 'Ban Tuyên huấn'},
    {'key': 'Ban Công tác quần chúng', 'label': 'Ban Công tác quần chúng'},
    {'key': 'Thủ trưởng Phòng Chính trị', 'label': 'TT phòng Chính trị'},
    {'key': 'Ban Công nghệ thông tin', 'label': 'Ban Công nghệ thông tin'},
    {'key': 'Ban Tác huấn', 'label': 'Ban Tác huấn'},
    {'key': 'Ban Quân lực', 'label': 'Ban Quân lực'},
    {'key': 'Thủ trưởng Phòng TM-HC', 'label': 'TT phòng TM-HC'},
    {'key': 'Phòng Đào tạo', 'label': '(Phòng Đào tạo)'},
    {'key': 'Phòng Khoa học', 'label': '(Phòng Khoa học quân sự)'},
    {'key': 'Ban Khảo thí', 'label': 'Ban Khảo thí'},
    {'key': 'Ủy ban Kiểm tra', 'label': 'Ủy ban Kiểm tra'},
]


def _is_auto_scope_approved(dept_name, doi_tuong):
    if dept_name == 'Ban Quân lực':
        return doi_tuong not in BAN_QUANLUC_DOI_TUONG
    if dept_name == 'Ban Cán bộ':
        return doi_tuong in BAN_QUANLUC_DOI_TUONG
    return False

# Đối tượng thuộc diện Ban Quân lực quản lý
BAN_QUANLUC_DOI_TUONG = ['Công nhân viên', 'Quân nhân chuyên nghiệp', 'Công chức quốc phòng']

# All criteria field labels (for detailed view)
ALL_FIELD_LABELS = {
    'muc_do_hoan_thanh': 'Hoàn thành NV',
    'phieu_tin_nhiem': 'Phiếu tín nhiệm',
    'kiem_tra_dieu_lenh': 'KT Điều lệnh',
    'ban_sung': 'Bắn súng',
    'the_luc': 'Thể lực',
    'kiem_tra_chinh_tri': 'KT Chính trị',
    'kiem_tra_tin_hoc': 'Kỹ năng số',
    'dia_ly_quan_su': 'ĐHQS',
    'xep_loai_dang_vien': 'Xếp loại đảng viên',
    'danh_hieu_gv_gioi': 'DH GV giỏi',
    'dinh_muc_giang_day': 'Định mức GD',
    'ket_qua_kiem_tra_giang': 'KT giảng',
    'thoi_gian_lao_dong_kh': 'LĐ khoa học',
    'tien_do_pgs': 'Tiến độ PGS',
    'danh_hieu_hv_gioi': 'DH HV giỏi',
    'diem_tong_ket': 'Điểm TK',
    'ket_qua_thuc_hanh': 'Thực hành',
    'ket_qua_doan_the': 'KQ Đoàn thể',
    'chu_tri_don_vi_danh_hieu': 'Chủ trì ĐV',
    'diem_nckh': 'Điểm NCKH',
    'nckh_noi_dung': 'ND NCKH',
    'nckh_minh_chung': 'MC NCKH',
    'thanh_tich_ca_nhan_khac': 'Thành tích khác',
}

# Criteria fields in display order
ALL_FIELDS = [
    'muc_do_hoan_thanh', 'phieu_tin_nhiem',
    'kiem_tra_chinh_tri', 'kiem_tra_dieu_lenh', 'kiem_tra_tin_hoc',
    'dia_ly_quan_su', 'ban_sung', 'the_luc',
    'xep_loai_dang_vien',
    'ket_qua_doan_the', 'chu_tri_don_vi_danh_hieu',
    'danh_hieu_gv_gioi', 'dinh_muc_giang_day', 'ket_qua_kiem_tra_giang',
    'tien_do_pgs', 'thoi_gian_lao_dong_kh',
    'danh_hieu_hv_gioi', 'diem_tong_ket', 'ket_qua_thuc_hanh',
    'diem_nckh', 'nckh_noi_dung', 'nckh_minh_chung',
    'thanh_tich_ca_nhan_khac',
]


@admin_bp.route('/tracking')
@login_required
@admin_required
def approval_tracking():
    """View all nominations with per-individual rows grouped by unit."""
    status_filter = request.args.get('status', '')
    unit_filter = request.args.get('unit', '', type=str)
    danh_hieu_filter = request.args.get('danh_hieu', '')
    search_query = request.args.get('q', '').strip()
    scope_filter = request.args.get('scope', '')  # 'quan_luc' or 'can_bo'
    view_mode = request.args.get('view', 'compact')  # 'compact' or 'detail'

    query = DeXuat.query.filter(
        DeXuat.trang_thai != TrangThaiDeXuat.NHAP.value
    )

    if status_filter:
        query = query.filter(DeXuat.trang_thai == status_filter)

    if unit_filter:
        query = query.join(DonVi).filter(DonVi.ten_don_vi == unit_filter)

    nominations = query.order_by(DeXuat.ngay_gui.desc()).all()

    status_list = [e.value for e in TrangThaiDeXuat if e != TrangThaiDeXuat.NHAP]

    # Get unit names for filter dropdown
    unit_names_query = db.session.query(DonVi.ten_don_vi).join(
        DeXuat, DeXuat.don_vi_id == DonVi.id
    ).filter(
        DeXuat.trang_thai != TrangThaiDeXuat.NHAP.value
    ).distinct().order_by(DonVi.ten_don_vi).all()
    unit_names = [u[0] for u in unit_names_query]

    # Summary stats (always from full dataset, not filtered)
    all_query = DeXuat.query.filter(DeXuat.trang_thai != TrangThaiDeXuat.NHAP.value)
    stats = {
        'total': all_query.count(),
        'pending': all_query.filter(DeXuat.trang_thai == TrangThaiDeXuat.CHO_DUYET.value).count(),
        'reviewing': all_query.filter(DeXuat.trang_thai == TrangThaiDeXuat.DANG_DUYET.value).count(),
        'dept_approved': all_query.filter(DeXuat.trang_thai == TrangThaiDeXuat.DA_DUYET.value).count(),
        'final_approved': all_query.filter(DeXuat.trang_thai == TrangThaiDeXuat.PHE_DUYET_CUOI.value).count(),
        'rejected': all_query.filter(DeXuat.trang_thai == TrangThaiDeXuat.TU_CHOI.value).count(),
    }

    # Get all chi_tiet_ids that already have KhenThuong records (already final-approved)
    approved_ct_ids = set(
        row[0] for row in db.session.query(KhenThuong.chi_tiet_id).all()
    )

    # Group nominations by unit and build per-individual dept results
    # Structure: unit_groups = [{ unit_name, nominations: [{dx, chi_tiets: [{ct, dept_results, can_final_approve}]}] }]
    unit_groups_dict = {}  # unit_name -> list of nomination data
    total_individuals = 0

    tracking_dept_names = [c['key'] for c in TRACKING_DEPT_COLUMNS]

    for dx in nominations:
        unit_name = dx.don_vi.ten_don_vi
        if unit_name not in unit_groups_dict:
            unit_groups_dict[unit_name] = []

        # Build dept approval lookup: dept_name -> {ct_id -> KetQuaDuyetChiTiet}
        dept_lookup = {}
        for pd in dx.phe_duyets:
            dept_lookup[pd.phong_duyet] = {
                'phe_duyet': pd,
                'items': {kq.chi_tiet_id: kq for kq in pd.chi_tiet_duyet},
            }

        chi_tiets_data = []
        for ct in dx.chi_tiets:
            # Skip individuals already in KhenThuong (already final-approved)
            if ct.id in approved_ct_ids:
                continue

            # Filter by danh_hieu if specified
            if danh_hieu_filter and ct.loai_danh_hieu != danh_hieu_filter:
                continue

            # Filter by scope (Quân lực / Cán bộ)
            if scope_filter == 'quan_luc' and ct.doi_tuong not in BAN_QUANLUC_DOI_TUONG:
                continue
            if scope_filter == 'can_bo' and ct.doi_tuong in BAN_QUANLUC_DOI_TUONG:
                continue

            # Filter by search query (ho_ten)
            if search_query:
                ho_ten = ct.quan_nhan.ho_ten if ct.quan_nhan else ''
                if search_query.lower() not in ho_ten.lower():
                    continue

            ct_dept_results = {}
            all_dept_ok = True
            for dept_name in tracking_dept_names:
                dept_data = dept_lookup.get(dept_name)
                is_auto = _is_auto_scope_approved(dept_name, ct.doi_tuong)
                if dept_data:
                    kq = dept_data['items'].get(ct.id)
                    if kq:
                        ct_dept_results[dept_name] = {
                            'ket_qua': kq.ket_qua,
                            'auto': is_auto,
                        }
                    elif is_auto:
                        ct_dept_results[dept_name] = {
                            'ket_qua': KetQuaDuyet.DONG_Y.value,
                            'auto': True,
                        }
                    else:
                        ct_dept_results[dept_name] = None
                    if dept_name in DEPT_NAMES and not is_auto and (not kq or kq.ket_qua != KetQuaDuyet.DONG_Y.value):
                        all_dept_ok = False
                else:
                    if is_auto:
                        ct_dept_results[dept_name] = {
                            'ket_qua': KetQuaDuyet.DONG_Y.value,
                            'auto': True,
                        }
                    else:
                        ct_dept_results[dept_name] = None
                    if dept_name in DEPT_NAMES and not is_auto:
                        all_dept_ok = False

            # Individual can be final-approved if all 6 depts approved this person
            ct_can_approve = all_dept_ok

            chi_tiets_data.append({
                'ct': ct,
                'dept_results': ct_dept_results,
                'can_final_approve': ct_can_approve,
                'dx_id': dx.id,
            })
            total_individuals += 1

        # Only add nomination if it still has individuals to show
        if chi_tiets_data:
            unit_groups_dict[unit_name].append({
                'dx': dx,
                'chi_tiets': chi_tiets_data,
            })

    # Convert to ordered list
    unit_groups = []
    for unit_name in sorted(unit_groups_dict.keys()):
        nom_list = unit_groups_dict[unit_name]
        if not nom_list:
            continue
        individual_count = sum(len(n['chi_tiets']) for n in nom_list)
        unit_groups.append({
            'unit_name': unit_name,
            'nominations': nom_list,
            'individual_count': individual_count,
        })

    danh_hieu_list = [e.value for e in LoaiDanhHieu]

    return render_template('admin/tracking.html',
                           unit_groups=unit_groups,
                           status_filter=status_filter,
                           unit_filter=unit_filter,
                           danh_hieu_filter=danh_hieu_filter,
                           search_query=search_query,
                           scope_filter=scope_filter,
                           view_mode=view_mode,
                           status_list=status_list,
                           unit_names=unit_names,
                           danh_hieu_list=danh_hieu_list,
                           stats=stats,
                           dept_names=DEPT_NAMES,
                           tracking_dept_columns=TRACKING_DEPT_COLUMNS,
                           total_individuals=total_individuals,
                           all_field_labels=ALL_FIELD_LABELS,
                           all_fields=ALL_FIELDS)


@admin_bp.route('/tracking/chi-tiet/<int:ct_id>')
@login_required
@admin_required
def tracking_detail(ct_id):
    """View detailed info for ONE individual (DeXuatChiTiet)."""
    ct = DeXuatChiTiet.query.get_or_404(ct_id)
    de_xuat = ct.de_xuat

    # Build per-department results for this individual
    dept_item_results = {}
    for pd in de_xuat.phe_duyets:
        kq = KetQuaDuyetChiTiet.query.filter_by(
            phe_duyet_id=pd.id, chi_tiet_id=ct.id
        ).first()
        is_auto = _is_auto_scope_approved(pd.phong_duyet, ct.doi_tuong)
        if not kq and is_auto:
            kq = SimpleNamespace(ket_qua=KetQuaDuyet.DONG_Y.value, ly_do='Tự động duyệt theo phạm vi')
        dept_item_results[pd.phong_duyet] = {
            'phe_duyet': pd,
            'item_result': kq,
            'is_auto': is_auto,
        }

    # Check if all 6 departments approved THIS individual
    all_dept_ok = True
    for dept_name in DEPT_NAMES:
        if dept_name not in dept_item_results:
            all_dept_ok = False
            break
        kq = dept_item_results[dept_name].get('item_result')
        is_auto = _is_auto_scope_approved(dept_name, ct.doi_tuong)
        if is_auto:
            continue
        if not kq or kq.ket_qua != KetQuaDuyet.DONG_Y.value:
            all_dept_ok = False
            break

    # Check if already in KhenThuong
    already_approved = KhenThuong.query.filter_by(chi_tiet_id=ct.id).first() is not None

    can_final_approve = all_dept_ok and not already_approved

    # All criteria field labels for the detail view
    all_field_labels = {
        'muc_do_hoan_thanh': 'Hoàn thành nhiệm vụ',
        'phieu_tin_nhiem': 'Phiếu tín nhiệm',
        'kiem_tra_dieu_lenh': 'Kiểm tra điều lệnh',
        'ban_sung': 'Bắn súng',
        'the_luc': 'Thể lực',
        'kiem_tra_chinh_tri': 'Kiểm tra chính trị',
        'kiem_tra_tin_hoc': 'Kỹ năng số',
        'dia_ly_quan_su': 'Địa hình quân sự',
        'danh_hieu_gv_gioi': 'Danh hiệu GV giỏi',
        'dinh_muc_giang_day': 'Định mức giảng dạy',
        'ket_qua_kiem_tra_giang': 'Kết quả kiểm tra giảng',
        'thoi_gian_lao_dong_kh': 'Thời gian LĐ khoa học',
        'tien_do_pgs': 'Tiến độ PGS',
        'danh_hieu_hv_gioi': 'Danh hiệu HV giỏi',
        'diem_tong_ket': 'Điểm tổng kết',
        'ket_qua_thuc_hanh': 'Kết quả thực hành',
        'ket_qua_doan_the': 'Kết quả đoàn thể',
        'chu_tri_don_vi_danh_hieu': 'Chủ trì ĐV danh hiệu',
        'diem_nckh': 'Điểm NCKH',
        'nckh_noi_dung': 'Nội dung NCKH',
        'nckh_minh_chung': 'Minh chứng NCKH',
        'thanh_tich_ca_nhan_khac': 'Thành tích cá nhân khác',
    }

    # All field names in display order
    all_fields = [
        'muc_do_hoan_thanh', 'phieu_tin_nhiem',
        'kiem_tra_chinh_tri', 'kiem_tra_dieu_lenh', 'kiem_tra_tin_hoc',
        'dia_ly_quan_su', 'ban_sung', 'the_luc',
        'ket_qua_doan_the', 'chu_tri_don_vi_danh_hieu',
        'danh_hieu_gv_gioi', 'dinh_muc_giang_day', 'ket_qua_kiem_tra_giang',
        'tien_do_pgs', 'thoi_gian_lao_dong_kh',
        'danh_hieu_hv_gioi', 'diem_tong_ket', 'ket_qua_thuc_hanh',
        'diem_nckh', 'nckh_noi_dung', 'nckh_minh_chung',
        'thanh_tich_ca_nhan_khac',
    ]

    return render_template('admin/tracking_detail.html',
                           ct=ct,
                           de_xuat=de_xuat,
                           dept_item_results=dept_item_results,
                           dept_names=DEPT_NAMES,
                           tracking_dept_columns=TRACKING_DEPT_COLUMNS,
                           can_final_approve=can_final_approve,
                           already_approved=already_approved,
                           all_field_labels=all_field_labels,
                           all_fields=all_fields)


@admin_bp.route('/tracking/chi-tiet/<int:ct_id>/final-approve', methods=['POST'])
@login_required
@admin_required
def final_approve_individual(ct_id):
    """Final approval for a single individual - creates one KhenThuong record."""
    ct = DeXuatChiTiet.query.get_or_404(ct_id)
    de_xuat = ct.de_xuat

    # Check if already approved
    existing = KhenThuong.query.filter_by(chi_tiet_id=ct.id).first()
    if existing:
        flash('Cá nhân này đã được phê duyệt cuối.', 'warning')
        return redirect(url_for('admin.tracking_detail', ct_id=ct_id))

    # Verify all 6 departments approved this individual
    for dept_name in DEPT_NAMES:
        pd = PheDuyet.query.filter_by(
            de_xuat_id=de_xuat.id, phong_duyet=dept_name
        ).first()
        if not pd:
            flash(f'{dept_name} chưa duyệt đề xuất này.', 'warning')
            return redirect(url_for('admin.tracking_detail', ct_id=ct_id))
        kq = KetQuaDuyetChiTiet.query.filter_by(
            phe_duyet_id=pd.id, chi_tiet_id=ct.id
        ).first()
        if not kq or kq.ket_qua != KetQuaDuyet.DONG_Y.value:
            flash(f'{dept_name} chưa đồng ý cho cá nhân này.', 'warning')
            return redirect(url_for('admin.tracking_detail', ct_id=ct_id))

    now = datetime.utcnow()
    ghi_chu = request.form.get('ghi_chu', '').strip() or None

    # Create KhenThuong record for this individual
    khen_thuong = KhenThuong(
        de_xuat_id=de_xuat.id,
        chi_tiet_id=ct.id,
        quan_nhan_id=ct.quan_nhan_id,
        don_vi_id=de_xuat.don_vi_id,
        ho_ten=ct.quan_nhan.ho_ten if ct.quan_nhan else de_xuat.don_vi.ten_don_vi,
        cap_bac=ct.quan_nhan.cap_bac if ct.quan_nhan else None,
        chuc_vu=ct.quan_nhan.chuc_vu if ct.quan_nhan else None,
        doi_tuong=ct.doi_tuong,
        loai_danh_hieu=ct.loai_danh_hieu,
        nam_hoc=de_xuat.nam_hoc,
        nguoi_duyet_id=current_user.id,
        ngay_duyet=now,
        ghi_chu=ghi_chu,
    )
    db.session.add(khen_thuong)

    # Check if ALL individuals in this nomination now have KhenThuong records
    # If so, update nomination status to 'Phê duyệt cuối'
    all_ct_ids = {c.id for c in de_xuat.chi_tiets}
    approved_ct_ids = set(
        row[0] for row in db.session.query(KhenThuong.chi_tiet_id).filter(
            KhenThuong.de_xuat_id == de_xuat.id
        ).all()
    )
    # Include the one we just created (not yet committed)
    approved_ct_ids.add(ct.id)

    if all_ct_ids <= approved_ct_ids:
        # All individuals approved -> mark nomination as final
        de_xuat.trang_thai = TrangThaiDeXuat.PHE_DUYET_CUOI.value
        # Update admin PheDuyet record
        admin_pd = PheDuyet.query.filter_by(
            de_xuat_id=de_xuat.id, phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value
        ).first()
        if admin_pd:
            admin_pd.ket_qua = KetQuaDuyet.DONG_Y.value
            admin_pd.nguoi_duyet_id = current_user.id
            admin_pd.ngay_duyet = now
            admin_pd.ghi_chu = ghi_chu

    db.session.commit()
    ho_ten = ct.quan_nhan.ho_ten if ct.quan_nhan else de_xuat.don_vi.ten_don_vi
    flash(f'Đã phê duyệt cuối cho "{ho_ten}" và lưu vào danh sách khen thưởng.', 'success')
    return redirect(url_for('admin.approval_tracking'))


@admin_bp.route('/tracking/<int:id>/final-approve', methods=['POST'])
@login_required
@admin_required
def final_approve_from_tracking(id):
    """Final approval for entire nomination - creates KhenThuong records for all approved individuals."""
    de_xuat = DeXuat.query.get_or_404(id)

    if de_xuat.trang_thai != TrangThaiDeXuat.DA_DUYET.value:
        flash('Đề xuất này chưa được tất cả các cơ quan phê duyệt.', 'warning')
        return redirect(url_for('admin.approval_tracking'))

    # Verify all 6 departments approved
    for dept_name in DEPT_NAMES:
        pd = PheDuyet.query.filter_by(
            de_xuat_id=id, phong_duyet=dept_name
        ).first()
        if not pd or pd.ket_qua != KetQuaDuyet.DONG_Y.value:
            flash(f'{dept_name} chưa phê duyệt xong.', 'warning')
            return redirect(url_for('admin.approval_tracking'))

    # Update admin PheDuyet record
    admin_pd = PheDuyet.query.filter_by(
        de_xuat_id=id, phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value
    ).first()

    if admin_pd:
        admin_pd.ket_qua = KetQuaDuyet.DONG_Y.value
        admin_pd.nguoi_duyet_id = current_user.id
        admin_pd.ngay_duyet = datetime.utcnow()
        admin_pd.ghi_chu = request.form.get('ghi_chu', '').strip() or None

    de_xuat.trang_thai = TrangThaiDeXuat.PHE_DUYET_CUOI.value

    # Create KhenThuong records for each approved individual
    now = datetime.utcnow()
    for ct in de_xuat.chi_tiets:
        # Check that this individual was approved by all departments
        all_approved = True
        for dept_name in DEPT_NAMES:
            pd = PheDuyet.query.filter_by(
                de_xuat_id=id, phong_duyet=dept_name
            ).first()
            if pd:
                kq = KetQuaDuyetChiTiet.query.filter_by(
                    phe_duyet_id=pd.id, chi_tiet_id=ct.id
                ).first()
                if not kq or kq.ket_qua != KetQuaDuyet.DONG_Y.value:
                    all_approved = False
                    break
            else:
                all_approved = False
                break

        if all_approved:
            # Check if record already exists
            existing = KhenThuong.query.filter_by(
                de_xuat_id=de_xuat.id, chi_tiet_id=ct.id
            ).first()
            if not existing:
                khen_thuong = KhenThuong(
                    de_xuat_id=de_xuat.id,
                    chi_tiet_id=ct.id,
                    quan_nhan_id=ct.quan_nhan_id,
                    don_vi_id=de_xuat.don_vi_id,
                    ho_ten=ct.quan_nhan.ho_ten if ct.quan_nhan else de_xuat.don_vi.ten_don_vi,
                    cap_bac=ct.quan_nhan.cap_bac if ct.quan_nhan else None,
                    chuc_vu=ct.quan_nhan.chuc_vu if ct.quan_nhan else None,
                    doi_tuong=ct.doi_tuong,
                    loai_danh_hieu=ct.loai_danh_hieu,
                    nam_hoc=de_xuat.nam_hoc,
                    nguoi_duyet_id=current_user.id,
                    ngay_duyet=now,
                    ghi_chu=request.form.get('ghi_chu', '').strip() or None,
                )
                db.session.add(khen_thuong)

    db.session.commit()
    flash('Đã phê duyệt cuối cùng và lưu kết quả khen thưởng.', 'success')
    return redirect(url_for('admin.approval_tracking'))


@admin_bp.route('/tracking/<int:id>/reject', methods=['POST'])
@login_required
@admin_required
def reject_from_tracking(id):
    """Admin rejects a nomination from tracking detail page."""
    de_xuat = DeXuat.query.get_or_404(id)

    ly_do = request.form.get('ly_do', '').strip()
    if not ly_do:
        flash('Vui lòng nhập lý do từ chối.', 'danger')
        return redirect(url_for('admin.approval_tracking'))

    admin_pd = PheDuyet.query.filter_by(
        de_xuat_id=id, phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value
    ).first()

    if admin_pd:
        admin_pd.ket_qua = KetQuaDuyet.TU_CHOI.value
        admin_pd.nguoi_duyet_id = current_user.id
        admin_pd.ngay_duyet = datetime.utcnow()
        admin_pd.ly_do = ly_do

    de_xuat.trang_thai = TrangThaiDeXuat.TU_CHOI.value
    db.session.commit()
    flash('Đã từ chối đề xuất.', 'warning')
    return redirect(url_for('admin.approval_tracking'))


@admin_bp.route('/batch-final-approve', methods=['POST'])
@login_required
@admin_required
def batch_final_approve():
    """Batch final approval for multiple nominations at once."""
    data = request.get_json()
    ids = data.get('ids', [])
    ghi_chu = data.get('ghi_chu', '').strip() or None

    if not ids:
        return jsonify({'success': False, 'message': 'Không có đề xuất nào được chọn.'}), 400

    approved_count = 0
    now = datetime.utcnow()

    for dx_id in ids:
        de_xuat = DeXuat.query.get(dx_id)
        if not de_xuat or de_xuat.trang_thai != TrangThaiDeXuat.DA_DUYET.value:
            continue

        # Verify all 6 departments approved
        all_ok = True
        for dept_name in DEPT_NAMES:
            pd = PheDuyet.query.filter_by(
                de_xuat_id=dx_id, phong_duyet=dept_name
            ).first()
            if not pd or pd.ket_qua != KetQuaDuyet.DONG_Y.value:
                all_ok = False
                break

        if not all_ok:
            continue

        # Update admin PheDuyet record
        admin_pd = PheDuyet.query.filter_by(
            de_xuat_id=dx_id, phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value
        ).first()
        if admin_pd:
            admin_pd.ket_qua = KetQuaDuyet.DONG_Y.value
            admin_pd.nguoi_duyet_id = current_user.id
            admin_pd.ngay_duyet = now
            admin_pd.ghi_chu = ghi_chu

        de_xuat.trang_thai = TrangThaiDeXuat.PHE_DUYET_CUOI.value

        # Create KhenThuong records
        for ct in de_xuat.chi_tiets:
            all_approved = True
            for dept_name in DEPT_NAMES:
                pd = PheDuyet.query.filter_by(
                    de_xuat_id=dx_id, phong_duyet=dept_name
                ).first()
                if pd:
                    kq = KetQuaDuyetChiTiet.query.filter_by(
                        phe_duyet_id=pd.id, chi_tiet_id=ct.id
                    ).first()
                    if not kq or kq.ket_qua != KetQuaDuyet.DONG_Y.value:
                        all_approved = False
                        break
                else:
                    all_approved = False
                    break

            if all_approved:
                existing = KhenThuong.query.filter_by(
                    de_xuat_id=de_xuat.id, chi_tiet_id=ct.id
                ).first()
                if not existing:
                    khen_thuong = KhenThuong(
                        de_xuat_id=de_xuat.id,
                        chi_tiet_id=ct.id,
                        quan_nhan_id=ct.quan_nhan_id,
                        don_vi_id=de_xuat.don_vi_id,
                        ho_ten=ct.quan_nhan.ho_ten if ct.quan_nhan else de_xuat.don_vi.ten_don_vi,
                        cap_bac=ct.quan_nhan.cap_bac if ct.quan_nhan else None,
                        chuc_vu=ct.quan_nhan.chuc_vu if ct.quan_nhan else None,
                        doi_tuong=ct.doi_tuong,
                        loai_danh_hieu=ct.loai_danh_hieu,
                        nam_hoc=de_xuat.nam_hoc,
                        nguoi_duyet_id=current_user.id,
                        ngay_duyet=now,
                        ghi_chu=ghi_chu,
                    )
                    db.session.add(khen_thuong)

        approved_count += 1

    db.session.commit()
    return jsonify({
        'success': True,
        'message': f'Đã phê duyệt cuối {approved_count} đề xuất.',
        'approved_count': approved_count,
    })


@admin_bp.route('/revoke-final/<int:id>', methods=['POST'])
@login_required
@admin_required
def revoke_final_approval(id):
    """Revoke final approval - undo 'Phê duyệt cuối' back to 'Đã duyệt', delete KhenThuong records."""
    de_xuat = DeXuat.query.get_or_404(id)

    if de_xuat.trang_thai != TrangThaiDeXuat.PHE_DUYET_CUOI.value:
        flash('Đề xuất này chưa được phê duyệt cuối.', 'warning')
        return redirect(url_for('admin.reward_list'))

    # Delete all KhenThuong records for this nomination
    KhenThuong.query.filter_by(de_xuat_id=id).delete()

    # Reset admin PheDuyet back to pending
    admin_pd = PheDuyet.query.filter_by(
        de_xuat_id=id, phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value
    ).first()
    if admin_pd:
        admin_pd.ket_qua = KetQuaDuyet.CHO_DUYET.value
        admin_pd.nguoi_duyet_id = None
        admin_pd.ngay_duyet = None
        admin_pd.ly_do = None
        admin_pd.ghi_chu = None

    # Revert status back to 'Đã duyệt' (all 6 depts still approved)
    de_xuat.trang_thai = TrangThaiDeXuat.DA_DUYET.value

    db.session.commit()
    flash(f'Đã thu hồi phê duyệt cuối cho đề xuất của {de_xuat.don_vi.ten_don_vi}.', 'success')
    return redirect(url_for('admin.reward_list'))


@admin_bp.route('/reward-list')
@login_required
@admin_or_reward_viewer_required
def reward_list():
    """View all finalized awards (KhenThuong records)."""
    page = request.args.get('page', 1, type=int)
    nam_hoc_filter = request.args.get('nam_hoc', '')
    unit_filter = request.args.get('unit', '')
    danh_hieu_filter = request.args.get('danh_hieu', '')
    search_query = request.args.get('q', '').strip()

    query = KhenThuong.query

    if nam_hoc_filter:
        query = query.filter(KhenThuong.nam_hoc == nam_hoc_filter)
    if unit_filter:
        query = query.join(DonVi).filter(DonVi.ten_don_vi == unit_filter)
    if danh_hieu_filter:
        query = query.filter(KhenThuong.loai_danh_hieu == danh_hieu_filter)
    if search_query:
        query = query.filter(KhenThuong.ho_ten.ilike(f'%{search_query}%'))

    rewards = query.order_by(KhenThuong.ngay_duyet.desc())\
        .paginate(page=page, per_page=20, error_out=False)

    # Get filter options
    nam_hoc_list = db.session.query(KhenThuong.nam_hoc).distinct()\
        .order_by(KhenThuong.nam_hoc.desc()).all()
    nam_hoc_list = [n[0] for n in nam_hoc_list]

    unit_names = db.session.query(DonVi.ten_don_vi).join(
        KhenThuong, KhenThuong.don_vi_id == DonVi.id
    ).distinct().order_by(DonVi.ten_don_vi).all()
    unit_names = [u[0] for u in unit_names]

    danh_hieu_list = db.session.query(KhenThuong.loai_danh_hieu).distinct().all()
    danh_hieu_list = [d[0] for d in danh_hieu_list]

    # Summary stats
    total_rewards = KhenThuong.query.count()
    stats_by_danh_hieu = {}
    for dh in danh_hieu_list:
        stats_by_danh_hieu[dh] = KhenThuong.query.filter_by(loai_danh_hieu=dh).count()

    pending_final_nominations = []
    if current_user.is_admin:
        pending_final_nominations = []
        nominations_waiting = DeXuat.query.filter_by(
            trang_thai=TrangThaiDeXuat.DA_DUYET.value
        ).order_by(DeXuat.ngay_gui.desc()).all()

        for dx in nominations_waiting:
            for ct in dx.chi_tiets:
                if KhenThuong.query.filter_by(chi_tiet_id=ct.id).first():
                    continue
                all_dept_ok = True
                for dept_name in DEPT_NAMES:
                    pd = PheDuyet.query.filter_by(de_xuat_id=dx.id, phong_duyet=dept_name).first()
                    if not pd:
                        all_dept_ok = False
                        break
                    kq = KetQuaDuyetChiTiet.query.filter_by(phe_duyet_id=pd.id, chi_tiet_id=ct.id).first()
                    if not kq or kq.ket_qua != KetQuaDuyet.DONG_Y.value:
                        all_dept_ok = False
                        break
                if all_dept_ok:
                    pending_final_nominations.append({'dx': dx, 'ct': ct})

    # Statistics: personnel with >=3 CSTD (consecutive / non-consecutive)
    cstd_rows = db.session.query(KhenThuong.quan_nhan_id, KhenThuong.nam_hoc).filter(
        KhenThuong.loai_danh_hieu == LoaiDanhHieu.CHIEN_SI_THI_DUA.value,
        KhenThuong.quan_nhan_id.isnot(None)
    ).all()

    by_person = {}
    for qn_id, nam_hoc in cstd_rows:
        by_person.setdefault(qn_id, set()).add(nam_hoc)

    def _nam_hoc_start(nh):
        try:
            return int(str(nh).split('-')[0])
        except Exception:
            return 0

    cstd_non_consecutive = []
    cstd_consecutive = []
    for qn_id, years in by_person.items():
        if len(years) < 3:
            continue
        qn = QuanNhan.query.get(qn_id)
        if not qn:
            continue
        years_sorted = sorted(list(years), key=_nam_hoc_start)
        cstd_non_consecutive.append({'qn': qn, 'years': years_sorted, 'count': len(years_sorted)})

        starts = [_nam_hoc_start(y) for y in years_sorted if _nam_hoc_start(y) > 0]
        starts = sorted(starts)
        streak = 1
        max_streak = 1
        for i in range(1, len(starts)):
            if starts[i] == starts[i - 1] + 1:
                streak += 1
                max_streak = max(max_streak, streak)
            else:
                streak = 1
        if max_streak >= 3:
            cstd_consecutive.append({'qn': qn, 'years': years_sorted, 'max_streak': max_streak})

    return render_template('admin/reward_list.html',
                           rewards=rewards,
                           nam_hoc_filter=nam_hoc_filter,
                           unit_filter=unit_filter,
                           danh_hieu_filter=danh_hieu_filter,
                           search_query=search_query,
                           nam_hoc_list=nam_hoc_list,
                           unit_names=unit_names,
                           danh_hieu_list=danh_hieu_list,
                           total_rewards=total_rewards,
                           stats_by_danh_hieu=stats_by_danh_hieu,
                           can_admin_action=current_user.is_admin,
                           pending_final_nominations=pending_final_nominations,
                           cstd_non_consecutive=cstd_non_consecutive,
                           cstd_consecutive=cstd_consecutive)


@admin_bp.route('/reward-list/export')
@login_required
@admin_or_reward_viewer_required
def export_reward_list():
    """Export danh sách khen thưởng ra file Excel với đầy đủ thông tin."""
    nam_hoc_filter = request.args.get('nam_hoc', '')
    unit_filter = request.args.get('unit', '')
    danh_hieu_filter = request.args.get('danh_hieu', '')
    search_query = request.args.get('q', '').strip()

    query = KhenThuong.query

    if nam_hoc_filter:
        query = query.filter(KhenThuong.nam_hoc == nam_hoc_filter)
    if unit_filter:
        query = query.join(DonVi).filter(DonVi.ten_don_vi == unit_filter)
    if danh_hieu_filter:
        query = query.filter(KhenThuong.loai_danh_hieu == danh_hieu_filter)
    if search_query:
        query = query.filter(KhenThuong.ho_ten.ilike(f'%{search_query}%'))

    rewards = query.order_by(KhenThuong.ngay_duyet.desc()).all()

    # Create Excel workbook
    wb = Workbook()
    ws = wb.active
    ws.title = 'Danh sách khen thưởng'

    # Styles
    header_font = Font(name='Times New Roman', bold=True, size=13)
    sub_header_font = Font(name='Times New Roman', bold=True, size=11)
    col_header_font = Font(name='Times New Roman', bold=True, size=10)
    col_header_fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
    col_header_font_white = Font(name='Times New Roman', bold=True, size=10, color='FFFFFF')
    data_font = Font(name='Times New Roman', size=10)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin'),
    )
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # Title rows
    ws.merge_cells('A1:S1')
    cell_title = ws['A1']
    cell_title.value = 'TRƯỜNG SĨ QUAN CHÍNH TRỊ'
    cell_title.font = Font(name='Times New Roman', bold=True, size=13)
    cell_title.alignment = center_align

    ws.merge_cells('A2:S2')
    cell_sub = ws['A2']
    cell_sub.value = 'DANH SÁCH KHEN THƯỞNG'
    cell_sub.font = Font(name='Times New Roman', bold=True, size=14)
    cell_sub.alignment = center_align

    # Filter info row
    filter_parts = []
    if nam_hoc_filter:
        filter_parts.append(f'Năm học: {nam_hoc_filter}')
    if unit_filter:
        filter_parts.append(f'Đơn vị: {unit_filter}')
    if danh_hieu_filter:
        filter_parts.append(f'Danh hiệu: {danh_hieu_filter}')
    if search_query:
        filter_parts.append(f'Tìm kiếm: {search_query}')

    ws.merge_cells('A3:S3')
    cell_filter = ws['A3']
    cell_filter.value = ' | '.join(filter_parts) if filter_parts else 'Tất cả dữ liệu'
    cell_filter.font = Font(name='Times New Roman', italic=True, size=10)
    cell_filter.alignment = center_align

    # Column headers (row 5)
    headers = [
        ('STT', 6),
        ('Họ và tên', 22),
        ('Cấp bậc', 14),
        ('Chức vụ', 20),
        ('Đối tượng', 18),
        ('Đơn vị', 30),
        ('Danh hiệu thi đua', 22),
        ('Năm học', 12),
        ('Hoàn thành NV', 16),
        ('Phiếu tín nhiệm', 14),
        ('KT Chính trị', 14),
        ('KT Điều lệnh', 14),
        ('Kỹ năng số', 14),
        ('ĐHQS', 14),
        ('Bắn súng', 12),
        ('Thể lực', 12),
        ('Điểm NCKH', 12),
        ('Nội dung NCKH', 30),
        ('Ngày phê duyệt', 16),
    ]

    for col_idx, (header_text, width) in enumerate(headers, 1):
        cell = ws.cell(row=5, column=col_idx, value=header_text)
        cell.font = col_header_font_white
        cell.fill = col_header_fill
        cell.alignment = center_align
        cell.border = thin_border
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # Data rows
    for row_idx, kt in enumerate(rewards, 1):
        row_num = row_idx + 5  # data starts at row 6

        # Load original DeXuatChiTiet for criteria fields
        ct = DeXuatChiTiet.query.get(kt.chi_tiet_id) if kt.chi_tiet_id else None

        row_data = [
            row_idx,                                                    # STT
            kt.ho_ten,                                                  # Họ tên
            kt.cap_bac or '',                                           # Cấp bậc
            kt.chuc_vu or '',                                           # Chức vụ
            kt.doi_tuong or '',                                         # Đối tượng
            kt.don_vi.ten_don_vi if kt.don_vi else '',                 # Đơn vị
            kt.loai_danh_hieu or '',                                    # Danh hiệu
            kt.nam_hoc or '',                                           # Năm học
            (ct.muc_do_hoan_thanh or '') if ct else '',                # HTNV
            (ct.phieu_tin_nhiem or '') if ct else '',                   # Phiếu tín nhiệm
            (ct.kiem_tra_chinh_tri or '') if ct else '',                # KT CT
            (ct.kiem_tra_dieu_lenh or '') if ct else '',                # KT ĐL
            (ct.kiem_tra_tin_hoc or '') if ct else '',                  # Kỹ năng số
            (ct.dia_ly_quan_su or '') if ct else '',                    # ĐHQS
            (ct.ban_sung or '') if ct else '',                          # Bắn súng
            (ct.the_luc or '') if ct else '',                           # Thể lực
            (ct.diem_nckh or '') if ct else '',                         # Điểm NCKH
            (ct.nckh_noi_dung or '') if ct else '',                    # Nội dung NCKH
            kt.ngay_duyet.strftime('%d/%m/%Y') if kt.ngay_duyet else '',  # Ngày duyệt
        ]

        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_num, column=col_idx, value=value)
            cell.font = data_font
            cell.border = thin_border
            if col_idx == 1:  # STT
                cell.alignment = center_align
            elif col_idx in (8, 10, 11, 12, 13, 14, 15, 16, 17, 19):  # center-aligned cols
                cell.alignment = center_align
            else:
                cell.alignment = left_align

    # Summary row
    summary_row = len(rewards) + 6
    ws.merge_cells(f'A{summary_row}:F{summary_row}')
    cell_sum = ws.cell(row=summary_row, column=1, value=f'Tổng cộng: {len(rewards)} cá nhân')
    cell_sum.font = Font(name='Times New Roman', bold=True, size=10)
    cell_sum.alignment = left_align

    # Signature section
    sig_row = summary_row + 2
    ws.merge_cells(f'N{sig_row}:S{sig_row}')
    cell_date = ws.cell(row=sig_row, column=14,
                        value=f'Ngày {datetime.now().day} tháng {datetime.now().month} năm {datetime.now().year}')
    cell_date.font = Font(name='Times New Roman', italic=True, size=10)
    cell_date.alignment = center_align

    ws.merge_cells(f'N{sig_row+1}:S{sig_row+1}')
    cell_signer = ws.cell(row=sig_row + 1, column=14, value='CƠ QUAN TUYÊN HUẤN')
    cell_signer.font = Font(name='Times New Roman', bold=True, size=11)
    cell_signer.alignment = center_align

    # Set print area and page setup
    ws.sheet_properties.pageSetUpPr = None
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1

    # Freeze panes
    ws.freeze_panes = 'A6'

    # Write to BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    # Generate filename
    filename_parts = ['DanhSachKhenThuong']
    if nam_hoc_filter:
        filename_parts.append(nam_hoc_filter.replace('-', '_'))
    filename_parts.append(datetime.now().strftime('%d%m%Y'))
    filename = '_'.join(filename_parts) + '.xlsx'

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename,
    )


@admin_bp.route('/reward-detail/<int:kt_id>')
@login_required
@admin_or_reward_viewer_required
def reward_detail(kt_id):
    """View detailed info for one KhenThuong record (individual award)."""
    kt = KhenThuong.query.get_or_404(kt_id)

    # Load the original DeXuatChiTiet for full criteria
    ct = DeXuatChiTiet.query.get(kt.chi_tiet_id) if kt.chi_tiet_id else None
    de_xuat = DeXuat.query.get(kt.de_xuat_id) if kt.de_xuat_id else None

    # Build per-department results for this individual
    dept_item_results = {}
    if de_xuat:
        for pd in de_xuat.phe_duyets:
            kq = None
            if ct:
                kq = KetQuaDuyetChiTiet.query.filter_by(
                    phe_duyet_id=pd.id, chi_tiet_id=ct.id
                ).first()
            dept_item_results[pd.phong_duyet] = {
                'phe_duyet': pd,
                'item_result': kq,
            }

    # All criteria field labels
    all_field_labels = {
        'muc_do_hoan_thanh': 'Hoàn thành nhiệm vụ',
        'phieu_tin_nhiem': 'Phiếu tín nhiệm',
        'kiem_tra_dieu_lenh': 'Kiểm tra điều lệnh',
        'ban_sung': 'Bắn súng',
        'the_luc': 'Thể lực',
        'kiem_tra_chinh_tri': 'Kiểm tra chính trị',
        'kiem_tra_tin_hoc': 'Kỹ năng số',
        'dia_ly_quan_su': 'Địa hình quân sự',
        'danh_hieu_gv_gioi': 'Danh hiệu GV giỏi',
        'dinh_muc_giang_day': 'Định mức giảng dạy',
        'ket_qua_kiem_tra_giang': 'Kết quả kiểm tra giảng',
        'thoi_gian_lao_dong_kh': 'Thời gian LĐ khoa học',
        'tien_do_pgs': 'Tiến độ PGS',
        'danh_hieu_hv_gioi': 'Danh hiệu HV giỏi',
        'diem_tong_ket': 'Điểm tổng kết',
        'ket_qua_thuc_hanh': 'Kết quả thực hành',
        'ket_qua_doan_the': 'Kết quả đoàn thể',
        'chu_tri_don_vi_danh_hieu': 'Chủ trì ĐV danh hiệu',
        'diem_nckh': 'Điểm NCKH',
        'nckh_noi_dung': 'Nội dung NCKH',
        'nckh_minh_chung': 'Minh chứng NCKH',
        'thanh_tich_ca_nhan_khac': 'Thành tích cá nhân khác',
    }

    all_fields = [
        'muc_do_hoan_thanh', 'phieu_tin_nhiem',
        'kiem_tra_chinh_tri', 'kiem_tra_dieu_lenh', 'kiem_tra_tin_hoc',
        'dia_ly_quan_su', 'ban_sung', 'the_luc',
        'ket_qua_doan_the', 'chu_tri_don_vi_danh_hieu',
        'danh_hieu_gv_gioi', 'dinh_muc_giang_day', 'ket_qua_kiem_tra_giang',
        'tien_do_pgs', 'thoi_gian_lao_dong_kh',
        'danh_hieu_hv_gioi', 'diem_tong_ket', 'ket_qua_thuc_hanh',
        'diem_nckh', 'nckh_noi_dung', 'nckh_minh_chung',
        'thanh_tich_ca_nhan_khac',
    ]

    return render_template('admin/reward_detail.html',
                           kt=kt,
                           ct=ct,
                           de_xuat=de_xuat,
                           dept_item_results=dept_item_results,
                           dept_names=DEPT_NAMES,
                           all_field_labels=all_field_labels,
                           all_fields=all_fields)


# ===== Legacy routes - redirect to new locations =====

@admin_bp.route('/final-review')
@login_required
@admin_required
def final_review_list():
    """Redirect to tracking page filtered for awaiting final approval."""
    return redirect(url_for('admin.approval_tracking', status='Đã duyệt'))


@admin_bp.route('/final-review/<int:id>')
@login_required
@admin_required
def final_review_detail(id):
    """Redirect to tracking page (legacy route)."""
    return redirect(url_for('admin.approval_tracking'))


@admin_bp.route('/final-review/<int:id>/approve', methods=['POST'])
@login_required
@admin_required
def final_approve(id):
    """Legacy final approve - redirect to new endpoint."""
    return redirect(url_for('admin.final_approve_from_tracking', id=id), code=307)


@admin_bp.route('/final-review/<int:id>/reject', methods=['POST'])
@login_required
@admin_required
def final_reject(id):
    """Legacy final reject - redirect to new endpoint."""
    return redirect(url_for('admin.reject_from_tracking', id=id), code=307)


# ===== User/Unit management (unchanged) =====

@admin_bp.route('/users')
@login_required
@admin_required
def manage_users():
    users = User.query.order_by(User.role, User.username).all()
    units = DonVi.query.filter_by(is_active=True).order_by(DonVi.thu_tu).all()
    roles = [(r.value, ROLE_DISPLAY.get(r, r.name)) for r in Role]
    return render_template('admin/manage_users.html',
                           users=users, units=units, roles=roles)


@admin_bp.route('/users/create', methods=['POST'])
@login_required
@admin_required
def create_user():
    username = request.form.get('username', '').strip()
    password = request.form.get('password', '').strip()
    ho_ten = request.form.get('ho_ten', '').strip()
    role_val = request.form.get('role', '').strip()
    don_vi_id = request.form.get('don_vi_id', type=int)

    if not username or not password or not ho_ten or not role_val:
        flash('Vui lòng điền đầy đủ thông tin.', 'danger')
        return redirect(url_for('admin.manage_users'))

    if User.query.filter_by(username=username).first():
        flash('Tên đăng nhập đã tồn tại.', 'danger')
        return redirect(url_for('admin.manage_users'))

    try:
        role = Role(role_val)
    except ValueError:
        flash('Vai trò không hợp lệ.', 'danger')
        return redirect(url_for('admin.manage_users'))

    user = User(
        username=username,
        ho_ten=ho_ten,
        role=role,
        don_vi_id=don_vi_id if role == Role.UNIT_USER else None,
    )
    user.set_password(password)
    db.session.add(user)
    db.session.commit()
    flash(f'Đã tạo tài khoản: {username}', 'success')
    return redirect(url_for('admin.manage_users'))


@admin_bp.route('/users/<int:id>/toggle', methods=['POST'])
@login_required
@admin_required
def toggle_user(id):
    user = User.query.get_or_404(id)
    if user.id == current_user.id:
        flash('Không thể tự vô hiệu hóa tài khoản của mình.', 'danger')
        return redirect(url_for('admin.manage_users'))

    user.is_active_account = not user.is_active_account
    db.session.commit()
    status = 'kích hoạt' if user.is_active_account else 'vô hiệu hóa'
    flash(f'Đã {status} tài khoản: {user.username}', 'success')
    return redirect(url_for('admin.manage_users'))


@admin_bp.route('/users/<int:id>/reset-password', methods=['POST'])
@login_required
@admin_required
def reset_password(id):
    user = User.query.get_or_404(id)
    new_pw = request.form.get('new_password', '').strip()
    if not new_pw:
        flash('Mật khẩu mới không được để trống.', 'danger')
        return redirect(url_for('admin.manage_users'))

    user.set_password(new_pw)
    db.session.commit()
    flash(f'Đã đặt lại mật khẩu cho: {user.username}', 'success')
    return redirect(url_for('admin.manage_users'))


@admin_bp.route('/units')
@login_required
@admin_required
def manage_units():
    units_by_type = {}
    for loai in LoaiDonVi:
        units_by_type[loai] = DonVi.query.filter_by(loai_don_vi=loai)\
            .order_by(DonVi.thu_tu).all()

    loai_don_vi_list = [(l.value, l.name) for l in LoaiDonVi]
    return render_template('admin/manage_units.html',
                           units_by_type=units_by_type,
                           loai_don_vi_list=loai_don_vi_list)


@admin_bp.route('/units/create', methods=['POST'])
@login_required
@admin_required
def create_unit():
    ma_don_vi = request.form.get('ma_don_vi', '').strip()
    ten_don_vi = request.form.get('ten_don_vi', '').strip()
    loai_don_vi_val = request.form.get('loai_don_vi', '').strip()
    thu_tu = request.form.get('thu_tu', 0, type=int)

    if not ma_don_vi or not ten_don_vi or not loai_don_vi_val:
        flash('Vui lòng điền đầy đủ thông tin.', 'danger')
        return redirect(url_for('admin.manage_units'))

    # Check unique ma_don_vi
    if DonVi.query.filter_by(ma_don_vi=ma_don_vi).first():
        flash(f'Mã đơn vị "{ma_don_vi}" đã tồn tại.', 'danger')
        return redirect(url_for('admin.manage_units'))

    try:
        loai = LoaiDonVi(loai_don_vi_val)
    except ValueError:
        flash('Loại đơn vị không hợp lệ.', 'danger')
        return redirect(url_for('admin.manage_units'))

    unit = DonVi(
        ma_don_vi=ma_don_vi,
        ten_don_vi=ten_don_vi,
        loai_don_vi=loai,
        thu_tu=thu_tu,
        is_active=True,
    )
    db.session.add(unit)
    db.session.commit()
    flash(f'Đã tạo đơn vị: {ten_don_vi}', 'success')
    return redirect(url_for('admin.manage_units'))


@admin_bp.route('/units/<int:id>/edit', methods=['POST'])
@login_required
@admin_required
def edit_unit(id):
    unit = DonVi.query.get_or_404(id)
    ma_don_vi = request.form.get('ma_don_vi', '').strip()
    ten_don_vi = request.form.get('ten_don_vi', '').strip()
    loai_don_vi_val = request.form.get('loai_don_vi', '').strip()
    thu_tu = request.form.get('thu_tu', 0, type=int)

    if not ma_don_vi or not ten_don_vi or not loai_don_vi_val:
        flash('Vui lòng điền đầy đủ thông tin.', 'danger')
        return redirect(url_for('admin.manage_units'))

    # Check unique ma_don_vi (exclude current unit)
    existing = DonVi.query.filter_by(ma_don_vi=ma_don_vi).first()
    if existing and existing.id != unit.id:
        flash(f'Mã đơn vị "{ma_don_vi}" đã được sử dụng.', 'danger')
        return redirect(url_for('admin.manage_units'))

    try:
        loai = LoaiDonVi(loai_don_vi_val)
    except ValueError:
        flash('Loại đơn vị không hợp lệ.', 'danger')
        return redirect(url_for('admin.manage_units'))

    unit.ma_don_vi = ma_don_vi
    unit.ten_don_vi = ten_don_vi
    unit.loai_don_vi = loai
    unit.thu_tu = thu_tu
    db.session.commit()
    flash(f'Đã cập nhật đơn vị: {ten_don_vi}', 'success')
    return redirect(url_for('admin.manage_units'))


@admin_bp.route('/units/<int:id>/toggle', methods=['POST'])
@login_required
@admin_required
def toggle_unit(id):
    unit = DonVi.query.get_or_404(id)
    unit.is_active = not unit.is_active
    db.session.commit()
    status = 'kích hoạt' if unit.is_active else 'ngừng hoạt động'
    flash(f'Đã {status} đơn vị: {unit.ten_don_vi}', 'success')
    return redirect(url_for('admin.manage_units'))


@admin_bp.route('/report')
@login_required
@admin_required
def report_summary():
    nam_hoc_filter = request.args.get('nam_hoc', '')

    # Get available nam_hoc values
    nam_hoc_list = db.session.query(DeXuat.nam_hoc).distinct()\
        .order_by(DeXuat.nam_hoc.desc()).all()
    nam_hoc_list = [n[0] for n in nam_hoc_list if n[0]]

    # Base queries (optionally filtered by nam_hoc)
    dx_query = DeXuat.query
    kt_query = KhenThuong.query
    if nam_hoc_filter:
        dx_query = dx_query.filter(DeXuat.nam_hoc == nam_hoc_filter)
        kt_query = kt_query.filter(KhenThuong.nam_hoc == nam_hoc_filter)

    # Overall stats
    total_personnel = QuanNhan.query.filter_by(is_active=True).count()
    total_nominations = dx_query.count()
    total_approved = dx_query.filter(
        DeXuat.trang_thai == TrangThaiDeXuat.PHE_DUYET_CUOI.value
    ).count()
    total_rewards = kt_query.count()

    # Stats by nomination status
    status_stats = {}
    for status in TrangThaiDeXuat:
        if status == TrangThaiDeXuat.NHAP:
            continue
        count = dx_query.filter(DeXuat.trang_thai == status.value).count()
        status_stats[status.value] = count

    # Stats by danh_hieu (from KhenThuong)
    danh_hieu_stats = {}
    for dh in LoaiDanhHieu:
        count = kt_query.filter(KhenThuong.loai_danh_hieu == dh.value).count()
        danh_hieu_stats[dh.value] = count

    # Stats by unit type (from KhenThuong)
    loai_dv_stats = {}
    for ldv in LoaiDonVi:
        count = kt_query.join(DonVi, KhenThuong.don_vi_id == DonVi.id)\
            .filter(DonVi.loai_don_vi == ldv).count()
        loai_dv_stats[ldv.value] = count

    # Per-unit stats
    units = DonVi.query.filter_by(is_active=True).order_by(DonVi.thu_tu).all()
    unit_stats = []
    for unit in units:
        personnel_count = QuanNhan.query.filter_by(don_vi_id=unit.id, is_active=True).count()

        nom_q = DeXuat.query.filter_by(don_vi_id=unit.id)
        kt_q = KhenThuong.query.filter_by(don_vi_id=unit.id)
        if nam_hoc_filter:
            nom_q = nom_q.filter(DeXuat.nam_hoc == nam_hoc_filter)
            kt_q = kt_q.filter(KhenThuong.nam_hoc == nam_hoc_filter)

        nomination_count = nom_q.count()
        approved_count = nom_q.filter(
            DeXuat.trang_thai == TrangThaiDeXuat.PHE_DUYET_CUOI.value
        ).count()
        reward_count = kt_q.count()

        # Per-danh-hieu breakdown for this unit
        unit_dh = {}
        for dh in LoaiDanhHieu:
            unit_dh[dh.value] = kt_q.filter(KhenThuong.loai_danh_hieu == dh.value).count()

        unit_stats.append({
            'unit': unit,
            'personnel': personnel_count,
            'nominations': nomination_count,
            'approved': approved_count,
            'rewards': reward_count,
            'by_danh_hieu': unit_dh,
        })

    # Department approval stats (how many individuals each dept has reviewed)
    dept_stats = {}
    for dept_name in DEPT_NAMES:
        total_reviewed = db.session.query(KetQuaDuyetChiTiet).join(
            PheDuyet, KetQuaDuyetChiTiet.phe_duyet_id == PheDuyet.id
        ).filter(PheDuyet.phong_duyet == dept_name)

        if nam_hoc_filter:
            total_reviewed = total_reviewed.join(
                DeXuat, PheDuyet.de_xuat_id == DeXuat.id
            ).filter(DeXuat.nam_hoc == nam_hoc_filter)

        total_count = total_reviewed.count()
        approved_count = total_reviewed.filter(
            KetQuaDuyetChiTiet.ket_qua == KetQuaDuyet.DONG_Y.value
        ).count()
        rejected_count = total_reviewed.filter(
            KetQuaDuyetChiTiet.ket_qua == KetQuaDuyet.TU_CHOI.value
        ).count()

        dept_stats[dept_name] = {
            'total': total_count,
            'approved': approved_count,
            'rejected': rejected_count,
        }

    # Chart data: rewards per unit (top units with rewards)
    chart_unit_names = []
    chart_unit_rewards = []
    for s in unit_stats:
        if s['rewards'] > 0:
            chart_unit_names.append(s['unit'].ma_don_vi)
            chart_unit_rewards.append(s['rewards'])

    return render_template('admin/report.html',
                           unit_stats=unit_stats,
                           total_personnel=total_personnel,
                           total_nominations=total_nominations,
                           total_approved=total_approved,
                           total_rewards=total_rewards,
                           status_stats=status_stats,
                           danh_hieu_stats=danh_hieu_stats,
                           loai_dv_stats=loai_dv_stats,
                           dept_stats=dept_stats,
                           dept_names=DEPT_NAMES,
                           nam_hoc_filter=nam_hoc_filter,
                           nam_hoc_list=nam_hoc_list,
                            chart_unit_names=chart_unit_names,
                           chart_unit_rewards=chart_unit_rewards)


# ------------------------------------------------------------------
# Admin: View all personnel across all units
# ------------------------------------------------------------------
@admin_bp.route('/personnel')
@login_required
@admin_required
def all_personnel():
    search = request.args.get('search', '').strip()
    don_vi_id = request.args.get('don_vi_id', '', type=str)
    doi_tuong = request.args.get('doi_tuong', '').strip()
    page = request.args.get('page', 1, type=int)

    query = QuanNhan.query.filter_by(is_active=True).join(DonVi)

    if search:
        query = query.filter(QuanNhan.ho_ten.ilike(f'%{search}%'))
    if don_vi_id:
        query = query.filter(QuanNhan.don_vi_id == int(don_vi_id))
    if doi_tuong:
        query = query.filter(QuanNhan.doi_tuong == doi_tuong)

    query = query.order_by(DonVi.thu_tu, DonVi.ten_don_vi, QuanNhan.ho_ten)
    personnel = query.paginate(page=page, per_page=30, error_out=False)

    units = DonVi.query.filter_by(is_active=True).order_by(DonVi.thu_tu, DonVi.ten_don_vi).all()
    doi_tuong_list = [e.value for e in DoiTuong]

    return render_template('admin/all_personnel.html',
                           personnel=personnel,
                           search=search,
                           don_vi_id=don_vi_id,
                           doi_tuong_filter=doi_tuong,
                           units=units,
                           doi_tuong_list=doi_tuong_list)


# ------------------------------------------------------------------
# Admin: View personnel detail (any unit)
# ------------------------------------------------------------------
@admin_bp.route('/personnel/<int:id>')
@login_required
@admin_required
def admin_personnel_detail(id):
    qn = QuanNhan.query.get_or_404(id)
    return render_template('admin/personnel_detail.html', qn=qn,
                           loai_chung_chi_list=[e.value for e in LoaiChungChi])


# ------------------------------------------------------------------
# Admin: Edit personnel (any unit)
# ------------------------------------------------------------------
@admin_bp.route('/personnel/<int:id>/edit', methods=['GET', 'POST'])
@login_required
@admin_required
def admin_personnel_edit(id):
    qn = QuanNhan.query.get_or_404(id)

    if request.method == 'POST':
        qn.ho_ten = request.form.get('ho_ten', '').strip()
        qn.cap_bac = request.form.get('cap_bac', '').strip() or None
        qn.chuc_danh = request.form.get('chuc_danh', '').strip() or None
        qn.chuc_vu = request.form.get('chuc_vu', '').strip() or None
        qn.can_cuoc_cong_dan = request.form.get('can_cuoc_cong_dan', '').strip() or None
        qn.doi_tuong = request.form.get('doi_tuong', '').strip() or None
        qn.hoc_ham = request.form.get('hoc_ham', 'Không').strip()
        qn.hoc_vi = request.form.get('hoc_vi', 'Không').strip()
        qn.trinh_do_hoc_van = request.form.get('trinh_do_hoc_van', '').strip() or None
        qn.ngoai_ngu = request.form.get('ngoai_ngu', '').strip() or None
        qn.ngay_nhap_ngu = request.form.get('ngay_nhap_ngu', '').strip() or None
        qn.la_chi_huy = 'la_chi_huy' in request.form
        qn.la_bi_thu = 'la_bi_thu' in request.form

        ns_str = request.form.get('ngay_sinh', '').strip()
        if ns_str:
            try:
                qn.ngay_sinh = datetime.strptime(ns_str, '%Y-%m-%d').date()
            except ValueError:
                pass
        else:
            qn.ngay_sinh = None

        if not qn.ho_ten:
            flash('Họ tên không được để trống.', 'danger')
        else:
            db.session.commit()
            flash('Đã cập nhật thông tin.', 'success')
            return redirect(url_for('admin.admin_personnel_detail', id=qn.id))

    return render_template('admin/personnel_edit.html', qn=qn,
                           cap_bac_list=[x.ten for x in CapBacOption.query.filter_by(is_active=True).order_by(CapBacOption.thu_tu, CapBacOption.ten).all()] or [e.value for e in CapBac],
                           hoc_ham_list=[e.value for e in HocHam],
                           hoc_vi_list=[e.value for e in HocVi],
                           doi_tuong_list=[e.value for e in DoiTuong],
                           chuc_vu_options=ChucVuOption.query.filter_by(is_active=True).order_by(ChucVuOption.thu_tu, ChucVuOption.ten).all())


@admin_bp.route('/cap-bac')
@login_required
@admin_required
def manage_cap_bac():
    items = CapBacOption.query.order_by(CapBacOption.thu_tu, CapBacOption.ten).all()
    return render_template('admin/manage_cap_bac.html', items=items)


@admin_bp.route('/cap-bac/create', methods=['POST'])
@login_required
@admin_required
def create_cap_bac():
    ten = request.form.get('ten', '').strip()
    thu_tu = request.form.get('thu_tu', 0, type=int)
    if not ten:
        flash('Tên cấp bậc không được để trống.', 'danger')
        return redirect(url_for('admin.manage_cap_bac'))
    if CapBacOption.query.filter_by(ten=ten).first():
        flash('Cấp bậc đã tồn tại.', 'warning')
        return redirect(url_for('admin.manage_cap_bac'))
    db.session.add(CapBacOption(ten=ten, thu_tu=thu_tu, is_active=True))
    db.session.commit()
    flash('Đã thêm cấp bậc.', 'success')
    return redirect(url_for('admin.manage_cap_bac'))


@admin_bp.route('/cap-bac/<int:id>/edit', methods=['POST'])
@login_required
@admin_required
def edit_cap_bac(id):
    item = CapBacOption.query.get_or_404(id)
    ten = request.form.get('ten', '').strip()
    thu_tu = request.form.get('thu_tu', 0, type=int)
    if not ten:
        flash('Tên cấp bậc không được để trống.', 'danger')
        return redirect(url_for('admin.manage_cap_bac'))
    dup = CapBacOption.query.filter(CapBacOption.ten == ten, CapBacOption.id != id).first()
    if dup:
        flash('Tên cấp bậc đã tồn tại.', 'warning')
        return redirect(url_for('admin.manage_cap_bac'))
    item.ten = ten
    item.thu_tu = thu_tu
    db.session.commit()
    flash('Đã cập nhật cấp bậc.', 'success')
    return redirect(url_for('admin.manage_cap_bac'))


@admin_bp.route('/cap-bac/<int:id>/toggle', methods=['POST'])
@login_required
@admin_required
def toggle_cap_bac(id):
    item = CapBacOption.query.get_or_404(id)
    item.is_active = not item.is_active
    db.session.commit()
    flash('Đã cập nhật trạng thái cấp bậc.', 'success')
    return redirect(url_for('admin.manage_cap_bac'))


@admin_bp.route('/cap-bac/<int:id>/delete', methods=['POST'])
@login_required
@admin_required
def delete_cap_bac(id):
    item = CapBacOption.query.get_or_404(id)
    db.session.delete(item)
    db.session.commit()
    flash('Đã xóa cấp bậc.', 'success')
    return redirect(url_for('admin.manage_cap_bac'))


@admin_bp.route('/chuc-vu')
@login_required
@admin_required
def manage_chuc_vu():
    items = ChucVuOption.query.order_by(ChucVuOption.thu_tu, ChucVuOption.ten).all()
    return render_template('admin/manage_chuc_vu.html', items=items)


@admin_bp.route('/chuc-vu/create', methods=['POST'])
@login_required
@admin_required
def create_chuc_vu():
    ten = request.form.get('ten', '').strip()
    thu_tu = request.form.get('thu_tu', 0, type=int)
    if not ten:
        flash('Tên chức vụ không được để trống.', 'danger')
        return redirect(url_for('admin.manage_chuc_vu'))
    if ChucVuOption.query.filter_by(ten=ten).first():
        flash('Chức vụ đã tồn tại.', 'warning')
        return redirect(url_for('admin.manage_chuc_vu'))
    db.session.add(ChucVuOption(ten=ten, thu_tu=thu_tu, is_active=True))
    db.session.commit()
    flash('Đã thêm chức vụ.', 'success')
    return redirect(url_for('admin.manage_chuc_vu'))


@admin_bp.route('/chuc-vu/<int:id>/edit', methods=['POST'])
@login_required
@admin_required
def edit_chuc_vu(id):
    item = ChucVuOption.query.get_or_404(id)
    ten = request.form.get('ten', '').strip()
    thu_tu = request.form.get('thu_tu', 0, type=int)
    if not ten:
        flash('Tên chức vụ không được để trống.', 'danger')
        return redirect(url_for('admin.manage_chuc_vu'))
    dup = ChucVuOption.query.filter(ChucVuOption.ten == ten, ChucVuOption.id != id).first()
    if dup:
        flash('Tên chức vụ đã tồn tại.', 'warning')
        return redirect(url_for('admin.manage_chuc_vu'))
    item.ten = ten
    item.thu_tu = thu_tu
    db.session.commit()
    flash('Đã cập nhật chức vụ.', 'success')
    return redirect(url_for('admin.manage_chuc_vu'))


@admin_bp.route('/chuc-vu/<int:id>/toggle', methods=['POST'])
@login_required
@admin_required
def toggle_chuc_vu(id):
    item = ChucVuOption.query.get_or_404(id)
    item.is_active = not item.is_active
    db.session.commit()
    flash('Đã cập nhật trạng thái chức vụ.', 'success')
    return redirect(url_for('admin.manage_chuc_vu'))


@admin_bp.route('/chuc-vu/<int:id>/delete', methods=['POST'])
@login_required
@admin_required
def delete_chuc_vu(id):
    item = ChucVuOption.query.get_or_404(id)
    db.session.delete(item)
    db.session.commit()
    flash('Đã xóa chức vụ.', 'success')
    return redirect(url_for('admin.manage_chuc_vu'))


# ------------------------------------------------------------------
# Admin: Add certificate for any personnel
# ------------------------------------------------------------------
@admin_bp.route('/personnel/<int:id>/certificate', methods=['POST'])
@login_required
@admin_required
def admin_add_certificate(id):
    qn = QuanNhan.query.get_or_404(id)

    ten = request.form.get('ten_chung_chi', '').strip()
    if not ten:
        flash('Tên chứng chỉ không được để trống.', 'danger')
        return redirect(url_for('admin.admin_personnel_detail', id=qn.id))

    duong_dan_anh = None
    file = request.files.get('anh_minh_chung')
    if file and file.filename:
        duong_dan_anh = save_upload(file, 'certificates')

    ngay_cap = None
    nc_str = request.form.get('ngay_cap', '').strip()
    if nc_str:
        try:
            ngay_cap = datetime.strptime(nc_str, '%Y-%m-%d').date()
        except ValueError:
            pass

    cc = ChungChi(
        quan_nhan_id=qn.id,
        loai=request.form.get('loai', LoaiChungChi.THANH_TICH_KHAC.value),
        ten_chung_chi=ten,
        so_hieu=request.form.get('so_hieu', '').strip() or None,
        ngay_cap=ngay_cap,
        co_quan_cap=request.form.get('co_quan_cap', '').strip() or None,
        duong_dan_anh=duong_dan_anh,
        ghi_chu=request.form.get('ghi_chu', '').strip() or None,
    )
    db.session.add(cc)
    db.session.commit()
    flash(f'Đã thêm: {ten}', 'success')
    return redirect(url_for('admin.admin_personnel_detail', id=qn.id))


# ------------------------------------------------------------------
# Admin: Delete certificate for any personnel
# ------------------------------------------------------------------
@admin_bp.route('/certificate/<int:id>/delete', methods=['POST'])
@login_required
@admin_required
def admin_delete_certificate(id):
    cc = ChungChi.query.get_or_404(id)
    qn_id = cc.quan_nhan_id

    if cc.duong_dan_anh:
        delete_upload(cc.duong_dan_anh)

    db.session.delete(cc)
    db.session.commit()
    flash('Đã xóa chứng chỉ.', 'success')
    return redirect(url_for('admin.admin_personnel_detail', id=qn_id))


# ------------------------------------------------------------------
# Admin: Manage DanhHieu (Award Titles) - List
# ------------------------------------------------------------------
# Master list of all possible criteria fields (column names on DeXuatChiTiet)
TIEU_CHI_OPTIONS = [
    ('muc_do_hoan_thanh', 'Mức độ hoàn thành nhiệm vụ'),
    ('phieu_tin_nhiem', 'Phiếu tín nhiệm'),
    ('kiem_tra_chinh_tri', 'Kiểm tra chính trị'),
    ('diem_kiem_tra_chinh_tri', 'Điểm kiểm tra chính trị'),
    ('kiem_tra_dieu_lenh', 'Kiểm tra điều lệnh'),
    ('diem_kiem_tra_dieu_lenh', 'Điểm kiểm tra điều lệnh'),
    ('kiem_tra_tin_hoc', 'Kỹ năng số'),
    ('diem_kiem_tra_tin_hoc', 'Điểm kỹ năng số'),
    ('dia_ly_quan_su', 'Địa hình quân sự'),
    ('diem_dia_ly_quan_su', 'Điểm địa hình quân sự'),
    ('ban_sung', 'Bắn súng'),
    ('diem_ban_sung', 'Điểm bắn súng'),
    ('the_luc', 'Thể lực'),
    ('diem_the_luc', 'Điểm thể lực'),
    ('xep_loai_dang_vien', 'Xếp loại đảng viên hằng năm'),
    ('ket_qua_doan_the', 'Kết quả đoàn thể'),
    ('chu_tri_don_vi_danh_hieu', 'Chủ trì đơn vị đạt danh hiệu'),
    ('danh_hieu_gv_gioi', 'Danh hiệu GV giỏi'),
    ('dinh_muc_giang_day', 'Định mức giảng dạy'),
    ('ket_qua_kiem_tra_giang', 'Kết quả kiểm tra giảng'),
    ('tien_do_pgs', 'Tiến độ PGS'),
    ('thoi_gian_lao_dong_kh', 'Thời gian lao động khoa học'),
    ('danh_hieu_hv_gioi', 'Danh hiệu HV giỏi'),
    ('diem_tong_ket', 'Điểm tổng kết'),
    ('ket_qua_thuc_hanh', 'Kết quả thực hành'),
    ('diem_nckh', 'Điểm NCKH'),
    ('nckh_noi_dung', 'Nội dung NCKH'),
    ('nckh_minh_chung', 'Minh chứng NCKH'),
    ('thanh_tich_ca_nhan_khac', 'Thành tích cá nhân khác'),
]


@admin_bp.route('/danh-hieu')
@login_required
@admin_required
def manage_danh_hieu():
    danh_hieus = DanhHieu.query.order_by(DanhHieu.thu_tu, DanhHieu.ten_danh_hieu).all()
    # Build tieu_chi_options from DB (fallback to hardcoded TIEU_CHI_OPTIONS for safety)
    db_tieu_chi = TieuChi.query.filter_by(is_active=True).order_by(TieuChi.thu_tu).all()
    if db_tieu_chi:
        tc_options = [(tc.ma_truong, tc.ten) for tc in db_tieu_chi]
    else:
        tc_options = TIEU_CHI_OPTIONS
    return render_template('admin/manage_danh_hieu.html',
                           danh_hieus=danh_hieus,
                           tieu_chi_options=tc_options)


# ------------------------------------------------------------------
# Admin: Create DanhHieu
# ------------------------------------------------------------------
@admin_bp.route('/danh-hieu/create', methods=['POST'])
@login_required
@admin_required
def create_danh_hieu():
    ten = request.form.get('ten_danh_hieu', '').strip()
    ma = request.form.get('ma_danh_hieu', '').strip()
    pham_vi = request.form.get('pham_vi', 'Cá nhân').strip()
    thu_tu = request.form.get('thu_tu', 0, type=int)
    tieu_chi = request.form.getlist('tieu_chi')

    if not ten or not ma:
        flash('Tên và mã danh hiệu không được để trống.', 'danger')
        return redirect(url_for('admin.manage_danh_hieu'))

    # Check uniqueness
    if DanhHieu.query.filter_by(ten_danh_hieu=ten).first():
        flash(f'Danh hiệu "{ten}" đã tồn tại.', 'danger')
        return redirect(url_for('admin.manage_danh_hieu'))
    if DanhHieu.query.filter_by(ma_danh_hieu=ma).first():
        flash(f'Mã "{ma}" đã tồn tại.', 'danger')
        return redirect(url_for('admin.manage_danh_hieu'))

    dh = DanhHieu(
        ten_danh_hieu=ten,
        ma_danh_hieu=ma,
        pham_vi=pham_vi,
        thu_tu=thu_tu,
        tieu_chi=tieu_chi,
    )
    db.session.add(dh)
    db.session.commit()
    flash(f'Đã thêm danh hiệu: {ten}', 'success')
    return redirect(url_for('admin.manage_danh_hieu'))


# ------------------------------------------------------------------
# Admin: Edit DanhHieu
# ------------------------------------------------------------------
@admin_bp.route('/danh-hieu/<int:id>/edit', methods=['GET', 'POST'])
@login_required
@admin_required
def edit_danh_hieu(id):
    dh = DanhHieu.query.get_or_404(id)

    if request.method == 'POST':
        ten = request.form.get('ten_danh_hieu', '').strip()
        ma = request.form.get('ma_danh_hieu', '').strip()
        pham_vi = request.form.get('pham_vi', 'Cá nhân').strip()
        thu_tu = request.form.get('thu_tu', 0, type=int)
        tieu_chi = request.form.getlist('tieu_chi')

        if not ten or not ma:
            flash('Tên và mã danh hiệu không được để trống.', 'danger')
            return redirect(url_for('admin.edit_danh_hieu', id=id))

        # Check uniqueness (excluding current)
        dup_ten = DanhHieu.query.filter(DanhHieu.ten_danh_hieu == ten, DanhHieu.id != id).first()
        if dup_ten:
            flash(f'Danh hiệu "{ten}" đã tồn tại.', 'danger')
            return redirect(url_for('admin.edit_danh_hieu', id=id))
        dup_ma = DanhHieu.query.filter(DanhHieu.ma_danh_hieu == ma, DanhHieu.id != id).first()
        if dup_ma:
            flash(f'Mã "{ma}" đã tồn tại.', 'danger')
            return redirect(url_for('admin.edit_danh_hieu', id=id))

        dh.ten_danh_hieu = ten
        dh.ma_danh_hieu = ma
        dh.pham_vi = pham_vi
        dh.thu_tu = thu_tu
        dh.tieu_chi = tieu_chi
        db.session.commit()
        flash(f'Đã cập nhật danh hiệu: {ten}', 'success')
        return redirect(url_for('admin.manage_danh_hieu'))

    return render_template('admin/edit_danh_hieu.html',
                           dh=dh,
                           tieu_chi_options=[(tc.ma_truong, tc.ten) for tc in TieuChi.query.filter_by(is_active=True).order_by(TieuChi.thu_tu).all()] or TIEU_CHI_OPTIONS,
                           tieu_chi_db=TieuChi.query.filter_by(is_active=True).order_by(TieuChi.thu_tu).all())


# ------------------------------------------------------------------
# Admin: Toggle DanhHieu active/inactive
# ------------------------------------------------------------------
@admin_bp.route('/danh-hieu/<int:id>/toggle', methods=['POST'])
@login_required
@admin_required
def toggle_danh_hieu(id):
    dh = DanhHieu.query.get_or_404(id)
    dh.is_active = not dh.is_active
    db.session.commit()
    status = 'kích hoạt' if dh.is_active else 'vô hiệu hóa'
    flash(f'Đã {status} danh hiệu: {dh.ten_danh_hieu}', 'success')
    return redirect(url_for('admin.manage_danh_hieu'))


# ------------------------------------------------------------------
# Admin: Delete DanhHieu
# ------------------------------------------------------------------
@admin_bp.route('/danh-hieu/<int:id>/delete', methods=['POST'])
@login_required
@admin_required
def delete_danh_hieu(id):
    dh = DanhHieu.query.get_or_404(id)

    # Check if any DeXuatChiTiet uses this danh_hieu
    usage_count = DeXuatChiTiet.query.filter_by(loai_danh_hieu=dh.ten_danh_hieu).count()
    if usage_count > 0:
        flash(f'Không thể xóa: Danh hiệu đang được sử dụng bởi {usage_count} đề xuất.', 'danger')
        return redirect(url_for('admin.manage_danh_hieu'))

    db.session.delete(dh)
    db.session.commit()
    flash(f'Đã xóa danh hiệu: {dh.ten_danh_hieu}', 'success')
    return redirect(url_for('admin.manage_danh_hieu'))


# ------------------------------------------------------------------
# Admin: Manage TieuChi (Criteria) - List
# ------------------------------------------------------------------
PHONG_DUYET_OPTIONS = [
    ('Phòng Khoa học', 'Phòng Khoa học'),
    ('Phòng Đào tạo', 'Phòng Đào tạo'),
    ('Thủ trưởng Phòng Chính trị', 'Thủ trưởng Phòng Chính trị'),
    ('Thủ trưởng Phòng TM-HC', 'Thủ trưởng Phòng TM-HC'),
    ('Ban Cán bộ', 'Ban Cán bộ'),
    ('Ban Tổ chức', 'Ban Tổ chức'),
    ('Ban Tuyên huấn', 'Ban Tuyên huấn'),
    ('Ban Công tác quần chúng', 'Ban Công tác quần chúng'),
    ('Ban Công nghệ thông tin', 'Ban Công nghệ thông tin'),
    ('Ban Tác huấn', 'Ban Tác huấn'),
    ('Ban Khảo thí', 'Ban Khảo thí'),
    ('Ban Quân lực', 'Ban Quân lực'),
]


@admin_bp.route('/tieu-chi')
@login_required
@admin_required
def manage_tieu_chi():
    tieu_chis = TieuChi.query.order_by(TieuChi.thu_tu, TieuChi.ten).all()

    model_fields = set(col.name for col in DeXuatChiTiet.__table__.columns)
    db_fields = set(tc.ma_truong for tc in tieu_chis)
    missing_fields = sorted([f for f in model_fields if f not in db_fields and f not in {
        'id', 'de_xuat_id', 'quan_nhan_id', 'loai_danh_hieu', 'doi_tuong', 'nam_hoc',
        'ghi_chu', 'created_at', 'updated_at'
    }])

    return render_template('admin/manage_tieu_chi.html',
                           tieu_chis=tieu_chis,
                           nhom_choices=TieuChi.NHOM_CHOICES,
                           phong_duyet_options=PHONG_DUYET_OPTIONS,
                           missing_fields=missing_fields)


# ------------------------------------------------------------------
# Admin: Create TieuChi
# ------------------------------------------------------------------
@admin_bp.route('/tieu-chi/create', methods=['POST'])
@login_required
@admin_required
def create_tieu_chi():
    ma = request.form.get('ma_truong', '').strip()
    ten = request.form.get('ten', '').strip()
    huong_dan = request.form.get('huong_dan', '').strip() or None
    nhom = request.form.get('nhom', 'chung').strip()
    co_minh_chung = request.form.get('co_minh_chung') == '1'
    phong_duyet = request.form.getlist('phong_duyet')
    thu_tu = request.form.get('thu_tu', 0, type=int)

    if not ma or not ten:
        flash('Mã trường và tên tiêu chí không được để trống.', 'danger')
        return redirect(url_for('admin.manage_tieu_chi'))

    if TieuChi.query.filter_by(ma_truong=ma).first():
        flash(f'Mã trường "{ma}" đã tồn tại.', 'danger')
        return redirect(url_for('admin.manage_tieu_chi'))

    tc = TieuChi(
        ma_truong=ma,
        ten=ten,
        huong_dan=huong_dan,
        nhom=nhom,
        co_minh_chung=co_minh_chung,
        thu_tu=thu_tu,
    )
    tc.phong_duyet = phong_duyet
    db.session.add(tc)
    db.session.commit()
    flash(f'Đã thêm tiêu chí: {ten}', 'success')
    return redirect(url_for('admin.manage_tieu_chi'))


# ------------------------------------------------------------------
# Admin: Edit TieuChi
# ------------------------------------------------------------------
@admin_bp.route('/tieu-chi/<int:id>/edit', methods=['GET', 'POST'])
@login_required
@admin_required
def edit_tieu_chi(id):
    tc = TieuChi.query.get_or_404(id)

    if request.method == 'POST':
        ma = request.form.get('ma_truong', '').strip()
        ten = request.form.get('ten', '').strip()
        huong_dan = request.form.get('huong_dan', '').strip() or None
        nhom = request.form.get('nhom', 'chung').strip()
        co_minh_chung = request.form.get('co_minh_chung') == '1'
        phong_duyet = request.form.getlist('phong_duyet')
        thu_tu = request.form.get('thu_tu', 0, type=int)

        if not ma or not ten:
            flash('Mã trường và tên tiêu chí không được để trống.', 'danger')
            return redirect(url_for('admin.edit_tieu_chi', id=id))

        dup = TieuChi.query.filter(TieuChi.ma_truong == ma, TieuChi.id != id).first()
        if dup:
            flash(f'Mã trường "{ma}" đã tồn tại.', 'danger')
            return redirect(url_for('admin.edit_tieu_chi', id=id))

        tc.ma_truong = ma
        tc.ten = ten
        tc.huong_dan = huong_dan
        tc.nhom = nhom
        tc.co_minh_chung = co_minh_chung
        tc.phong_duyet = phong_duyet
        tc.thu_tu = thu_tu
        db.session.commit()
        flash(f'Đã cập nhật tiêu chí: {ten}', 'success')
        return redirect(url_for('admin.manage_tieu_chi'))

    return render_template('admin/edit_tieu_chi.html',
                           tc=tc,
                           nhom_choices=TieuChi.NHOM_CHOICES,
                           phong_duyet_options=PHONG_DUYET_OPTIONS)


# ------------------------------------------------------------------
# Admin: Toggle TieuChi active/inactive
# ------------------------------------------------------------------
@admin_bp.route('/tieu-chi/<int:id>/toggle', methods=['POST'])
@login_required
@admin_required
def toggle_tieu_chi(id):
    tc = TieuChi.query.get_or_404(id)
    tc.is_active = not tc.is_active
    db.session.commit()
    status = 'kích hoạt' if tc.is_active else 'vô hiệu hóa'
    flash(f'Đã {status} tiêu chí: {tc.ten}', 'success')
    return redirect(url_for('admin.manage_tieu_chi'))


# ------------------------------------------------------------------
# Admin: Delete TieuChi
# ------------------------------------------------------------------
@admin_bp.route('/tieu-chi/<int:id>/delete', methods=['POST'])
@login_required
@admin_required
def delete_tieu_chi(id):
    tc = TieuChi.query.get_or_404(id)
    db.session.delete(tc)
    db.session.commit()
    flash(f'Đã xóa tiêu chí: {tc.ten}', 'success')
    return redirect(url_for('admin.manage_tieu_chi'))
