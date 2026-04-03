from io import BytesIO
from flask import Blueprint, render_template, redirect, url_for, flash, request, jsonify, send_file
from flask_login import login_required, current_user
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from app.extensions import db
from app.models.user import User, Role
from app.models.unit import DonVi, LoaiDonVi
from app.models.personnel import QuanNhan
from app.models.nomination import DeXuat, DeXuatChiTiet, TrangThaiDeXuat, LoaiDanhHieu
from app.models.approval import PheDuyet, PhongDuyet, KetQuaDuyet, KetQuaDuyetChiTiet
from app.models.reward import KhenThuong
from app.utils.decorators import admin_required
from datetime import datetime

admin_bp = Blueprint('admin', __name__)

# The six reviewing departments (excluding admin)
DEPT_NAMES = [
    'Phòng Chính trị', 'Phòng Tham mưu', 'Phòng Khoa học', 'Phòng Đào tạo',
    'Ban Cán bộ', 'Ban Quân lực'
]

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
}

# Criteria fields in display order
ALL_FIELDS = [
    'muc_do_hoan_thanh', 'phieu_tin_nhiem',
    'kiem_tra_chinh_tri', 'kiem_tra_dieu_lenh', 'kiem_tra_tin_hoc',
    'dia_ly_quan_su', 'ban_sung', 'the_luc',
    'ket_qua_doan_the', 'chu_tri_don_vi_danh_hieu',
    'danh_hieu_gv_gioi', 'dinh_muc_giang_day', 'ket_qua_kiem_tra_giang',
    'tien_do_pgs', 'thoi_gian_lao_dong_kh',
    'danh_hieu_hv_gioi', 'diem_tong_ket', 'ket_qua_thuc_hanh',
    'diem_nckh', 'nckh_noi_dung', 'nckh_minh_chung',
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

    for dx in nominations:
        unit_name = dx.don_vi.ten_don_vi
        if unit_name not in unit_groups_dict:
            unit_groups_dict[unit_name] = []

        # Build dept approval lookup: dept_name -> {ct_id -> KetQuaDuyetChiTiet}
        dept_lookup = {}
        for pd in dx.phe_duyets:
            if pd.phong_duyet in DEPT_NAMES:
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
            for dept_name in DEPT_NAMES:
                dept_data = dept_lookup.get(dept_name)
                if dept_data:
                    kq = dept_data['items'].get(ct.id)
                    ct_dept_results[dept_name] = kq.ket_qua if kq else None
                    if not kq or kq.ket_qua != KetQuaDuyet.DONG_Y.value:
                        all_dept_ok = False
                else:
                    ct_dept_results[dept_name] = None
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
        dept_item_results[pd.phong_duyet] = {
            'phe_duyet': pd,
            'item_result': kq,
        }

    # Check if all 6 departments approved THIS individual
    all_dept_ok = True
    for dept_name in DEPT_NAMES:
        if dept_name not in dept_item_results:
            all_dept_ok = False
            break
        kq = dept_item_results[dept_name].get('item_result')
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
    ]

    return render_template('admin/tracking_detail.html',
                           ct=ct,
                           de_xuat=de_xuat,
                           dept_item_results=dept_item_results,
                           dept_names=DEPT_NAMES,
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
@admin_required
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
                           stats_by_danh_hieu=stats_by_danh_hieu)


@admin_bp.route('/reward-list/export')
@login_required
@admin_required
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
@admin_required
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
    roles = [(r.value, r.name) for r in Role]
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
