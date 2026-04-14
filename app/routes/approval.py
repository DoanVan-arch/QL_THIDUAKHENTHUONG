from flask import Blueprint, render_template, redirect, url_for, flash, request, jsonify
from flask_login import login_required, current_user
from app.extensions import db
from app.models.user import User, Role
from app.models.unit import DonVi
from app.models.nomination import DeXuat, DeXuatChiTiet, TrangThaiDeXuat, TieuChi
from app.models.approval import PheDuyet, PhongDuyet, KetQuaDuyet, KetQuaDuyetChiTiet
from app.models.notification import ThongBao
from app.utils.decorators import department_required
from datetime import datetime

approval_bp = Blueprint('approval', __name__)

ROLE_TO_PHONG = {
    Role.PHONG_CHINHTRI: PhongDuyet.PHONG_CHINHTRI.value,
    Role.PHONG_THAMMUU: PhongDuyet.PHONG_THAMMUU.value,
    Role.PHONG_KHOAHOC: PhongDuyet.PHONG_KHOAHOC.value,
    Role.PHONG_DAOTAO: PhongDuyet.PHONG_DAOTAO.value,
    Role.THU_TRUONG_PHONG_CHINHTRI: PhongDuyet.THU_TRUONG_PHONG_CHINHTRI.value,
    Role.THU_TRUONG_PHONG_TMHC: PhongDuyet.THU_TRUONG_PHONG_TMHC.value,
    Role.BAN_CANBO: PhongDuyet.BAN_CANBO.value,
    Role.BAN_TOCHUC: PhongDuyet.BAN_TOCHUC.value,
    Role.BAN_TUYENHUAN: PhongDuyet.BAN_TUYENHUAN.value,
    Role.BAN_CTCQ: PhongDuyet.BAN_CTCQ.value,
    Role.BAN_CNTT: PhongDuyet.BAN_CNTT.value,
    Role.BAN_TAC_HUAN: PhongDuyet.BAN_TAC_HUAN.value,
    Role.BAN_KHAOTHI: PhongDuyet.BAN_KHAOTHI.value,
    Role.BAN_QUANLUC: PhongDuyet.BAN_QUANLUC.value,
}

_GROUP_CONFIRMATION = {
    Role.THU_TRUONG_PHONG_CHINHTRI: {
        PhongDuyet.BAN_CANBO.value,
        PhongDuyet.BAN_TOCHUC.value,
        PhongDuyet.BAN_TUYENHUAN.value,
        PhongDuyet.BAN_CTCQ.value,
    },
    Role.THU_TRUONG_PHONG_TMHC: {
        PhongDuyet.BAN_TAC_HUAN.value,
        PhongDuyet.BAN_QUANLUC.value,
        PhongDuyet.BAN_CNTT.value,
    },
}

# Reverse map: PhongDuyet display name -> Role
_PHONG_TO_ROLE = {v: k for k, v in ROLE_TO_PHONG.items()}

# --- Hardcoded fallbacks (used when DB has no tieu_chi rows) ---
_FALLBACK_PHONG_FIELDS = {
    Role.PHONG_DAOTAO: [
        'danh_hieu_gv_gioi', 'tien_do_pgs', 'dinh_muc_giang_day',
        'ket_qua_kiem_tra_giang', 'danh_hieu_hv_gioi', 'diem_tong_ket', 'ket_qua_thuc_hanh',
    ],
    Role.PHONG_KHOAHOC: ['thoi_gian_lao_dong_kh', 'diem_nckh', 'nckh_noi_dung', 'nckh_minh_chung'],
    Role.PHONG_THAMMUU: [
        'kiem_tra_tin_hoc', 'kiem_tra_dieu_lenh', 'dia_ly_quan_su', 'ban_sung', 'the_luc', 'ket_qua_doan_the',
    ],
    Role.PHONG_CHINHTRI: ['kiem_tra_chinh_tri', 'ket_qua_doan_the', 'xep_loai_dang_vien'],
    Role.BAN_CANBO: ['muc_do_hoan_thanh'],
    Role.BAN_QUANLUC: ['muc_do_hoan_thanh'],
}

_FALLBACK_FIELD_LABELS = {
    'muc_do_hoan_thanh': 'Hoàn thành NV', 'phieu_tin_nhiem': 'Tín nhiệm',
    'kiem_tra_dieu_lenh': 'Điều lệnh', 'ban_sung': 'Bắn súng', 'the_luc': 'Thể lực',
    'kiem_tra_chinh_tri': 'Chính trị', 'kiem_tra_tin_hoc': 'Kỹ năng số',
    'dia_ly_quan_su': 'Địa hình QS', 'danh_hieu_gv_gioi': 'GV giỏi',
    'xep_loai_dang_vien': 'Xếp loại ĐV',
    'dinh_muc_giang_day': 'Định mức GD', 'ket_qua_kiem_tra_giang': 'KT giảng',
    'thoi_gian_lao_dong_kh': 'LĐ KH', 'tien_do_pgs': 'Tiến độ PGS',
    'danh_hieu_hv_gioi': 'HV giỏi', 'diem_tong_ket': 'Điểm TK',
    'ket_qua_thuc_hanh': 'Thực hành', 'ket_qua_doan_the': 'Đoàn thể',
    'chu_tri_don_vi_danh_hieu': 'Chủ trì ĐV', 'diem_nckh': 'Điểm KH',
    'nckh_noi_dung': 'NCKH', 'nckh_minh_chung': 'MC NCKH',
    'thanh_tich_ca_nhan_khac': 'Thành tích khác',
}

# Long text / file fields excluded from table columns
_LONG_TEXT_FIELDS = {'nckh_noi_dung', 'nckh_minh_chung', 'tien_do_pgs', 'thanh_tich_ca_nhan_khac'}


def _load_phong_fields_from_db():
    """Build PHONG_FIELDS dict from TieuChi table. Returns None if table is empty."""
    rows = TieuChi.query.filter_by(is_active=True).all()
    if not rows:
        return None
    result = {}
    for tc in rows:
        for pd_name in tc.phong_duyet:
            role = _PHONG_TO_ROLE.get(pd_name)
            if role:
                result.setdefault(role, [])
                if tc.ma_truong not in result[role]:
                    result[role].append(tc.ma_truong)
    return result


def _load_field_labels_from_db():
    """Build FIELD_LABELS dict from TieuChi table. Returns None if table is empty."""
    rows = TieuChi.query.filter_by(is_active=True).all()
    if not rows:
        return None
    return {tc.ma_truong: tc.ten for tc in rows}


def get_phong_fields():
    """Get department -> fields mapping, preferring DB over hardcoded fallback."""
    result = _load_phong_fields_from_db()
    return result if result else _FALLBACK_PHONG_FIELDS


def get_field_labels():
    """Get field -> label mapping, preferring DB over hardcoded fallback."""
    result = _load_field_labels_from_db()
    return result if result else _FALLBACK_FIELD_LABELS


def get_phong_table_columns():
    """Get department -> table column fields (excluding long text/file fields)."""
    phong_fields = get_phong_fields()
    return {role: [f for f in fields if f not in _LONG_TEXT_FIELDS]
            for role, fields in phong_fields.items()}


# Conditional field visibility by doi_tuong (remains hardcoded — specific to business logic)
PHONG_FIELD_CONDITIONS = {
    Role.BAN_CANBO: {
        'muc_do_hoan_thanh': ['Giảng viên', 'Cán bộ'],
    },
    Role.BAN_QUANLUC: {
        'muc_do_hoan_thanh': ['Công nhân viên', 'Quân nhân chuyên nghiệp', 'Công chức quốc phòng'],
    },
}

# doi_tuong scope: which doi_tuong values each department approves
# Departments not listed approve ALL doi_tuong values
BAN_QUANLUC_DOI_TUONG = ['Công nhân viên', 'Quân nhân chuyên nghiệp', 'Công chức quốc phòng']

DEPT_DOI_TUONG_SCOPE = {
    Role.BAN_QUANLUC: BAN_QUANLUC_DOI_TUONG,
    # BAN_CANBO approves all EXCEPT BAN_QUANLUC's scope
}


def _is_in_dept_scope(role, doi_tuong):
    """Check if an individual's doi_tuong falls within the department's approval scope."""
    if role == Role.BAN_QUANLUC:
        return doi_tuong in BAN_QUANLUC_DOI_TUONG
    elif role == Role.BAN_CANBO:
        return doi_tuong not in BAN_QUANLUC_DOI_TUONG
    return True  # All other departments approve all doi_tuong


def _notify_rejections(phe_duyet):
    """Create ThongBao notifications for unit account when individuals are rejected."""
    de_xuat = phe_duyet.de_xuat
    # Find the unit user account for this don_vi
    unit_user = User.query.filter_by(
        don_vi_id=de_xuat.don_vi_id, role=Role.UNIT_USER
    ).first()
    if not unit_user:
        return

    phong_name = phe_duyet.phong_duyet
    for kq in phe_duyet.chi_tiet_duyet:
        if kq.ket_qua == KetQuaDuyet.TU_CHOI.value:
            ct = kq.chi_tiet
            name = ct.quan_nhan.ho_ten if ct.quan_nhan else de_xuat.don_vi.ten_don_vi
            thong_bao = ThongBao(
                user_id=unit_user.id,
                de_xuat_id=de_xuat.id,
                chi_tiet_id=ct.id,
                loai='tu_choi',
                tieu_de=f'{phong_name} không nhất trí: {name}',
                noi_dung=f'Lý do: {kq.ly_do or "Không rõ"}. Đề xuất năm học {de_xuat.nam_hoc}.',
            )
            db.session.add(thong_bao)


@approval_bp.route('/pending')
@login_required
@department_required
def pending_list():
    phong_name = ROLE_TO_PHONG.get(current_user.role, '')

    if current_user.role in _GROUP_CONFIRMATION:
        all_pending = PheDuyet.query.filter_by(
            phong_duyet=phong_name,
            ket_qua=KetQuaDuyet.CHO_DUYET.value
        ).order_by(PheDuyet.created_at.desc()).all()
        pending_reviews = []
        required_groups = _GROUP_CONFIRMATION[current_user.role]
        for pd in all_pending:
            group_reviews = PheDuyet.query.filter(
                PheDuyet.de_xuat_id == pd.de_xuat_id,
                PheDuyet.phong_duyet.in_(list(required_groups))
            ).all()
            if group_reviews and all(g.ket_qua == KetQuaDuyet.DONG_Y.value for g in group_reviews):
                pending_reviews.append(pd)
    else:
        pending_reviews = PheDuyet.query.filter_by(
            phong_duyet=phong_name,
            ket_qua=KetQuaDuyet.CHO_DUYET.value
        ).order_by(PheDuyet.created_at.desc()).all()

    # Ensure per-item records exist for all chi_tiets
    # For BAN_QUANLUC/BAN_CANBO: auto-approve out-of-scope items
    auto_finalized_ids = []
    for pd in pending_reviews:
        existing_ct_ids = {kq.chi_tiet_id for kq in pd.chi_tiet_duyet}
        for ct in pd.de_xuat.chi_tiets:
            if ct.id not in existing_ct_ids:
                in_scope = _is_in_dept_scope(current_user.role, ct.doi_tuong)
                kq = KetQuaDuyetChiTiet(
                    phe_duyet_id=pd.id,
                    chi_tiet_id=ct.id,
                    ket_qua=KetQuaDuyet.DONG_Y.value if not in_scope else KetQuaDuyet.CHO_DUYET.value,
                )
                db.session.add(kq)
    db.session.commit()

    # Auto-finalize departments where ALL items are out-of-scope (all auto-approved)
    for pd in pending_reviews:
        db.session.refresh(pd)
        if pd.ket_qua != KetQuaDuyet.CHO_DUYET.value:
            continue
        pending_in_scope = [
            kq for kq in pd.chi_tiet_duyet
            if kq.ket_qua == KetQuaDuyet.CHO_DUYET.value
        ]
        if not pending_in_scope:
            # All items are auto-approved (out-of-scope) -> auto-finalize
            pd.ket_qua = KetQuaDuyet.DONG_Y.value
            pd.nguoi_duyet_id = None
            pd.ngay_duyet = datetime.utcnow()
            pd.ghi_chu = 'Tự động duyệt (không có đối tượng thuộc phạm vi)'

            de_xuat = pd.de_xuat
            all_dept = PheDuyet.query.filter_by(de_xuat_id=de_xuat.id).filter(
                PheDuyet.phong_duyet != PhongDuyet.ADMIN_TUYENHUAN.value
            ).all()
            if all(a.ket_qua == KetQuaDuyet.DONG_Y.value for a in all_dept):
                de_xuat.trang_thai = TrangThaiDeXuat.DA_DUYET.value
                existing_admin = PheDuyet.query.filter_by(
                    de_xuat_id=de_xuat.id,
                    phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value
                ).first()
                if not existing_admin:
                    admin_pd = PheDuyet(
                        de_xuat_id=de_xuat.id,
                        phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value,
                        ket_qua=KetQuaDuyet.CHO_DUYET.value,
                    )
                    db.session.add(admin_pd)
            auto_finalized_ids.append(pd.id)
    db.session.commit()

    # Remove auto-finalized from pending list
    pending_reviews = [pd for pd in pending_reviews if pd.id not in auto_finalized_ids]

    # Build item results: pd.id -> {ct.id -> KetQuaDuyetChiTiet}
    all_item_results = {}
    for pd in pending_reviews:
        db.session.refresh(pd)
        all_item_results[pd.id] = {kq.chi_tiet_id: kq for kq in pd.chi_tiet_duyet}

    # Build out-of-scope set for template rendering
    out_of_scope_ct_ids = set()
    if current_user.role in (Role.BAN_QUANLUC, Role.BAN_CANBO):
        for pd in pending_reviews:
            for ct in pd.de_xuat.chi_tiets:
                if not _is_in_dept_scope(current_user.role, ct.doi_tuong):
                    out_of_scope_ct_ids.add(ct.id)

    allowed_fields = get_phong_fields().get(current_user.role, [])
    table_columns = get_phong_table_columns().get(current_user.role, [])
    field_conditions = PHONG_FIELD_CONDITIONS.get(current_user.role, {})

    # Collect unique unit names for dropdown filter
    unit_names = []
    for pd in pending_reviews:
        name = pd.de_xuat.don_vi.ten_don_vi
        if name not in unit_names:
            unit_names.append(name)

    return render_template('approval/pending_list.html',
                           pending_reviews=pending_reviews,
                           all_item_results=all_item_results,
                           phong_name=phong_name,
                           allowed_fields=allowed_fields,
                           table_columns=table_columns,
                           field_labels=get_field_labels(),
                           field_conditions=field_conditions,
                           unit_names=unit_names,
                           out_of_scope_ct_ids=out_of_scope_ct_ids)


@approval_bp.route('/review/<int:id>', methods=['GET'])
@login_required
@department_required
def review_nomination(id):
    de_xuat = DeXuat.query.get_or_404(id)
    phong_name = ROLE_TO_PHONG.get(current_user.role, '')

    phe_duyet = PheDuyet.query.filter_by(
        de_xuat_id=id, phong_duyet=phong_name
    ).first_or_404()

    if current_user.role in _GROUP_CONFIRMATION:
        required_groups = _GROUP_CONFIRMATION[current_user.role]
        group_reviews = PheDuyet.query.filter(
            PheDuyet.de_xuat_id == id,
            PheDuyet.phong_duyet.in_(list(required_groups))
        ).all()
        if not group_reviews or not all(g.ket_qua == KetQuaDuyet.DONG_Y.value for g in group_reviews):
            flash('Chưa đủ kết quả nhất trí từ các ban thuộc nhóm xác nhận.', 'warning')
            return redirect(url_for('approval.pending_list'))

    # Ensure per-item records exist for all chi_tiets
    # For BAN_QUANLUC/BAN_CANBO: auto-approve out-of-scope items
    existing_ct_ids = {kq.chi_tiet_id for kq in phe_duyet.chi_tiet_duyet}
    for ct in de_xuat.chi_tiets:
        if ct.id not in existing_ct_ids:
            in_scope = _is_in_dept_scope(current_user.role, ct.doi_tuong)
            kq = KetQuaDuyetChiTiet(
                phe_duyet_id=phe_duyet.id,
                chi_tiet_id=ct.id,
                ket_qua=KetQuaDuyet.DONG_Y.value if not in_scope else KetQuaDuyet.CHO_DUYET.value,
            )
            db.session.add(kq)
    db.session.commit()

    # Reload
    phe_duyet = PheDuyet.query.get(phe_duyet.id)

    # Build lookup: chi_tiet_id -> KetQuaDuyetChiTiet
    item_results = {kq.chi_tiet_id: kq for kq in phe_duyet.chi_tiet_duyet}

    # Build out-of-scope set for template
    out_of_scope_ct_ids = set()
    if current_user.role in (Role.BAN_QUANLUC, Role.BAN_CANBO):
        for ct in de_xuat.chi_tiets:
            if not _is_in_dept_scope(current_user.role, ct.doi_tuong):
                out_of_scope_ct_ids.add(ct.id)

    allowed_fields = get_phong_fields().get(current_user.role, [])
    table_columns = get_phong_table_columns().get(current_user.role, [])
    field_conditions = PHONG_FIELD_CONDITIONS.get(current_user.role, {})

    return render_template('approval/review.html',
                           de_xuat=de_xuat, phe_duyet=phe_duyet,
                           phong_name=phong_name, item_results=item_results,
                           allowed_fields=allowed_fields,
                           table_columns=table_columns,
                           field_labels=get_field_labels(),
                           field_conditions=field_conditions,
                           out_of_scope_ct_ids=out_of_scope_ct_ids)


@approval_bp.route('/review/<int:id>/item/<int:ct_id>/approve', methods=['POST'])
@login_required
@department_required
def approve_item(id, ct_id):
    phong_name = ROLE_TO_PHONG.get(current_user.role, '')
    phe_duyet = PheDuyet.query.filter_by(
        de_xuat_id=id, phong_duyet=phong_name
    ).first_or_404()

    # Block if out of scope for BAN_QUANLUC/BAN_CANBO
    ct = DeXuatChiTiet.query.get_or_404(ct_id)
    if not _is_in_dept_scope(current_user.role, ct.doi_tuong):
        flash('Cá nhân này không thuộc phạm vi duyệt của bạn.', 'warning')
        return redirect(url_for('approval.review_nomination', id=id))

    kq = KetQuaDuyetChiTiet.query.filter_by(
        phe_duyet_id=phe_duyet.id, chi_tiet_id=ct_id
    ).first_or_404()

    kq.ket_qua = KetQuaDuyet.DONG_Y.value
    kq.ly_do = None
    db.session.commit()

    name = ct.quan_nhan.ho_ten if ct.quan_nhan else 'Đơn vị'
    flash(f'Đã nhất trí: {name}', 'success')
    return redirect(url_for('approval.review_nomination', id=id))


@approval_bp.route('/review/<int:id>/item/<int:ct_id>/reject', methods=['POST'])
@login_required
@department_required
def reject_item(id, ct_id):
    phong_name = ROLE_TO_PHONG.get(current_user.role, '')
    phe_duyet = PheDuyet.query.filter_by(
        de_xuat_id=id, phong_duyet=phong_name
    ).first_or_404()

    # Block if out of scope for BAN_QUANLUC/BAN_CANBO
    ct_obj = DeXuatChiTiet.query.get_or_404(ct_id)
    if not _is_in_dept_scope(current_user.role, ct_obj.doi_tuong):
        flash('Cá nhân này không thuộc phạm vi duyệt của bạn.', 'warning')
        return redirect(url_for('approval.review_nomination', id=id))

    kq = KetQuaDuyetChiTiet.query.filter_by(
        phe_duyet_id=phe_duyet.id, chi_tiet_id=ct_id
    ).first_or_404()

    ly_do = request.form.get('ly_do', '').strip()
    if not ly_do:
        flash('Vui lòng nhập lý do không nhất trí.', 'danger')
        return redirect(url_for('approval.review_nomination', id=id))

    kq.ket_qua = KetQuaDuyet.TU_CHOI.value
    kq.ly_do = ly_do
    db.session.commit()

    ct = DeXuatChiTiet.query.get(ct_id)
    name = ct.quan_nhan.ho_ten if ct and ct.quan_nhan else 'Đơn vị'
    flash(f'Đã không nhất trí: {name}', 'warning')
    return redirect(url_for('approval.review_nomination', id=id))


@approval_bp.route('/review/<int:id>/submit', methods=['POST'])
@login_required
@department_required
def submit_review(id):
    phong_name = ROLE_TO_PHONG.get(current_user.role, '')
    phe_duyet = PheDuyet.query.filter_by(
        de_xuat_id=id, phong_duyet=phong_name
    ).first_or_404()

    # Check all items have been reviewed
    pending_items = [kq for kq in phe_duyet.chi_tiet_duyet
                     if kq.ket_qua == KetQuaDuyet.CHO_DUYET.value]
    if pending_items:
        flash(f'Còn {len(pending_items)} cá nhân chưa được duyệt. Vui lòng duyệt tất cả trước khi hoàn tất.', 'danger')
        return redirect(url_for('approval.review_nomination', id=id))

    has_rejection = any(kq.ket_qua == KetQuaDuyet.TU_CHOI.value for kq in phe_duyet.chi_tiet_duyet)

    if has_rejection:
        phe_duyet.ket_qua = KetQuaDuyet.TU_CHOI.value
        phe_duyet.ly_do = 'Một số cá nhân không đạt yêu cầu'
    else:
        phe_duyet.ket_qua = KetQuaDuyet.DONG_Y.value

    phe_duyet.nguoi_duyet_id = current_user.id
    phe_duyet.ngay_duyet = datetime.utcnow()
    phe_duyet.ghi_chu = request.form.get('ghi_chu', '').strip() or None

    de_xuat = DeXuat.query.get(id)

    if has_rejection:
        de_xuat.trang_thai = TrangThaiDeXuat.TU_CHOI.value
    else:
        all_dept_approvals = PheDuyet.query.filter_by(de_xuat_id=id).filter(
            PheDuyet.phong_duyet != PhongDuyet.ADMIN_TUYENHUAN.value
        ).all()

        if all(a.ket_qua == KetQuaDuyet.DONG_Y.value for a in all_dept_approvals):
            de_xuat.trang_thai = TrangThaiDeXuat.DA_DUYET.value
            existing_admin = PheDuyet.query.filter_by(
                de_xuat_id=id, phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value
            ).first()
            if not existing_admin:
                admin_pd = PheDuyet(
                    de_xuat_id=id,
                    phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value,
                    ket_qua=KetQuaDuyet.CHO_DUYET.value,
                )
                db.session.add(admin_pd)
        else:
            de_xuat.trang_thai = TrangThaiDeXuat.DANG_DUYET.value

    db.session.commit()
    if has_rejection:
        _notify_rejections(phe_duyet)
        db.session.commit()
        flash(f'{phong_name} đã hoàn tất duyệt - có cá nhân không nhất trí.', 'warning')
    else:
        flash(f'{phong_name} đã nhất trí toàn bộ đề xuất.', 'success')
    return redirect(url_for('approval.pending_list'))


@approval_bp.route('/toggle/<int:pd_id>/<int:ct_id>', methods=['POST'])
@login_required
@department_required
def toggle_item(pd_id, ct_id):
    phong_name = ROLE_TO_PHONG.get(current_user.role, '')
    phe_duyet = PheDuyet.query.filter_by(
        id=pd_id, phong_duyet=phong_name
    ).first_or_404()

    # Block if out of scope for BAN_QUANLUC/BAN_CANBO
    ct_obj = DeXuatChiTiet.query.get_or_404(ct_id)
    if not _is_in_dept_scope(current_user.role, ct_obj.doi_tuong):
        return jsonify({'success': False, 'message': 'Cá nhân này không thuộc phạm vi duyệt của bạn.'}), 403

    kq = KetQuaDuyetChiTiet.query.filter_by(
        phe_duyet_id=phe_duyet.id, chi_tiet_id=ct_id
    ).first_or_404()

    data = request.get_json()
    approved = data.get('approved', True)
    ly_do = data.get('ly_do', '').strip()

    if approved:
        kq.ket_qua = KetQuaDuyet.DONG_Y.value
        kq.ly_do = None
    else:
        if not ly_do:
            return jsonify({'success': False, 'message': 'Vui lòng nhập lý do'}), 400
        kq.ket_qua = KetQuaDuyet.TU_CHOI.value
        kq.ly_do = ly_do

    db.session.commit()

    # Check if all items decided -> auto-finalize
    pending_count = KetQuaDuyetChiTiet.query.filter_by(
        phe_duyet_id=phe_duyet.id,
        ket_qua=KetQuaDuyet.CHO_DUYET.value
    ).count()

    auto_finalized = False
    if pending_count == 0:
        has_rejection = KetQuaDuyetChiTiet.query.filter_by(
            phe_duyet_id=phe_duyet.id,
            ket_qua=KetQuaDuyet.TU_CHOI.value
        ).count() > 0

        if has_rejection:
            phe_duyet.ket_qua = KetQuaDuyet.TU_CHOI.value
            phe_duyet.ly_do = 'Có cá nhân không đạt yêu cầu'
        else:
            phe_duyet.ket_qua = KetQuaDuyet.DONG_Y.value

        phe_duyet.nguoi_duyet_id = current_user.id
        phe_duyet.ngay_duyet = datetime.utcnow()

        de_xuat = phe_duyet.de_xuat
        if has_rejection:
            de_xuat.trang_thai = TrangThaiDeXuat.TU_CHOI.value
        else:
            all_dept = PheDuyet.query.filter_by(de_xuat_id=de_xuat.id).filter(
                PheDuyet.phong_duyet != PhongDuyet.ADMIN_TUYENHUAN.value
            ).all()
            if all(a.ket_qua == KetQuaDuyet.DONG_Y.value for a in all_dept):
                de_xuat.trang_thai = TrangThaiDeXuat.DA_DUYET.value
                existing_admin = PheDuyet.query.filter_by(
                    de_xuat_id=de_xuat.id, phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value
                ).first()
                if not existing_admin:
                    admin_pd = PheDuyet(
                        de_xuat_id=de_xuat.id,
                        phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value,
                        ket_qua=KetQuaDuyet.CHO_DUYET.value,
                    )
                    db.session.add(admin_pd)
            else:
                de_xuat.trang_thai = TrangThaiDeXuat.DANG_DUYET.value

        db.session.commit()
        auto_finalized = True

        if has_rejection:
            _notify_rejections(phe_duyet)
            db.session.commit()

    # Build stats
    all_kq = phe_duyet.chi_tiet_duyet
    total = len(all_kq)
    approved_count = sum(1 for k in all_kq if k.ket_qua == KetQuaDuyet.DONG_Y.value)
    rejected_count = sum(1 for k in all_kq if k.ket_qua == KetQuaDuyet.TU_CHOI.value)

    return jsonify({
        'success': True,
        'ket_qua': kq.ket_qua,
        'auto_finalized': auto_finalized,
        'stats': {
            'total': total,
            'reviewed': approved_count + rejected_count,
            'approved': approved_count,
            'rejected': rejected_count,
        }
    })


@approval_bp.route('/batch-approve', methods=['POST'])
@login_required
@department_required
def batch_approve():
    """Batch approve multiple items at once (by unit or all)."""
    phong_name = ROLE_TO_PHONG.get(current_user.role, '')
    data = request.get_json()
    pd_id = data.get('pd_id')
    ct_ids = data.get('ct_ids', [])

    if not pd_id or not ct_ids:
        return jsonify({'success': False, 'message': 'Thiếu dữ liệu'}), 400

    phe_duyet = PheDuyet.query.filter_by(
        id=pd_id, phong_duyet=phong_name
    ).first_or_404()

    if phe_duyet.ket_qua != KetQuaDuyet.CHO_DUYET.value:
        return jsonify({'success': False, 'message': 'Đã hoàn tất duyệt'}), 400

    approved_names = []
    for ct_id in ct_ids:
        # Skip out-of-scope items for BAN_QUANLUC/BAN_CANBO
        ct = DeXuatChiTiet.query.get(ct_id)
        if ct and not _is_in_dept_scope(current_user.role, ct.doi_tuong):
            continue
        kq = KetQuaDuyetChiTiet.query.filter_by(
            phe_duyet_id=phe_duyet.id, chi_tiet_id=ct_id
        ).first()
        if kq and kq.ket_qua == KetQuaDuyet.CHO_DUYET.value:
            kq.ket_qua = KetQuaDuyet.DONG_Y.value
            kq.ly_do = None
            name = ct.quan_nhan.ho_ten if ct and ct.quan_nhan else 'Đơn vị'
            approved_names.append(name)

    db.session.commit()

    # Check if all items decided -> auto-finalize
    pending_count = KetQuaDuyetChiTiet.query.filter_by(
        phe_duyet_id=phe_duyet.id,
        ket_qua=KetQuaDuyet.CHO_DUYET.value
    ).count()

    auto_finalized = False
    if pending_count == 0:
        has_rejection = KetQuaDuyetChiTiet.query.filter_by(
            phe_duyet_id=phe_duyet.id,
            ket_qua=KetQuaDuyet.TU_CHOI.value
        ).count() > 0

        if has_rejection:
            phe_duyet.ket_qua = KetQuaDuyet.TU_CHOI.value
            phe_duyet.ly_do = 'Có cá nhân không đạt yêu cầu'
        else:
            phe_duyet.ket_qua = KetQuaDuyet.DONG_Y.value

        phe_duyet.nguoi_duyet_id = current_user.id
        phe_duyet.ngay_duyet = datetime.utcnow()

        de_xuat = phe_duyet.de_xuat
        if has_rejection:
            de_xuat.trang_thai = TrangThaiDeXuat.TU_CHOI.value
        else:
            all_dept = PheDuyet.query.filter_by(de_xuat_id=de_xuat.id).filter(
                PheDuyet.phong_duyet != PhongDuyet.ADMIN_TUYENHUAN.value
            ).all()
            if all(a.ket_qua == KetQuaDuyet.DONG_Y.value for a in all_dept):
                de_xuat.trang_thai = TrangThaiDeXuat.DA_DUYET.value
                existing_admin = PheDuyet.query.filter_by(
                    de_xuat_id=de_xuat.id, phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value
                ).first()
                if not existing_admin:
                    admin_pd = PheDuyet(
                        de_xuat_id=de_xuat.id,
                        phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value,
                        ket_qua=KetQuaDuyet.CHO_DUYET.value,
                    )
                    db.session.add(admin_pd)
            else:
                de_xuat.trang_thai = TrangThaiDeXuat.DANG_DUYET.value

        db.session.commit()
        auto_finalized = True

        if has_rejection:
            _notify_rejections(phe_duyet)
            db.session.commit()

    # Build stats
    all_kq = phe_duyet.chi_tiet_duyet
    total = len(all_kq)
    approved_count = sum(1 for k in all_kq if k.ket_qua == KetQuaDuyet.DONG_Y.value)
    rejected_count = sum(1 for k in all_kq if k.ket_qua == KetQuaDuyet.TU_CHOI.value)

    return jsonify({
        'success': True,
        'approved_count': len(approved_names),
        'auto_finalized': auto_finalized,
        'stats': {
            'total': total,
            'reviewed': approved_count + rejected_count,
            'approved': approved_count,
            'rejected': rejected_count,
        }
    })


@approval_bp.route('/history')
@login_required
@department_required
def history():
    phong_name = ROLE_TO_PHONG.get(current_user.role, '')

    page = request.args.get('page', 1, type=int)
    unit_filter = request.args.get('unit', '')
    danh_hieu_filter = request.args.get('danh_hieu', '')
    ket_qua_filter = request.args.get('ket_qua', '')

    # Query per-individual results (KetQuaDuyetChiTiet) for this department
    query = db.session.query(KetQuaDuyetChiTiet).join(
        PheDuyet, KetQuaDuyetChiTiet.phe_duyet_id == PheDuyet.id
    ).filter(
        PheDuyet.phong_duyet == phong_name,
        PheDuyet.ket_qua != KetQuaDuyet.CHO_DUYET.value,
    )

    # Apply filters
    if ket_qua_filter:
        query = query.filter(KetQuaDuyetChiTiet.ket_qua == ket_qua_filter)

    if unit_filter:
        query = query.join(DeXuat, PheDuyet.de_xuat_id == DeXuat.id).join(
            DonVi, DeXuat.don_vi_id == DonVi.id
        ).filter(DonVi.ten_don_vi == unit_filter)
    else:
        query = query.join(DeXuat, PheDuyet.de_xuat_id == DeXuat.id)

    if danh_hieu_filter:
        query = query.join(
            DeXuatChiTiet, KetQuaDuyetChiTiet.chi_tiet_id == DeXuatChiTiet.id
        ).filter(DeXuatChiTiet.loai_danh_hieu == danh_hieu_filter)

    individual_results = query.order_by(
        PheDuyet.ngay_duyet.desc(), KetQuaDuyetChiTiet.id.desc()
    ).paginate(page=page, per_page=20, error_out=False)

    # Get filter options
    unit_names_q = db.session.query(DonVi.ten_don_vi).join(
        DeXuat, DeXuat.don_vi_id == DonVi.id
    ).join(PheDuyet, PheDuyet.de_xuat_id == DeXuat.id).filter(
        PheDuyet.phong_duyet == phong_name,
        PheDuyet.ket_qua != KetQuaDuyet.CHO_DUYET.value,
    ).distinct().order_by(DonVi.ten_don_vi).all()
    unit_names = [u[0] for u in unit_names_q]

    # Summary stats (unfiltered)
    base_q = db.session.query(KetQuaDuyetChiTiet).join(
        PheDuyet, KetQuaDuyetChiTiet.phe_duyet_id == PheDuyet.id
    ).filter(
        PheDuyet.phong_duyet == phong_name,
        PheDuyet.ket_qua != KetQuaDuyet.CHO_DUYET.value,
    )
    stats = {
        'total': base_q.count(),
        'approved': base_q.filter(KetQuaDuyetChiTiet.ket_qua == KetQuaDuyet.DONG_Y.value).count(),
        'rejected': base_q.filter(KetQuaDuyetChiTiet.ket_qua == KetQuaDuyet.TU_CHOI.value).count(),
    }

    allowed_fields = get_phong_fields().get(current_user.role, [])
    table_columns = get_phong_table_columns().get(current_user.role, [])

    return render_template('approval/history.html',
                           individual_results=individual_results,
                           phong_name=phong_name,
                           unit_filter=unit_filter,
                           danh_hieu_filter=danh_hieu_filter,
                           ket_qua_filter=ket_qua_filter,
                           unit_names=unit_names,
                           stats=stats,
                           allowed_fields=allowed_fields,
                           table_columns=table_columns,
                           field_labels=get_field_labels())


@approval_bp.route('/history/chi-tiet/<int:ct_id>')
@login_required
@department_required
def history_detail(ct_id):
    """View detailed info for one individual in history, scoped to current department."""
    phong_name = ROLE_TO_PHONG.get(current_user.role, '')
    ct = DeXuatChiTiet.query.get_or_404(ct_id)
    de_xuat = ct.de_xuat

    # Get this department's PheDuyet and individual result
    phe_duyet = PheDuyet.query.filter_by(
        de_xuat_id=de_xuat.id, phong_duyet=phong_name
    ).first_or_404()

    kq = KetQuaDuyetChiTiet.query.filter_by(
        phe_duyet_id=phe_duyet.id, chi_tiet_id=ct.id
    ).first()

    allowed_fields = get_phong_fields().get(current_user.role, [])

    return render_template('approval/history_detail.html',
                           ct=ct,
                           de_xuat=de_xuat,
                           phe_duyet=phe_duyet,
                           kq=kq,
                           phong_name=phong_name,
                           allowed_fields=allowed_fields,
                           field_labels=get_field_labels())


@approval_bp.route('/revoke/<int:pd_id>', methods=['POST'])
@login_required
@department_required
def revoke_review(pd_id):
    """Revoke (thu hồi) a completed department review, resetting it to pending."""
    phong_name = ROLE_TO_PHONG.get(current_user.role, '')

    phe_duyet = PheDuyet.query.filter_by(
        id=pd_id, phong_duyet=phong_name
    ).first_or_404()

    de_xuat = phe_duyet.de_xuat

    # Cannot revoke after admin final approval
    if de_xuat.trang_thai == TrangThaiDeXuat.PHE_DUYET_CUOI.value:
        flash('Không thể thu hồi - đề xuất đã được phê duyệt cuối cùng.', 'danger')
        return redirect(url_for('approval.history'))

    # Can only revoke a completed review (not one that's still pending)
    if phe_duyet.ket_qua == KetQuaDuyet.CHO_DUYET.value:
        flash('Kết quả duyệt này vẫn đang chờ, không cần thu hồi.', 'warning')
        return redirect(url_for('approval.history'))

    # 1. Reset the PheDuyet record
    phe_duyet.ket_qua = KetQuaDuyet.CHO_DUYET.value
    phe_duyet.nguoi_duyet_id = None
    phe_duyet.ngay_duyet = None
    phe_duyet.ly_do = None
    phe_duyet.ghi_chu = None

    # 2. Reset KetQuaDuyetChiTiet records back to pending
    #    For BAN_QUANLUC/BAN_CANBO: only reset in-scope items, keep out-of-scope as DONG_Y
    for kq in phe_duyet.chi_tiet_duyet:
        ct = kq.chi_tiet
        if ct and not _is_in_dept_scope(current_user.role, ct.doi_tuong):
            # Out-of-scope: keep auto-approved
            continue
        kq.ket_qua = KetQuaDuyet.CHO_DUYET.value
        kq.ly_do = None

    # 3. Handle DeXuat status changes
    old_status = de_xuat.trang_thai

    # If admin Tuyên huấn PheDuyet was already created (status was 'Đã duyệt'),
    # delete it since not all 6 depts approve anymore
    if old_status == TrangThaiDeXuat.DA_DUYET.value:
        admin_pd = PheDuyet.query.filter_by(
            de_xuat_id=de_xuat.id,
            phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value
        ).first()
        if admin_pd:
            # Delete admin's chi_tiet_duyet records if any
            KetQuaDuyetChiTiet.query.filter_by(phe_duyet_id=admin_pd.id).delete()
            db.session.delete(admin_pd)

    # Determine new status: check other departments
    other_depts = PheDuyet.query.filter_by(de_xuat_id=de_xuat.id).filter(
        PheDuyet.phong_duyet != PhongDuyet.ADMIN_TUYENHUAN.value,
        PheDuyet.id != phe_duyet.id  # exclude the one we just revoked
    ).all()

    has_other_completed = any(
        pd.ket_qua != KetQuaDuyet.CHO_DUYET.value for pd in other_depts
    )

    if has_other_completed:
        # At least one other dept has completed their review
        de_xuat.trang_thai = TrangThaiDeXuat.DANG_DUYET.value
    else:
        # No dept has completed review - back to waiting
        de_xuat.trang_thai = TrangThaiDeXuat.CHO_DUYET.value

    db.session.commit()
    flash(f'{phong_name} đã thu hồi kết quả duyệt cho đề xuất của {de_xuat.don_vi.ten_don_vi}.', 'success')
    return redirect(url_for('approval.history'))
