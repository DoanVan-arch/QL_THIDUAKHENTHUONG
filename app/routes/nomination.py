from flask import Blueprint, render_template, redirect, url_for, flash, request, jsonify
from flask_login import login_required, current_user
from app.extensions import db
from app.models.personnel import QuanNhan, DoiTuong, MucDoHoanThanh
from app.models.nomination import DeXuat, DeXuatChiTiet, MinhChung, LoaiDanhHieu, TrangThaiDeXuat, DanhHieu, TieuChi
from app.models.evaluation import NhomTieuChi
from app.models.approval import PheDuyet, PhongDuyet, KetQuaDuyet, KetQuaDuyetChiTiet
from app.models.reward import KhenThuong
from app.models.notification import ThongBao
from app.utils.decorators import unit_user_required
from app.utils.file_upload import save_upload
from datetime import datetime
from sqlalchemy.exc import ProgrammingError, OperationalError

# The six reviewing departments (excluding admin)
DEPT_NAMES = [
    'Phòng Khoa học', 'Phòng Đào tạo',
    'Thủ trưởng Phòng Chính trị', 'Thủ trưởng Phòng TM-HC',
    'Ban Cán bộ', 'Ban Tổ chức', 'Ban Tuyên huấn', 'Ban Công tác quần chúng',
    'Ban Công nghệ thông tin', 'Ban Tác huấn', 'Ban Khảo thí', 'Ủy ban Kiểm tra', 'Ban Quân lực'
]

nomination_bp = Blueprint('nomination', __name__)


def _get_evidence_required_fields():
    rows = TieuChi.query.filter_by(is_active=True, co_minh_chung=True).order_by(TieuChi.thu_tu, TieuChi.ten).all()
    return [{'ma_truong': tc.ma_truong, 'ten': tc.ten, 'nhom': tc.nhom} for tc in rows]


@nomination_bp.route('/history')
@login_required
@unit_user_required
def nomination_history():
    """History of submitted nominations for the unit."""
    if not current_user.don_vi:
        flash('Tài khoản chưa được gán đơn vị.', 'warning')
        return redirect(url_for('dashboard.index'))

    page = request.args.get('page', 1, type=int)
    status_filter = request.args.get('status', '')
    nam_hoc_filter = request.args.get('nam_hoc', '')

    query = DeXuat.query.filter(
        DeXuat.don_vi_id == current_user.don_vi_id,
        DeXuat.trang_thai != TrangThaiDeXuat.NHAP.value,
    )

    if status_filter:
        query = query.filter(DeXuat.trang_thai == status_filter)
    if nam_hoc_filter:
        query = query.filter(DeXuat.nam_hoc == nam_hoc_filter)

    nominations = query.order_by(DeXuat.ngay_gui.desc()).paginate(
        page=page, per_page=10, error_out=False
    )

    # Get distinct nam_hoc values for filter
    nam_hoc_list = [row[0] for row in db.session.query(DeXuat.nam_hoc).filter(
        DeXuat.don_vi_id == current_user.don_vi_id,
        DeXuat.trang_thai != TrangThaiDeXuat.NHAP.value,
    ).distinct().order_by(DeXuat.nam_hoc.desc()).all()]

    # Count notifications (will be set up with ThongBao model)
    unread_count = 0
    try:
        from app.models.notification import ThongBao
        unread_count = ThongBao.query.filter_by(
            user_id=current_user.id, da_doc=False
        ).count()
    except Exception:
        pass

    return render_template('nomination/history.html',
                           nominations=nominations,
                           status_filter=status_filter,
                           nam_hoc_filter=nam_hoc_filter,
                           nam_hoc_list=nam_hoc_list,
                           unread_count=unread_count)


@nomination_bp.route('/')
@login_required
@unit_user_required
def list_nominations():
    if not current_user.don_vi:
        flash('Tài khoản chưa được gán đơn vị.', 'warning')
        return redirect(url_for('dashboard.index'))

    page = request.args.get('page', 1, type=int)
    nominations = DeXuat.query.filter_by(don_vi_id=current_user.don_vi_id)\
        .order_by(DeXuat.ngay_tao.desc())\
        .paginate(page=page, per_page=10, error_out=False)

    return render_template('nomination/list.html', nominations=nominations)


@nomination_bp.route('/<int:id>/delete', methods=['POST'])
@login_required
@unit_user_required
def delete_nomination(id):
    de_xuat = DeXuat.query.get_or_404(id)

    if de_xuat.don_vi_id != current_user.don_vi_id:
        flash('Không có quyền thao tác.', 'danger')
        return redirect(url_for('nomination.list_nominations'))

    if de_xuat.trang_thai not in (TrangThaiDeXuat.NHAP.value, TrangThaiDeXuat.TU_CHOI.value):
        flash('Chỉ được xóa đề xuất ở trạng thái Nháp hoặc Từ chối.', 'warning')
        return redirect(url_for('nomination.list_nominations'))

    chi_tiet_ids = [ct.id for ct in de_xuat.chi_tiets]
    if chi_tiet_ids:
        ThongBao.query.filter(ThongBao.chi_tiet_id.in_(chi_tiet_ids)).delete(synchronize_session=False)
    ThongBao.query.filter_by(de_xuat_id=de_xuat.id).delete(synchronize_session=False)

    db.session.delete(de_xuat)
    db.session.commit()
    flash('Đã xóa đề xuất.', 'success')
    return redirect(url_for('nomination.list_nominations'))


@nomination_bp.route('/create', methods=['GET', 'POST'])
@login_required
@unit_user_required
def create_nomination():
    if not current_user.don_vi:
        flash('Tài khoản chưa được gán đơn vị.', 'warning')
        return redirect(url_for('dashboard.index'))

    if request.method == 'POST':
        nam_hoc = request.form.get('nam_hoc', '').strip()
        if not nam_hoc:
            flash('Năm học không được để trống.', 'danger')
            return redirect(url_for('nomination.create_nomination'))

        de_xuat = DeXuat(
            don_vi_id=current_user.don_vi_id,
            nam_hoc=nam_hoc,
            nguoi_tao_id=current_user.id,
            trang_thai=TrangThaiDeXuat.NHAP.value,
            ghi_chu=request.form.get('ghi_chu', '').strip() or None,
        )
        db.session.add(de_xuat)
        db.session.commit()
        flash('Đã tạo đề xuất mới. Vui lòng thêm cá nhân vào đề xuất.', 'success')
        return redirect(url_for('nomination.edit_nomination', id=de_xuat.id))

    return render_template('nomination/create.html')


@nomination_bp.route('/<int:id>')
@login_required
def detail_nomination(id):
    de_xuat = DeXuat.query.get_or_404(id)

    # Unit users can only see their own
    if current_user.is_unit_user and de_xuat.don_vi_id != current_user.don_vi_id:
        flash('Không có quyền truy cập.', 'danger')
        return redirect(url_for('dashboard.index'))

    # Build per-individual department approval results for tracking
    # Structure: {ct.id: {dept_name: ket_qua_string}}
    dept_lookup = {}
    for pd in de_xuat.phe_duyets:
        if pd.phong_duyet in DEPT_NAMES:
            dept_lookup[pd.phong_duyet] = {
                'phe_duyet': pd,
                'items': {kq.chi_tiet_id: kq for kq in pd.chi_tiet_duyet},
            }

    chi_tiets_data = []
    # Get all chi_tiet_ids that already have KhenThuong records
    approved_ct_ids = set(
        row[0] for row in db.session.query(KhenThuong.chi_tiet_id).filter(
            KhenThuong.de_xuat_id == de_xuat.id
        ).all()
    )

    for ct in de_xuat.chi_tiets:
        ct_dept_results = {}
        for dept_name in DEPT_NAMES:
            dept_data = dept_lookup.get(dept_name)
            if dept_data:
                kq = dept_data['items'].get(ct.id)
                ct_dept_results[dept_name] = kq.ket_qua if kq else None
            else:
                ct_dept_results[dept_name] = None

        chi_tiets_data.append({
            'ct': ct,
            'dept_results': ct_dept_results,
            'is_rewarded': ct.id in approved_ct_ids,
        })

    return render_template('nomination/detail.html',
                           de_xuat=de_xuat,
                           chi_tiets_data=chi_tiets_data,
                           dept_names=DEPT_NAMES)


@nomination_bp.route('/<int:id>/edit', methods=['GET', 'POST'])
@login_required
@unit_user_required
def edit_nomination(id):
    de_xuat = DeXuat.query.get_or_404(id)
    if de_xuat.don_vi_id != current_user.don_vi_id:
        flash('Không có quyền truy cập.', 'danger')
        return redirect(url_for('nomination.list_nominations'))

    if not de_xuat.is_editable:
        flash('Đề xuất này không thể sửa (đã gửi duyệt).', 'warning')
        return redirect(url_for('nomination.detail_nomination', id=id))

    if request.method == 'POST':
        de_xuat.nam_hoc = request.form.get('nam_hoc', de_xuat.nam_hoc).strip()
        de_xuat.ghi_chu = request.form.get('ghi_chu', '').strip() or None
        db.session.commit()
        flash('Đã cập nhật thông tin đề xuất.', 'success')
        return redirect(url_for('nomination.edit_nomination', id=id))

    personnel = QuanNhan.query.filter_by(
        don_vi_id=current_user.don_vi_id, is_active=True
    ).order_by(QuanNhan.ho_ten).all()

    # Find personnel already nominated in this nam_hoc (in submitted nominations or current draft)
    # 1. Already in the current draft
    already_in_current = set(
        ct.quan_nhan_id for ct in de_xuat.chi_tiets if ct.quan_nhan_id
    )
    # 2. Already in other submitted nominations for same nam_hoc
    already_in_other_q = db.session.query(DeXuatChiTiet.quan_nhan_id).join(
        DeXuat, DeXuatChiTiet.de_xuat_id == DeXuat.id
    ).filter(
        DeXuatChiTiet.nam_hoc == de_xuat.nam_hoc,
        DeXuatChiTiet.quan_nhan_id.isnot(None),
        DeXuat.id != de_xuat.id,
        DeXuat.trang_thai != TrangThaiDeXuat.NHAP.value,
    ).all()
    already_in_other = set(row[0] for row in already_in_other_q)

    already_nominated_ids = already_in_current | already_in_other

    danh_hieu_db = DanhHieu.query.filter_by(is_active=True).order_by(DanhHieu.thu_tu).all()
    danh_hieu_list = [dh.ten_danh_hieu for dh in danh_hieu_db]
    # Build a mapping of danh_hieu -> criteria fields for JS
    danh_hieu_tieu_chi = {dh.ten_danh_hieu: dh.tieu_chi for dh in danh_hieu_db}
    doi_tuong_list = [e.value for e in DoiTuong]
    muc_do_list = [e.value for e in MucDoHoanThanh]

    # Build tooltips from TieuChi DB: {ma_truong: huong_dan}
    tieu_chi_db = TieuChi.query.filter_by(is_active=True).all()
    tieu_chi_tooltips = {tc.ma_truong: tc.huong_dan for tc in tieu_chi_db if tc.huong_dan}

    # Build criteria/group metadata to keep unit UI aligned with admin group configuration
    nhom_meta = {}
    try:
        nhom_rows = NhomTieuChi.query.filter_by(is_active=True).order_by(NhomTieuChi.thu_tu, NhomTieuChi.ten_nhom).all()
        nhom_meta = {
            row.ma_nhom: {
                'ten_nhom': row.ten_nhom,
                'doi_tuong_ap_dung': row.doi_tuong_ap_dung or [],
            }
            for row in nhom_rows
        }
    except (ProgrammingError, OperationalError):
        db.session.rollback()

    if not nhom_meta:
        nhom_meta = {
            key: {'ten_nhom': label, 'doi_tuong_ap_dung': []}
            for key, label in TieuChi.NHOM_CHOICES.items()
        }

    criteria_meta = {}
    for tc in tieu_chi_db:
        nhom_info = nhom_meta.get(tc.nhom, {'ten_nhom': tc.nhom, 'doi_tuong_ap_dung': []})
        criteria_meta[tc.ma_truong] = {
            'nhom': tc.nhom,
            'nhom_ten': nhom_info.get('ten_nhom') or tc.nhom,
            'doi_tuong_ap_dung': nhom_info.get('doi_tuong_ap_dung') or [],
        }

    return render_template('nomination/edit.html',
                           de_xuat=de_xuat,
                           personnel=personnel,
                           already_nominated_ids=already_nominated_ids,
                           danh_hieu_list=danh_hieu_list,
                           danh_hieu_tieu_chi=danh_hieu_tieu_chi,
                           doi_tuong_list=doi_tuong_list,
                           muc_do_list=muc_do_list,
                           tieu_chi_tooltips=tieu_chi_tooltips,
                           criteria_meta=criteria_meta,
                           nhom_meta=nhom_meta,
                           evidence_fields=_get_evidence_required_fields())


@nomination_bp.route('/<int:id>/add-item', methods=['POST'])
@login_required
@unit_user_required
def add_nomination_item(id):
    de_xuat = DeXuat.query.get_or_404(id)
    if de_xuat.don_vi_id != current_user.don_vi_id or not de_xuat.is_editable:
        flash('Không có quyền thao tác.', 'danger')
        return redirect(url_for('nomination.list_nominations'))

    quan_nhan_id = request.form.get('quan_nhan_id', type=int)
    loai_danh_hieu = request.form.get('loai_danh_hieu', '').strip()

    if not loai_danh_hieu:
        flash('Vui lòng chọn danh hiệu đề xuất.', 'danger')
        return redirect(url_for('nomination.edit_nomination', id=id))

    # --- Duplicate prevention ---
    if quan_nhan_id:
        # Check if this person is already in the current nomination
        existing_in_current = DeXuatChiTiet.query.filter_by(
            de_xuat_id=de_xuat.id, quan_nhan_id=quan_nhan_id
        ).first()
        if existing_in_current:
            qn_name = QuanNhan.query.get(quan_nhan_id)
            name = qn_name.ho_ten if qn_name else 'Cá nhân'
            flash(f'{name} đã có trong đề xuất này. Không thể thêm trùng.', 'danger')
            return redirect(url_for('nomination.edit_nomination', id=id))

        # Check if this person + nam_hoc already exists in another submitted nomination
        existing_other = db.session.query(DeXuatChiTiet).join(
            DeXuat, DeXuatChiTiet.de_xuat_id == DeXuat.id
        ).filter(
            DeXuatChiTiet.quan_nhan_id == quan_nhan_id,
            DeXuatChiTiet.nam_hoc == de_xuat.nam_hoc,
            DeXuat.id != de_xuat.id,
            DeXuat.trang_thai != TrangThaiDeXuat.NHAP.value,
        ).first()
        if existing_other:
            qn_name = QuanNhan.query.get(quan_nhan_id)
            name = qn_name.ho_ten if qn_name else 'Cá nhân'
            other_unit = existing_other.de_xuat.don_vi.ten_don_vi
            flash(
                f'{name} đã được đề xuất trong năm học {de_xuat.nam_hoc} '
                f'(bởi {other_unit}). Không thể đề xuất trùng.',
                'danger'
            )
            return redirect(url_for('nomination.edit_nomination', id=id))

    # Get personnel info and doi_tuong from form
    qn = QuanNhan.query.get(quan_nhan_id) if quan_nhan_id else None
    doi_tuong = request.form.get('doi_tuong', '').strip() or (qn.doi_tuong if qn else None)

    # Validate score/rating paired fields: must fill both sides together
    pair_rules = [
        ('diem_kiem_tra_tin_hoc', 'kiem_tra_tin_hoc', 'kỹ năng số'),
        ('diem_kiem_tra_dieu_lenh', 'kiem_tra_dieu_lenh', 'điều lệnh'),
        ('diem_dia_ly_quan_su', 'dia_ly_quan_su', 'địa hình quân sự'),
        ('diem_ban_sung', 'ban_sung', 'bắn súng'),
        ('diem_the_luc', 'the_luc', 'thể lực'),
        ('diem_kiem_tra_chinh_tri', 'kiem_tra_chinh_tri', 'chính trị'),
    ]
    for diem_field, xeploai_field, label in pair_rules:
        diem_val = request.form.get(diem_field, '').strip()
        xeploai_val = request.form.get(xeploai_field, '').strip()
        if (diem_val and not xeploai_val) or (xeploai_val and not diem_val):
            flash(f'Tiêu chí {label}: phải nhập đầy đủ cả Điểm và Xếp loại.', 'danger')
            return redirect(url_for('nomination.edit_nomination', id=id))

    chi_tiet = DeXuatChiTiet(
        de_xuat_id=de_xuat.id,
        quan_nhan_id=quan_nhan_id,
        loai_danh_hieu=loai_danh_hieu,
        doi_tuong=doi_tuong,
        nam_hoc=de_xuat.nam_hoc,
        muc_do_hoan_thanh=request.form.get('muc_do_hoan_thanh', '').strip() or None,
        kiem_tra_tin_hoc=request.form.get('kiem_tra_tin_hoc', '').strip() or None,
        diem_kiem_tra_tin_hoc=request.form.get('diem_kiem_tra_tin_hoc', '').strip() or None,
        kiem_tra_dieu_lenh=request.form.get('kiem_tra_dieu_lenh', '').strip() or None,
        diem_kiem_tra_dieu_lenh=request.form.get('diem_kiem_tra_dieu_lenh', '').strip() or None,
        dia_ly_quan_su=request.form.get('dia_ly_quan_su', '').strip() or None,
        diem_dia_ly_quan_su=request.form.get('diem_dia_ly_quan_su', '').strip() or None,
        ban_sung=request.form.get('ban_sung', '').strip() or None,
        diem_ban_sung=request.form.get('diem_ban_sung', '').strip() or None,
        the_luc=request.form.get('the_luc', '').strip() or None,
        diem_the_luc=request.form.get('diem_the_luc', '').strip() or None,
        kiem_tra_chinh_tri=request.form.get('kiem_tra_chinh_tri', '').strip() or None,
        diem_kiem_tra_chinh_tri=request.form.get('diem_kiem_tra_chinh_tri', '').strip() or None,
        phieu_tin_nhiem=request.form.get('phieu_tin_nhiem', '').strip() or None,
        xep_loai_dang_vien=request.form.get('xep_loai_dang_vien', '').strip() or None,
        ket_qua_doan_the=request.form.get('ket_qua_doan_the', '').strip() or None,
        chu_tri_don_vi_danh_hieu=request.form.get('chu_tri_don_vi_danh_hieu', '').strip() or None,
        # Lecturer fields
        danh_hieu_gv_gioi=request.form.get('danh_hieu_gv_gioi', '').strip() or None,
        tien_do_pgs=request.form.get('tien_do_pgs', '').strip() or None,
        dinh_muc_giang_day=request.form.get('dinh_muc_giang_day', '').strip() or None,
        thoi_gian_lao_dong_kh=request.form.get('thoi_gian_lao_dong_kh', '').strip() or None,
        ket_qua_kiem_tra_giang=request.form.get('ket_qua_kiem_tra_giang', '').strip() or None,
        # Student fields
        danh_hieu_hv_gioi=request.form.get('danh_hieu_hv_gioi', '').strip() or None,
        diem_tong_ket=request.form.get('diem_tong_ket', '').strip() or None,
        ket_qua_thuc_hanh=request.form.get('ket_qua_thuc_hanh', '').strip() or None,
        # NCKH
        nckh_noi_dung=request.form.get('nckh_noi_dung', '').strip() or None,
        diem_nckh=float(request.form.get('diem_nckh')) if request.form.get('diem_nckh', '').strip() else None,
        thanh_tich_ca_nhan_khac=request.form.get('thanh_tich_ca_nhan_khac', '').strip() or None,
        ghi_chu=request.form.get('ghi_chu_item', '').strip() or None,
    )

    db.session.add(chi_tiet)
    db.session.flush()

    # Handle file uploads for evidence
    evidence_fields = ['minh_chung_chung']
    for field_name in evidence_fields:
        file = request.files.get(field_name)
        if file and file.filename:
            path = save_upload(file, 'evidence')
            if path:
                mc = MinhChung(
                    chi_tiet_id=chi_tiet.id,
                    loai_minh_chung=field_name,
                    duong_dan=path,
                    ten_file_goc=file.filename,
                )
                db.session.add(mc)

    # NCKH evidence supports one or multiple files
    nckh_files = request.files.getlist('nckh_minh_chung')
    for file in nckh_files:
        if file and file.filename:
            path = save_upload(file, 'evidence')
            if path:
                mc = MinhChung(
                    chi_tiet_id=chi_tiet.id,
                    loai_minh_chung='nckh_minh_chung',
                    duong_dan=path,
                    ten_file_goc=file.filename,
                )
                db.session.add(mc)

    # Dynamic evidence files by configured criteria requiring evidence
    for ef in _get_evidence_required_fields():
        field_key = ef['ma_truong']
        files = request.files.getlist(f'minh_chung_{field_key}')
        for file in files:
            if file and file.filename:
                path = save_upload(file, 'evidence')
                if path:
                    mc = MinhChung(
                        chi_tiet_id=chi_tiet.id,
                        loai_minh_chung=f'minh_chung_{field_key}',
                        duong_dan=path,
                        ten_file_goc=file.filename,
                    )
                    db.session.add(mc)

    # Handle multiple evidence files
    evidence_files = request.files.getlist('minh_chung_files')
    for file in evidence_files:
        if file and file.filename:
            path = save_upload(file, 'evidence')
            if path:
                mc = MinhChung(
                    chi_tiet_id=chi_tiet.id,
                    loai_minh_chung='minh_chung_bo_sung',
                    duong_dan=path,
                    ten_file_goc=file.filename,
                )
                db.session.add(mc)

    # Handle CSTT evidence files
    cstt_files = request.files.getlist('minh_chung_files_cstt')
    for file in cstt_files:
        if file and file.filename:
            path = save_upload(file, 'evidence')
            if path:
                mc = MinhChung(
                    chi_tiet_id=chi_tiet.id,
                    loai_minh_chung='minh_chung_bo_sung',
                    duong_dan=path,
                    ten_file_goc=file.filename,
                )
                db.session.add(mc)

    # Handle evidence files for ket_qua_doan_the
    doan_the_files = request.files.getlist('minh_chung_doan_the')
    for file in doan_the_files:
        if file and file.filename:
            path = save_upload(file, 'evidence')
            if path:
                mc = MinhChung(
                    chi_tiet_id=chi_tiet.id,
                    loai_minh_chung='minh_chung_doan_the',
                    duong_dan=path,
                    ten_file_goc=file.filename,
                )
                db.session.add(mc)

    # Handle evidence files for thanh_tich_ca_nhan_khac
    thanh_tich_files = request.files.getlist('minh_chung_thanh_tich_khac')
    for file in thanh_tich_files:
        if file and file.filename:
            path = save_upload(file, 'evidence')
            if path:
                mc = MinhChung(
                    chi_tiet_id=chi_tiet.id,
                    loai_minh_chung='minh_chung_thanh_tich_khac',
                    duong_dan=path,
                    ten_file_goc=file.filename,
                )
                db.session.add(mc)

    db.session.commit()

    name = qn.ho_ten if qn else 'Đơn vị'
    flash(f'Đã thêm đề xuất cho: {name} - {loai_danh_hieu}', 'success')
    return redirect(url_for('nomination.edit_nomination', id=id))


@nomination_bp.route('/item/<int:id>/delete', methods=['POST'])
@login_required
@unit_user_required
def delete_nomination_item(id):
    chi_tiet = DeXuatChiTiet.query.get_or_404(id)
    de_xuat = chi_tiet.de_xuat

    if de_xuat.don_vi_id != current_user.don_vi_id or not de_xuat.is_editable:
        flash('Không có quyền thao tác.', 'danger')
        return redirect(url_for('nomination.list_nominations'))

    ThongBao.query.filter_by(chi_tiet_id=chi_tiet.id).delete(synchronize_session=False)

    db.session.delete(chi_tiet)
    db.session.commit()
    flash('Đã xóa mục đề xuất.', 'success')
    return redirect(url_for('nomination.edit_nomination', id=de_xuat.id))


@nomination_bp.route('/<int:id>/submit', methods=['POST'])
@login_required
@unit_user_required
def submit_nomination(id):
    de_xuat = DeXuat.query.get_or_404(id)
    if de_xuat.don_vi_id != current_user.don_vi_id:
        flash('Không có quyền thao tác.', 'danger')
        return redirect(url_for('nomination.list_nominations'))

    if not de_xuat.is_editable:
        flash('Đề xuất này không thể gửi (đã gửi duyệt).', 'warning')
        return redirect(url_for('nomination.detail_nomination', id=id))

    if not de_xuat.chi_tiets:
        flash('Đề xuất phải có ít nhất một cá nhân/đơn vị.', 'danger')
        return redirect(url_for('nomination.edit_nomination', id=id))

    # Validate: commander/secretary must have DON_VI_QUYET_THANG
    has_unit_award = any(
        ct.loai_danh_hieu == 'Đơn vị quyết thắng'
        for ct in de_xuat.chi_tiets
    )
    for ct in de_xuat.chi_tiets:
        if ct.quan_nhan and (ct.quan_nhan.la_chi_huy or ct.quan_nhan.la_bi_thu):
            if not has_unit_award:
                flash(
                    'Có chỉ huy/bí thư trong danh sách đề xuất - cần phải có đề xuất "Đơn vị quyết thắng" đi kèm.',
                    'danger'
                )
                return redirect(url_for('nomination.edit_nomination', id=id))

    # Create pending approval records
    has_any_doan_the = any((ct.ket_qua_doan_the or '').strip() for ct in de_xuat.chi_tiets)

    for phong in [PhongDuyet.PHONG_KHOAHOC, PhongDuyet.PHONG_DAOTAO,
                  PhongDuyet.THU_TRUONG_PHONG_CHINHTRI, PhongDuyet.THU_TRUONG_PHONG_TMHC,
                  PhongDuyet.BAN_CANBO, PhongDuyet.BAN_TOCHUC,
                  PhongDuyet.BAN_TUYENHUAN, PhongDuyet.BAN_CTCQ,
                  PhongDuyet.BAN_CNTT, PhongDuyet.BAN_TAC_HUAN,
                  PhongDuyet.BAN_KHAOTHI, PhongDuyet.UY_BAN_KIEMTRA,
                  PhongDuyet.BAN_QUANLUC]:
        existing = PheDuyet.query.filter_by(de_xuat_id=de_xuat.id, phong_duyet=phong.value).first()
        if not existing:
            initial_ket_qua = KetQuaDuyet.CHO_DUYET.value
            initial_ghi_chu = None
            if phong == PhongDuyet.BAN_CTCQ and not has_any_doan_the:
                initial_ket_qua = KetQuaDuyet.DONG_Y.value
                initial_ghi_chu = 'Tự động duyệt (không có dữ liệu kết quả đoàn thể)'

            pd = PheDuyet(
                de_xuat_id=de_xuat.id,
                phong_duyet=phong.value,
                ket_qua=initial_ket_qua,
                ghi_chu=initial_ghi_chu,
            )
            db.session.add(pd)
            db.session.flush()

            if phong == PhongDuyet.BAN_CTCQ and not has_any_doan_the:
                for ct in de_xuat.chi_tiets:
                    exists_item = KetQuaDuyetChiTiet.query.filter_by(
                        phe_duyet_id=pd.id,
                        chi_tiet_id=ct.id,
                    ).first()
                    if not exists_item:
                        db.session.add(KetQuaDuyetChiTiet(
                            phe_duyet_id=pd.id,
                            chi_tiet_id=ct.id,
                            ket_qua=KetQuaDuyet.DONG_Y.value,
                        ))

    de_xuat.trang_thai = TrangThaiDeXuat.CHO_DUYET.value
    de_xuat.ngay_gui = datetime.utcnow()
    db.session.commit()
    flash('Đã gửi đề xuất lên cấp trên. Chờ các cơ quan phê duyệt.', 'success')
    return redirect(url_for('nomination.detail_nomination', id=id))


@nomination_bp.route('/notifications')
@login_required
@unit_user_required
def notifications():
    """Notification list for unit accounts."""
    from app.models.notification import ThongBao

    page = request.args.get('page', 1, type=int)
    thong_baos = ThongBao.query.filter_by(
        user_id=current_user.id
    ).order_by(ThongBao.created_at.desc()).paginate(
        page=page, per_page=20, error_out=False
    )

    return render_template('nomination/notifications.html', thong_baos=thong_baos)


@nomination_bp.route('/notifications/mark-read', methods=['POST'])
@login_required
@unit_user_required
def mark_notifications_read():
    """Mark all notifications as read."""
    from app.models.notification import ThongBao

    ThongBao.query.filter_by(
        user_id=current_user.id, da_doc=False
    ).update({'da_doc': True})
    db.session.commit()
    flash('Đã đánh dấu tất cả thông báo đã đọc.', 'success')
    return redirect(url_for('nomination.notifications'))


@nomination_bp.route('/<int:id>/revoke', methods=['POST'])
@login_required
@unit_user_required
def revoke_nomination(id):
    """Revoke (thu hồi) a submitted nomination back to draft status.
    Only allowed if no department has approved/rejected yet (all still CHO_DUYET).
    """
    de_xuat = DeXuat.query.get_or_404(id)
    if de_xuat.don_vi_id != current_user.don_vi_id:
        flash('Không có quyền thao tác.', 'danger')
        return redirect(url_for('nomination.list_nominations'))

    # Only allow revoking submitted nominations (not drafts, not already final-approved)
    if de_xuat.trang_thai not in (
        TrangThaiDeXuat.CHO_DUYET.value,
        TrangThaiDeXuat.DANG_DUYET.value,
        TrangThaiDeXuat.TU_CHOI.value,
    ):
        flash('Không thể thu hồi đề xuất ở trạng thái này.', 'warning')
        return redirect(url_for('nomination.detail_nomination', id=id))

    # Check that no department has finalized their review (all still CHO_DUYET)
    dept_reviews = PheDuyet.query.filter_by(de_xuat_id=de_xuat.id).filter(
        PheDuyet.phong_duyet != 'Tuyên huấn'
    ).all()

    has_any_decided = any(
        pd.ket_qua != KetQuaDuyet.CHO_DUYET.value for pd in dept_reviews
    )
    if has_any_decided:
        flash('Không thể thu hồi - đã có cơ quan hoàn tất duyệt.', 'danger')
        return redirect(url_for('nomination.detail_nomination', id=id))

    # Delete all PheDuyet and their KetQuaDuyetChiTiet records
    for pd in dept_reviews:
        KetQuaDuyetChiTiet.query.filter_by(phe_duyet_id=pd.id).delete()
        db.session.delete(pd)

    # Reset nomination status
    de_xuat.trang_thai = TrangThaiDeXuat.NHAP.value
    de_xuat.ngay_gui = None

    db.session.commit()
    flash('Đã thu hồi đề xuất. Đề xuất đã chuyển về trạng thái Nháp.', 'success')
    return redirect(url_for('nomination.detail_nomination', id=id))
