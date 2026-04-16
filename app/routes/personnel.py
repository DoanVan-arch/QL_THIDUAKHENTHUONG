from flask import Blueprint, render_template, redirect, url_for, flash, request, send_file
from flask_login import login_required, current_user
from app.extensions import db
from app.models.user import Role
from app.models.personnel import QuanNhan, CapBac, HocHam, HocVi, DoiTuong
from app.models.certificate import ChungChi, LoaiChungChi
from app.models.catalog import ChucVuOption, CapBacOption
from app.models.evaluation import DanhGiaHangNam
from app.utils.decorators import unit_user_required
from app.utils.file_upload import save_upload, delete_upload
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from sqlalchemy import case, collate
from sqlalchemy.orm import aliased
from sqlalchemy.exc import ProgrammingError, OperationalError, IntegrityError

personnel_bp = Blueprint('personnel', __name__)


@personnel_bp.route('/')
@login_required
@unit_user_required
def list_personnel():
    if not current_user.don_vi:
        flash('Tài khoản chưa được gán đơn vị.', 'warning')
        return redirect(url_for('dashboard.index'))

    search = request.args.get('search', '').strip()
    cap_bac_filter = request.args.get('cap_bac', '').strip()
    chuc_vu_filter = request.args.get('chuc_vu', '').strip()
    doi_tuong_filter = request.args.get('doi_tuong', '').strip()
    sort_by = request.args.get('sort_by', 'chuc_vu_order').strip()
    page = request.args.get('page', 1, type=int)

    query = QuanNhan.query.filter_by(don_vi_id=current_user.don_vi_id, is_active=True)
    if search:
        query = query.filter(QuanNhan.ho_ten.ilike(f'%{search}%'))
    if cap_bac_filter:
        query = query.filter(QuanNhan.cap_bac == cap_bac_filter)
    if chuc_vu_filter:
        query = query.filter(QuanNhan.chuc_vu == chuc_vu_filter)
    if doi_tuong_filter:
        query = query.filter(QuanNhan.doi_tuong == doi_tuong_filter)

    chuc_vu_alias = aliased(ChucVuOption)
    query = query.outerjoin(
        chuc_vu_alias,
        collate(chuc_vu_alias.ten, 'utf8mb4_unicode_ci') == collate(QuanNhan.chuc_vu, 'utf8mb4_unicode_ci')
    )

    if sort_by == 'ho_ten':
        query = query.order_by(QuanNhan.ho_ten.asc())
    elif sort_by == 'cap_bac':
        query = query.order_by(
            case((QuanNhan.cap_bac.is_(None), 1), else_=0),
            QuanNhan.cap_bac.asc(),
            QuanNhan.ho_ten.asc()
        )
    elif sort_by == 'doi_tuong':
        query = query.order_by(
            case((QuanNhan.doi_tuong.is_(None), 1), else_=0),
            QuanNhan.doi_tuong.asc(),
            QuanNhan.ho_ten.asc()
        )
    else:
        query = query.order_by(
            case((chuc_vu_alias.thu_tu.is_(None), 1), else_=0),
            chuc_vu_alias.thu_tu.asc(),
            QuanNhan.chuc_vu.asc(),
            QuanNhan.ho_ten.asc()
        )

    personnel = query.paginate(page=page, per_page=20, error_out=False)

    return render_template('personnel/list.html',
                           personnel=personnel,
                           search=search,
                           cap_bac_filter=cap_bac_filter,
                           chuc_vu_filter=chuc_vu_filter,
                           doi_tuong_filter=doi_tuong_filter,
                           sort_by=sort_by,
                           cap_bac_list=_get_cap_bac_list(),
                           chuc_vu_list=[x.ten for x in _get_chuc_vu_options()],
                           doi_tuong_list=[e.value for e in DoiTuong])


def _get_chuc_vu_options():
    return ChucVuOption.query.filter_by(is_active=True).order_by(ChucVuOption.thu_tu, ChucVuOption.ten).all()


def _get_cap_bac_list():
    db_values = [x.ten for x in CapBacOption.query.filter_by(is_active=True).order_by(CapBacOption.thu_tu, CapBacOption.ten).all()]
    return db_values if db_values else [e.value for e in CapBac]


@personnel_bp.route('/create', methods=['GET', 'POST'])
@login_required
@unit_user_required
def create_personnel():
    if not current_user.don_vi:
        flash('Tài khoản chưa được gán đơn vị.', 'warning')
        return redirect(url_for('dashboard.index'))

    if request.method == 'POST':
        ngay_sinh = None
        ns_str = request.form.get('ngay_sinh', '').strip()
        if ns_str:
            try:
                ngay_sinh = datetime.strptime(ns_str, '%Y-%m-%d').date()
            except ValueError:
                pass

        qn = QuanNhan(
            don_vi_id=current_user.don_vi_id,
            ho_ten=request.form.get('ho_ten', '').strip(),
            cap_bac=request.form.get('cap_bac', '').strip() or None,
            chuc_danh=request.form.get('chuc_danh', '').strip() or None,
            chuc_vu=request.form.get('chuc_vu', '').strip() or None,
            don_vi_truc_thuoc=request.form.get('don_vi_truc_thuoc', '').strip() or None,
            can_cuoc_cong_dan=request.form.get('can_cuoc_cong_dan', '').strip() or None,
            ngay_sinh=ngay_sinh,
            ngay_nhap_ngu=request.form.get('ngay_nhap_ngu', '').strip() or None,
            doi_tuong=request.form.get('doi_tuong', '').strip() or None,
            hoc_ham=request.form.get('hoc_ham', 'Không').strip(),
            hoc_vi=request.form.get('hoc_vi', 'Không').strip(),
            trinh_do_hoc_van=request.form.get('trinh_do_hoc_van', '').strip() or None,
            ngoai_ngu=request.form.get('ngoai_ngu', '').strip() or None,
            la_chi_huy='la_chi_huy' in request.form,
            la_bi_thu='la_bi_thu' in request.form,
        )

        if not qn.ho_ten:
            flash('Họ tên không được để trống.', 'danger')
            return render_template('personnel/create.html',
                                   cap_bac_list=_get_cap_bac_list(),
                                   hoc_ham_list=[e.value for e in HocHam],
                                   hoc_vi_list=[e.value for e in HocVi],
                                   doi_tuong_list=[e.value for e in DoiTuong],
                                   chuc_vu_options=_get_chuc_vu_options())

        db.session.add(qn)
        db.session.commit()
        flash(f'Đã thêm quân nhân: {qn.ho_ten}', 'success')
        return redirect(url_for('personnel.detail_personnel', id=qn.id))

    return render_template('personnel/create.html',
                           cap_bac_list=_get_cap_bac_list(),
                           hoc_ham_list=[e.value for e in HocHam],
                           hoc_vi_list=[e.value for e in HocVi],
                           doi_tuong_list=[e.value for e in DoiTuong],
                           chuc_vu_options=_get_chuc_vu_options())


@personnel_bp.route('/<int:id>')
@login_required
@unit_user_required
def detail_personnel(id):
    qn = QuanNhan.query.get_or_404(id)
    if qn.don_vi_id != current_user.don_vi_id:
        flash('Không có quyền truy cập.', 'danger')
        return redirect(url_for('personnel.list_personnel'))

    return render_template('personnel/detail.html', qn=qn,
                           loai_chung_chi_list=[e.value for e in LoaiChungChi])


@personnel_bp.route('/<int:id>/edit', methods=['GET', 'POST'])
@login_required
@unit_user_required
def edit_personnel(id):
    qn = QuanNhan.query.get_or_404(id)
    if qn.don_vi_id != current_user.don_vi_id:
        flash('Không có quyền truy cập.', 'danger')
        return redirect(url_for('personnel.list_personnel'))

    if request.method == 'POST':
        qn.ho_ten = request.form.get('ho_ten', '').strip()
        qn.cap_bac = request.form.get('cap_bac', '').strip() or None
        qn.chuc_danh = request.form.get('chuc_danh', '').strip() or None
        qn.chuc_vu = request.form.get('chuc_vu', '').strip() or None
        qn.don_vi_truc_thuoc = request.form.get('don_vi_truc_thuoc', '').strip() or None
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
            return redirect(url_for('personnel.detail_personnel', id=qn.id))

    return render_template('personnel/edit.html', qn=qn,
                           cap_bac_list=_get_cap_bac_list(),
                           hoc_ham_list=[e.value for e in HocHam],
                           hoc_vi_list=[e.value for e in HocVi],
                           doi_tuong_list=[e.value for e in DoiTuong],
                           chuc_vu_options=_get_chuc_vu_options())


@personnel_bp.route('/<int:id>/delete', methods=['POST'])
@login_required
@unit_user_required
def delete_personnel(id):
    qn = QuanNhan.query.get_or_404(id)
    if qn.don_vi_id != current_user.don_vi_id:
        flash('Không có quyền truy cập.', 'danger')
        return redirect(url_for('personnel.list_personnel'))

    qn.is_active = False
    db.session.commit()
    flash(f'Đã xóa quân nhân: {qn.ho_ten}', 'success')
    return redirect(url_for('personnel.list_personnel'))


@personnel_bp.route('/<int:id>/certificate', methods=['POST'])
@login_required
@unit_user_required
def add_certificate(id):
    qn = QuanNhan.query.get_or_404(id)
    if qn.don_vi_id != current_user.don_vi_id:
        flash('Không có quyền truy cập.', 'danger')
        return redirect(url_for('personnel.list_personnel'))

    ten = request.form.get('ten_chung_chi', '').strip()
    if not ten:
        flash('Tên chứng chỉ không được để trống.', 'danger')
        return redirect(url_for('personnel.detail_personnel', id=qn.id))

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
    return redirect(url_for('personnel.detail_personnel', id=qn.id))


@personnel_bp.route('/certificate/<int:id>/delete', methods=['POST'])
@login_required
@unit_user_required
def delete_certificate(id):
    cc = ChungChi.query.get_or_404(id)
    qn = cc.quan_nhan
    if qn.don_vi_id != current_user.don_vi_id:
        flash('Không có quyền truy cập.', 'danger')
        return redirect(url_for('personnel.list_personnel'))

    if cc.duong_dan_anh:
        delete_upload(cc.duong_dan_anh)

    db.session.delete(cc)
    db.session.commit()
    flash('Đã xóa chứng chỉ.', 'success')
    return redirect(url_for('personnel.detail_personnel', id=qn.id))


def _parse_bool(value):
    if value is None:
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value == 1
    return str(value).strip().lower() in {'1', 'true', 'yes', 'y', 'x'}


def _parse_date(value):
    if value is None or value == '':
        return None
    if isinstance(value, datetime):
        return value.date()
    if hasattr(value, 'date'):
        try:
            return value.date()
        except Exception:
            pass
    text = str(value).strip()
    for fmt in ('%Y-%m-%d', '%d/%m/%Y'):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _normalize_cccd(value):
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    if text.endswith('.0') and text.replace('.', '', 1).isdigit():
        text = text[:-2]
    return text


@personnel_bp.route('/template')
@login_required
@unit_user_required
def download_personnel_template():
    wb = Workbook()
    ws = wb.active
    ws.title = 'QuanNhan'

    lookup = wb.create_sheet('DanhMuc')
    lookup.sheet_state = 'hidden'

    cap_bac_values = _get_cap_bac_list()
    doi_tuong_values = [e.value for e in DoiTuong]
    hoc_ham_values = [e.value for e in HocHam]
    hoc_vi_values = [e.value for e in HocVi]
    chuc_vu_values = [x.ten for x in _get_chuc_vu_options()]
    bool_values = ['0', '1']

    data_sets = [
        ('A', cap_bac_values),
        ('B', doi_tuong_values),
        ('C', hoc_ham_values),
        ('D', hoc_vi_values),
        ('E', chuc_vu_values),
        ('F', bool_values),
    ]
    for col, values in data_sets:
        for idx, value in enumerate(values, start=1):
            lookup[f'{col}{idx}'] = value

    headers = [
        'ho_ten', 'cap_bac', 'doi_tuong', 'chuc_danh', 'chuc_vu', 'can_cuoc_cong_dan',
        'don_vi_truc_thuoc',
        'ngay_sinh', 'ngay_nhap_ngu', 'hoc_ham', 'hoc_vi', 'trinh_do_hoc_van',
        'ngoai_ngu', 'la_chi_huy', 'la_bi_thu',
    ]
    ws.append(headers)
    ws.append([
        'Nguyễn Văn A', 'Trung úy', 'Giảng viên', 'Giảng viên', 'Trợ lý', '012345678901',
        'Đại đội 1', '1990-01-15', '09/2015', 'Không', 'Thạc sĩ', '12/12',
        'Anh B2', 1, 0,
    ])

    max_row = 500
    validations = [
        ('B', f"=DanhMuc!$A$1:$A${max(1, len(cap_bac_values))}"),
        ('C', f"=DanhMuc!$B$1:$B${max(1, len(doi_tuong_values))}"),
        ('J', f"=DanhMuc!$C$1:$C${max(1, len(hoc_ham_values))}"),
        ('K', f"=DanhMuc!$D$1:$D${max(1, len(hoc_vi_values))}"),
        ('E', f"=DanhMuc!$E$1:$E${max(1, len(chuc_vu_values))}"),
        ('N', "=DanhMuc!$F$1:$F$2"),
        ('O', "=DanhMuc!$F$1:$F$2"),
    ]
    for col, formula in validations:
        dv = DataValidation(type='list', formula1=formula, allow_blank=True)
        dv.prompt = 'Chọn giá trị từ danh sách'
        ws.add_data_validation(dv)
        dv.add(f'{col}2:{col}{max_row}')

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='mau_quan_nhan.xlsx',
    )


@personnel_bp.route('/import', methods=['POST'])
@login_required
@unit_user_required
def import_personnel_excel():
    if not current_user.don_vi:
        flash('Tài khoản chưa được gán đơn vị.', 'warning')
        return redirect(url_for('personnel.list_personnel'))

    file = request.files.get('excel_file')
    if not file or not file.filename:
        flash('Vui lòng chọn file Excel (.xlsx).', 'danger')
        return redirect(url_for('personnel.list_personnel'))

    if not file.filename.lower().endswith('.xlsx'):
        flash('Chỉ hỗ trợ file .xlsx.', 'danger')
        return redirect(url_for('personnel.list_personnel'))

    try:
        wb = load_workbook(file, data_only=True)
    except Exception:
        flash('Không thể đọc file Excel. Vui lòng kiểm tra định dạng.', 'danger')
        return redirect(url_for('personnel.list_personnel'))

    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        flash('File Excel trống.', 'warning')
        return redirect(url_for('personnel.list_personnel'))

    header_row = [str(c).strip() if c is not None else '' for c in rows[0]]
    headers = {h.lower(): idx for idx, h in enumerate(header_row) if h}

    required_cols = ['ho_ten']
    optional_cols = [
        'cap_bac', 'doi_tuong', 'chuc_danh', 'chuc_vu', 'can_cuoc_cong_dan',
        'don_vi_truc_thuoc',
        'ngay_sinh', 'ngay_nhap_ngu', 'hoc_ham', 'hoc_vi', 'trinh_do_hoc_van',
        'ngoai_ngu', 'la_chi_huy', 'la_bi_thu',
    ]
    missing = [c for c in required_cols if c not in headers]
    if missing:
        flash('Thiếu cột bắt buộc trong file Excel: ' + ', '.join(missing), 'danger')
        return redirect(url_for('personnel.list_personnel'))

    created = 0
    skipped = 0
    errors = 0
    replaced = 0

    cap_bac_values = set(_get_cap_bac_list())
    doi_tuong_values = {e.value for e in DoiTuong}
    hoc_ham_values = {e.value for e in HocHam}
    hoc_vi_values = {e.value for e in HocVi}

    for row in rows[1:]:
        try:
            ho_ten = row[headers['ho_ten']] if headers.get('ho_ten') is not None else None
            ho_ten = str(ho_ten).strip() if ho_ten is not None else ''
            if not ho_ten:
                skipped += 1
                continue

            cccd = None
            if 'can_cuoc_cong_dan' in headers:
                cccd = _normalize_cccd(row[headers['can_cuoc_cong_dan']])

            if cccd:
                duplicates = QuanNhan.query.filter_by(
                    don_vi_id=current_user.don_vi_id,
                    can_cuoc_cong_dan=cccd,
                    is_active=True
                ).all()
                if duplicates:
                    for duplicated_qn in duplicates:
                        duplicated_qn.is_active = False
                    replaced += len(duplicates)

            cap_bac_val = str(row[headers['cap_bac']]).strip() if headers.get('cap_bac') is not None and row[headers['cap_bac']] is not None else None
            doi_tuong_val = str(row[headers['doi_tuong']]).strip() if headers.get('doi_tuong') is not None and row[headers['doi_tuong']] is not None else None
            hoc_ham_val = str(row[headers['hoc_ham']]).strip() if headers.get('hoc_ham') is not None and row[headers['hoc_ham']] is not None else 'Không'
            hoc_vi_val = str(row[headers['hoc_vi']]).strip() if headers.get('hoc_vi') is not None and row[headers['hoc_vi']] is not None else 'Không'

            if cap_bac_val and cap_bac_val not in cap_bac_values:
                errors += 1
                continue
            if doi_tuong_val and doi_tuong_val not in doi_tuong_values:
                errors += 1
                continue
            if hoc_ham_val and hoc_ham_val not in hoc_ham_values:
                errors += 1
                continue
            if hoc_vi_val and hoc_vi_val not in hoc_vi_values:
                errors += 1
                continue

            qn = QuanNhan(
                don_vi_id=current_user.don_vi_id,
                ho_ten=ho_ten,
                cap_bac=cap_bac_val,
                doi_tuong=doi_tuong_val,
                chuc_danh=str(row[headers['chuc_danh']]).strip() if headers.get('chuc_danh') is not None and row[headers['chuc_danh']] is not None else None,
                chuc_vu=str(row[headers['chuc_vu']]).strip() if headers.get('chuc_vu') is not None and row[headers['chuc_vu']] is not None else None,
                can_cuoc_cong_dan=cccd or None,
                don_vi_truc_thuoc=str(row[headers['don_vi_truc_thuoc']]).strip() if headers.get('don_vi_truc_thuoc') is not None and row[headers['don_vi_truc_thuoc']] is not None else None,
                ngay_sinh=_parse_date(row[headers['ngay_sinh']]) if headers.get('ngay_sinh') is not None else None,
                ngay_nhap_ngu=str(row[headers['ngay_nhap_ngu']]).strip() if headers.get('ngay_nhap_ngu') is not None and row[headers['ngay_nhap_ngu']] is not None else None,
                hoc_ham=hoc_ham_val,
                hoc_vi=hoc_vi_val,
                trinh_do_hoc_van=str(row[headers['trinh_do_hoc_van']]).strip() if headers.get('trinh_do_hoc_van') is not None and row[headers['trinh_do_hoc_van']] is not None else None,
                ngoai_ngu=str(row[headers['ngoai_ngu']]).strip() if headers.get('ngoai_ngu') is not None and row[headers['ngoai_ngu']] is not None else None,
                la_chi_huy=_parse_bool(row[headers['la_chi_huy']]) if headers.get('la_chi_huy') is not None else False,
                la_bi_thu=_parse_bool(row[headers['la_bi_thu']]) if headers.get('la_bi_thu') is not None else False,
            )
            db.session.add(qn)
            created += 1
        except Exception:
            errors += 1

    if created > 0:
        db.session.commit()

    flash(
        f'Đã nhập {created} quân nhân. Bỏ qua: {skipped}. Thay thế theo CCCD: {replaced}. Lỗi: {errors}.',
        'success' if created > 0 else 'warning'
    )
    return redirect(url_for('personnel.list_personnel'))


@personnel_bp.route('/evaluations')
@login_required
@unit_user_required
def annual_evaluations():
    if not current_user.don_vi:
        flash('Tài khoản chưa được gán đơn vị.', 'warning')
        return redirect(url_for('dashboard.index'))

    nam_hoc = request.args.get('nam_hoc', '').strip()
    if not nam_hoc:
        current_year = datetime.now().year
        nam_hoc = f'{current_year}-{current_year + 1}'

    personnel = QuanNhan.query.filter_by(
        don_vi_id=current_user.don_vi_id,
        is_active=True
    ).order_by(QuanNhan.ho_ten.asc()).all()

    qn_ids = [qn.id for qn in personnel]
    existing_rows = {}
    nam_hoc_list = []
    migration_missing = False
    try:
        if qn_ids:
            rows = DanhGiaHangNam.query.filter(
                DanhGiaHangNam.nam_hoc == nam_hoc,
                DanhGiaHangNam.quan_nhan_id.in_(qn_ids)
            ).all()
            existing_rows = {row.quan_nhan_id: row for row in rows}

        nam_hoc_list = [row[0] for row in db.session.query(DanhGiaHangNam.nam_hoc).filter(
            DanhGiaHangNam.don_vi_id == current_user.don_vi_id
        ).distinct().order_by(DanhGiaHangNam.nam_hoc.desc()).all()]
    except (ProgrammingError, OperationalError):
        db.session.rollback()
        migration_missing = True
        flash('Thiếu bảng đánh giá hằng năm. Vui lòng chạy: flask db upgrade', 'danger')

    if nam_hoc not in nam_hoc_list:
        nam_hoc_list = [nam_hoc] + nam_hoc_list

    return render_template('personnel/evaluation_list.html',
                           personnel=personnel,
                           existing_rows=existing_rows,
                           nam_hoc=nam_hoc,
                           nam_hoc_list=nam_hoc_list,
                           migration_missing=migration_missing,
                           dang_vien_choices=DanhGiaHangNam.XEP_LOAI_DANG_VIEN_CHOICES,
                           can_bo_choices=DanhGiaHangNam.XEP_LOAI_CAN_BO_CHOICES)


@personnel_bp.route('/evaluations/save', methods=['POST'])
@login_required
@unit_user_required
def save_annual_evaluations():
    if not current_user.don_vi:
        flash('Tài khoản chưa được gán đơn vị.', 'warning')
        return redirect(url_for('dashboard.index'))

    nam_hoc = request.form.get('nam_hoc', '').strip()
    if not nam_hoc:
        flash('Năm học không được để trống.', 'danger')
        return redirect(url_for('personnel.annual_evaluations'))

    personnel = QuanNhan.query.filter_by(don_vi_id=current_user.don_vi_id, is_active=True).all()
    if not personnel:
        flash('Đơn vị chưa có quân nhân để đánh giá.', 'warning')
        return redirect(url_for('personnel.annual_evaluations', nam_hoc=nam_hoc))

    qn_ids = [qn.id for qn in personnel]
    try:
        existing_rows = DanhGiaHangNam.query.filter(
            DanhGiaHangNam.nam_hoc == nam_hoc,
            DanhGiaHangNam.quan_nhan_id.in_(qn_ids)
        ).all()
    except (ProgrammingError, OperationalError):
        db.session.rollback()
        flash('Thiếu bảng đánh giá hằng năm. Vui lòng chạy: flask db upgrade', 'danger')
        return redirect(url_for('personnel.annual_evaluations', nam_hoc=nam_hoc))
    existing_map = {row.quan_nhan_id: row for row in existing_rows}

    apply_all_dang_vien = request.form.get('bulk_xep_loai_dang_vien', '').strip()
    apply_all_can_bo = request.form.get('bulk_xep_loai_can_bo', '').strip()

    if apply_all_dang_vien and apply_all_dang_vien not in DanhGiaHangNam.XEP_LOAI_DANG_VIEN_CHOICES:
        flash('Giá trị áp dụng hàng loạt (đảng viên) không hợp lệ.', 'danger')
        return redirect(url_for('personnel.annual_evaluations', nam_hoc=nam_hoc))

    if apply_all_can_bo and apply_all_can_bo not in DanhGiaHangNam.XEP_LOAI_CAN_BO_CHOICES:
        flash('Giá trị áp dụng hàng loạt (cán bộ) không hợp lệ.', 'danger')
        return redirect(url_for('personnel.annual_evaluations', nam_hoc=nam_hoc))

    values_map = {}
    missing_names = []
    for qn in personnel:
        dang_vien_val = request.form.get(f'xep_loai_dang_vien_{qn.id}', '').strip() or apply_all_dang_vien
        can_bo_val = request.form.get(f'xep_loai_can_bo_{qn.id}', '').strip() or apply_all_can_bo

        if not dang_vien_val or not can_bo_val:
            missing_names.append(qn.ho_ten)
            continue

        if dang_vien_val not in DanhGiaHangNam.XEP_LOAI_DANG_VIEN_CHOICES:
            flash(f'Giá trị xếp loại đảng viên không hợp lệ cho {qn.ho_ten}.', 'danger')
            return redirect(url_for('personnel.annual_evaluations', nam_hoc=nam_hoc))
        if can_bo_val not in DanhGiaHangNam.XEP_LOAI_CAN_BO_CHOICES:
            flash(f'Giá trị xếp loại cán bộ không hợp lệ cho {qn.ho_ten}.', 'danger')
            return redirect(url_for('personnel.annual_evaluations', nam_hoc=nam_hoc))

        values_map[qn.id] = (dang_vien_val, can_bo_val)

    if missing_names:
        flash(f'Chưa đánh giá đủ toàn bộ quân nhân. Thiếu: {len(missing_names)} người.', 'danger')
        return redirect(url_for('personnel.annual_evaluations', nam_hoc=nam_hoc))

    saved = 0
    for qn in personnel:
        dang_vien_val, can_bo_val = values_map[qn.id]

        row = existing_map.get(qn.id)
        if not row:
            row = DanhGiaHangNam.query.filter_by(
                quan_nhan_id=qn.id,
                nam_hoc=nam_hoc,
            ).first()
        if not row:
            row = DanhGiaHangNam(
                quan_nhan_id=qn.id,
                don_vi_id=current_user.don_vi_id,
                nam_hoc=nam_hoc,
            )
            db.session.add(row)

        row.xep_loai_dang_vien = dang_vien_val
        row.xep_loai_can_bo = can_bo_val
        row.nguoi_cap_nhat_id = current_user.id
        saved += 1

    try:
        db.session.commit()
    except IntegrityError:
        db.session.rollback()

        # Retry safely in case another request inserted same (quan_nhan_id, nam_hoc)
        saved = 0
        for qn in personnel:
            dang_vien_val, can_bo_val = values_map[qn.id]
            row = DanhGiaHangNam.query.filter_by(
                quan_nhan_id=qn.id,
                nam_hoc=nam_hoc,
            ).first()
            if not row:
                row = DanhGiaHangNam(
                    quan_nhan_id=qn.id,
                    don_vi_id=current_user.don_vi_id,
                    nam_hoc=nam_hoc,
                )
                db.session.add(row)

            row.xep_loai_dang_vien = dang_vien_val
            row.xep_loai_can_bo = can_bo_val
            row.nguoi_cap_nhat_id = current_user.id
            saved += 1

        try:
            db.session.commit()
        except IntegrityError:
            db.session.rollback()
            flash('Có xung đột dữ liệu khi lưu đánh giá. Vui lòng tải lại trang và lưu lại.', 'danger')
            return redirect(url_for('personnel.annual_evaluations', nam_hoc=nam_hoc))

    flash(f'Đã lưu {saved} đánh giá năm học {nam_hoc}.', 'success')
    return redirect(url_for('personnel.annual_evaluations', nam_hoc=nam_hoc))
