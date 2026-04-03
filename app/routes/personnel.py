from flask import Blueprint, render_template, redirect, url_for, flash, request
from flask_login import login_required, current_user
from app.extensions import db
from app.models.user import Role
from app.models.personnel import QuanNhan, CapBac, HocHam, HocVi, DoiTuong
from app.models.certificate import ChungChi, LoaiChungChi
from app.utils.decorators import unit_user_required
from app.utils.file_upload import save_upload, delete_upload
from datetime import datetime

personnel_bp = Blueprint('personnel', __name__)


@personnel_bp.route('/')
@login_required
@unit_user_required
def list_personnel():
    if not current_user.don_vi:
        flash('Tài khoản chưa được gán đơn vị.', 'warning')
        return redirect(url_for('dashboard.index'))

    search = request.args.get('search', '').strip()
    page = request.args.get('page', 1, type=int)

    query = QuanNhan.query.filter_by(don_vi_id=current_user.don_vi_id, is_active=True)
    if search:
        query = query.filter(QuanNhan.ho_ten.ilike(f'%{search}%'))

    personnel = query.order_by(QuanNhan.ho_ten).paginate(page=page, per_page=20, error_out=False)

    return render_template('personnel/list.html',
                           personnel=personnel, search=search,
                           cap_bac_list=[e.value for e in CapBac],
                           doi_tuong_list=[e.value for e in DoiTuong])


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
                                   cap_bac_list=[e.value for e in CapBac],
                                   hoc_ham_list=[e.value for e in HocHam],
                                   hoc_vi_list=[e.value for e in HocVi],
                                   doi_tuong_list=[e.value for e in DoiTuong])

        db.session.add(qn)
        db.session.commit()
        flash(f'Đã thêm quân nhân: {qn.ho_ten}', 'success')
        return redirect(url_for('personnel.detail_personnel', id=qn.id))

    return render_template('personnel/create.html',
                           cap_bac_list=[e.value for e in CapBac],
                           hoc_ham_list=[e.value for e in HocHam],
                           hoc_vi_list=[e.value for e in HocVi],
                           doi_tuong_list=[e.value for e in DoiTuong])


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
                           cap_bac_list=[e.value for e in CapBac],
                           hoc_ham_list=[e.value for e in HocHam],
                           hoc_vi_list=[e.value for e in HocVi],
                           doi_tuong_list=[e.value for e in DoiTuong])


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
