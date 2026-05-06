from flask import Blueprint, render_template, redirect, url_for, flash, request, jsonify
from flask_login import login_required, current_user
from app.extensions import db
from app.models.nomination import DeXuat, DeXuatChiTiet, TrangThaiDeXuat
from app.models.hoi_dong import HoiDongBieuQuyet, HOI_DONG_VAI_TRO, HOI_DONG_VAI_TRO_DISPLAY, KET_QUA_DONG_Y, KET_QUA_KHONG_DONG_Y
from app.utils.decorators import hoi_dong_required
from datetime import datetime

hoi_dong_bp = Blueprint('hoi_dong', __name__)


@hoi_dong_bp.route('/')
@login_required
@hoi_dong_required
def list_nominations():
    """Flat per-individual list of all PHE_DUYET_CUOI individuals for Hội đồng voting,
    sorted by DanhHieu.thu_tu then unit name."""
    from app.models.nomination import DanhHieu
    from app.models.unit import DonVi

    nam_hoc_filter = request.args.get('nam_hoc', '')

    # Get available nam_hoc options
    nam_hoc_list = [n[0] for n in db.session.query(DeXuat.nam_hoc).filter_by(
        trang_thai=TrangThaiDeXuat.PHE_DUYET_CUOI.value
    ).distinct().order_by(DeXuat.nam_hoc.desc()).all()]

    rows = []
    if nam_hoc_filter:
        q = DeXuat.query.filter_by(trang_thai=TrangThaiDeXuat.PHE_DUYET_CUOI.value)\
            .filter(DeXuat.nam_hoc == nam_hoc_filter)\
            .join(DonVi, DeXuat.don_vi_id == DonVi.id).order_by(DonVi.ten_don_vi)
        nominations = q.all()

        vai_tro = current_user.hoi_dong_vai_tro
        danh_hieu_order = {dh.ten_danh_hieu: dh.thu_tu for dh in DanhHieu.query.all()}

        for dx in nominations:
            for ct in dx.chi_tiets:
                my_vote = HoiDongBieuQuyet.query.filter_by(
                    chi_tiet_id=ct.id, vai_tro=vai_tro
                ).first()
                all_votes = HoiDongBieuQuyet.query.filter_by(chi_tiet_id=ct.id).all()
                dong_y_count = sum(1 for v in all_votes if v.ket_qua == KET_QUA_DONG_Y)
                rows.append({
                    'dx': dx,
                    'ct': ct,
                    'my_vote': my_vote,
                    'dong_y_count': dong_y_count,
                    'total_vai_tro': len(HOI_DONG_VAI_TRO),
                    '_sort_award': danh_hieu_order.get(ct.loai_danh_hieu, 999),
                    '_sort_unit': dx.don_vi.ten_don_vi if dx.don_vi else '',
                })
        rows.sort(key=lambda r: (r['_sort_award'], r['_sort_unit']))
    else:
        vai_tro = current_user.hoi_dong_vai_tro

    return render_template('hoi_dong/list.html',
                           rows=rows,
                           vai_tro=vai_tro,
                           VAI_TRO_DISPLAY=HOI_DONG_VAI_TRO_DISPLAY,
                           KET_QUA_DONG_Y=KET_QUA_DONG_Y,
                           KET_QUA_KHONG_DONG_Y=KET_QUA_KHONG_DONG_Y,
                           nam_hoc_filter=nam_hoc_filter,
                           nam_hoc_list=nam_hoc_list)


@hoi_dong_bp.route('/<int:id>')
@login_required
@hoi_dong_required
def detail(id):
    """Detail view + voting UI for one nomination."""
    de_xuat = DeXuat.query.get_or_404(id)

    if de_xuat.trang_thai != TrangThaiDeXuat.PHE_DUYET_CUOI.value:
        flash('Đề xuất này không ở giai đoạn Hội đồng biểu quyết (Bảng 2).', 'warning')
        return redirect(url_for('hoi_dong.list_nominations'))

    vai_tro = current_user.hoi_dong_vai_tro

    # Sort chi_tiets by DanhHieu.thu_tu then name
    from app.models.nomination import DanhHieu
    danh_hieu_order = {dh.ten_danh_hieu: dh.thu_tu for dh in DanhHieu.query.all()}
    sorted_chi_tiets = sorted(
        de_xuat.chi_tiets,
        key=lambda ct: (danh_hieu_order.get(ct.loai_danh_hieu, 999),
                        ct.quan_nhan.ho_ten if ct.quan_nhan else ct.ten_don_vi_de_xuat or '')
    )

    # Build vote map: ct_id -> {vai_tro -> HoiDongBieuQuyet}
    all_votes = HoiDongBieuQuyet.query.filter_by(de_xuat_id=id).all()
    vote_map = {}  # ct_id -> {vai_tro: bq}
    for bq in all_votes:
        vote_map.setdefault(bq.chi_tiet_id, {})[bq.vai_tro] = bq

    # My votes for this nomination
    my_votes = {bq.chi_tiet_id: bq for bq in all_votes if bq.vai_tro == vai_tro}

    return render_template('hoi_dong/detail.html',
                           de_xuat=de_xuat,
                           sorted_chi_tiets=sorted_chi_tiets,
                           vai_tro=vai_tro,
                           all_vai_tro=HOI_DONG_VAI_TRO,
                           VAI_TRO_DISPLAY=HOI_DONG_VAI_TRO_DISPLAY,
                           vote_map=vote_map,
                           my_votes=my_votes,
                           KET_QUA_DONG_Y=KET_QUA_DONG_Y,
                           KET_QUA_KHONG_DONG_Y=KET_QUA_KHONG_DONG_Y)


@hoi_dong_bp.route('/<int:id>/vote/<int:ct_id>', methods=['POST'])
@login_required
@hoi_dong_required
def cast_vote(id, ct_id):
    """Cast or update a vote for one individual."""
    de_xuat = DeXuat.query.get_or_404(id)
    ct = DeXuatChiTiet.query.get_or_404(ct_id)

    if ct.de_xuat_id != id:
        flash('Dữ liệu không hợp lệ.', 'danger')
        return redirect(url_for('hoi_dong.detail', id=id))

    if de_xuat.trang_thai != TrangThaiDeXuat.PHE_DUYET_CUOI.value:
        flash('Đề xuất này không ở giai đoạn Hội đồng.', 'warning')
        return redirect(url_for('hoi_dong.detail', id=id))

    vai_tro = current_user.hoi_dong_vai_tro
    ket_qua = request.form.get('ket_qua', '').strip()
    ghi_chu = request.form.get('ghi_chu', '').strip() or None

    if ket_qua not in (KET_QUA_DONG_Y, KET_QUA_KHONG_DONG_Y):
        flash('Kết quả biểu quyết không hợp lệ.', 'danger')
        return redirect(url_for('hoi_dong.detail', id=id))

    existing = HoiDongBieuQuyet.query.filter_by(
        chi_tiet_id=ct_id, vai_tro=vai_tro
    ).first()

    if existing:
        existing.ket_qua = ket_qua
        existing.ghi_chu = ghi_chu
        existing.nguoi_bieu_quyet_id = current_user.id
        existing.updated_at = datetime.utcnow()
    else:
        bq = HoiDongBieuQuyet(
            de_xuat_id=id,
            chi_tiet_id=ct_id,
            nguoi_bieu_quyet_id=current_user.id,
            vai_tro=vai_tro,
            ket_qua=ket_qua,
            ghi_chu=ghi_chu,
        )
        db.session.add(bq)

    db.session.commit()

    name = ct.quan_nhan.ho_ten if ct.quan_nhan else ct.ten_don_vi_de_xuat or 'Đơn vị'
    flash(f'Đã biểu quyết "{ket_qua}" cho {name}.', 'success')
    # Redirect back to the referring page (list or detail), default to list
    ref = request.referrer or ''
    if f'/hoi-dong/{id}' in ref and 'vote' not in ref:
        return redirect(url_for('hoi_dong.detail', id=id))
    return redirect(url_for('hoi_dong.list_nominations'))


@hoi_dong_bp.route('/<int:id>/vote-all', methods=['POST'])
@login_required
@hoi_dong_required
def cast_vote_all(id):
    """Batch vote: cast the same result for all chi_tiets in nomination."""
    de_xuat = DeXuat.query.get_or_404(id)

    if de_xuat.trang_thai != TrangThaiDeXuat.PHE_DUYET_CUOI.value:
        flash('Đề xuất này không ở giai đoạn Hội đồng.', 'warning')
        return redirect(url_for('hoi_dong.detail', id=id))

    vai_tro = current_user.hoi_dong_vai_tro
    ket_qua = request.form.get('ket_qua', '').strip()
    ghi_chu = request.form.get('ghi_chu', '').strip() or None

    if ket_qua not in (KET_QUA_DONG_Y, KET_QUA_KHONG_DONG_Y):
        flash('Kết quả biểu quyết không hợp lệ.', 'danger')
        return redirect(url_for('hoi_dong.detail', id=id))

    now = datetime.utcnow()
    count = 0
    for ct in de_xuat.chi_tiets:
        existing = HoiDongBieuQuyet.query.filter_by(
            chi_tiet_id=ct.id, vai_tro=vai_tro
        ).first()
        if existing:
            existing.ket_qua = ket_qua
            existing.ghi_chu = ghi_chu
            existing.nguoi_bieu_quyet_id = current_user.id
            existing.updated_at = now
        else:
            bq = HoiDongBieuQuyet(
                de_xuat_id=id,
                chi_tiet_id=ct.id,
                nguoi_bieu_quyet_id=current_user.id,
                vai_tro=vai_tro,
                ket_qua=ket_qua,
                ghi_chu=ghi_chu,
            )
            db.session.add(bq)
        count += 1

    db.session.commit()
    flash(f'Đã biểu quyết "{ket_qua}" cho {count} cá nhân/đơn vị.', 'success')
    return redirect(url_for('hoi_dong.detail', id=id))
