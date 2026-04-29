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
    """List all nominations at PHE_DUYET_CUOI stage (Bảng 2) for Hội đồng voting."""
    nominations = DeXuat.query.filter_by(
        trang_thai=TrangThaiDeXuat.PHE_DUYET_CUOI.value
    ).order_by(DeXuat.ngay_gui.desc()).all()

    vai_tro = current_user.hoi_dong_vai_tro

    # For each nomination, build per-individual vote status for current user's vai_tro
    vote_summary = {}  # de_xuat_id -> {ct_id -> bieu_quyet or None}
    for dx in nominations:
        ct_votes = {}
        for ct in dx.chi_tiets:
            bq = HoiDongBieuQuyet.query.filter_by(
                chi_tiet_id=ct.id, vai_tro=vai_tro
            ).first()
            ct_votes[ct.id] = bq
        vote_summary[dx.id] = ct_votes

    return render_template('hoi_dong/list.html',
                           nominations=nominations,
                           vote_summary=vote_summary,
                           vai_tro=vai_tro,
                           VAI_TRO_DISPLAY=HOI_DONG_VAI_TRO_DISPLAY)


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

    # Build vote map: ct_id -> {vai_tro -> HoiDongBieuQuyet}
    all_votes = HoiDongBieuQuyet.query.filter_by(de_xuat_id=id).all()
    vote_map = {}  # ct_id -> {vai_tro: bq}
    for bq in all_votes:
        vote_map.setdefault(bq.chi_tiet_id, {})[bq.vai_tro] = bq

    # My votes for this nomination
    my_votes = {bq.chi_tiet_id: bq for bq in all_votes if bq.vai_tro == vai_tro}

    return render_template('hoi_dong/detail.html',
                           de_xuat=de_xuat,
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
    return redirect(url_for('hoi_dong.detail', id=id))


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
