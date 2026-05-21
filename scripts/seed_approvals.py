# -*- coding: utf-8 -*-
"""Seed 10 approval records. Run: python scripts/seed_approvals.py"""
import sys, os, random
from datetime import datetime, timedelta
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import create_app
app = create_app()

with app.app_context():
    from app.models.nomination import DeXuat, DeXuatChiTiet, TrangThaiDeXuat
    from app.models.approval import PheDuyet, KetQuaDuyetChiTiet, KetQuaDuyet
    from app.routes.approval import ROLE_TO_PHONG
    from app.extensions import db

    REAL_DEPTS = list(set(ROLE_TO_PHONG.values()))

    # Clear all
    KetQuaDuyetChiTiet.query.delete()
    PheDuyet.query.delete()
    DeXuat.query.update({DeXuat.trang_thai: TrangThaiDeXuat.NHAP.value})
    DeXuatChiTiet.query.update({DeXuatChiTiet.admin_approved: False})
    db.session.commit()

    # IDs with chi_tiets (verified earlier)
    dx_ids = [300, 301, 304, 308, 309, 310, 311, 313, 317, 318]
    random.seed(42)
    random.shuffle(dx_ids)

    # (n_depts or 'reject', final_status_value)
    scenarios = [
        (2,        TrangThaiDeXuat.DANG_DUYET.value),
        (4,        TrangThaiDeXuat.DANG_DUYET.value),
        (7,        TrangThaiDeXuat.DANG_DUYET.value),
        (10,       TrangThaiDeXuat.DANG_DUYET.value),
        (12,       TrangThaiDeXuat.DANG_DUYET.value),
        (15,       TrangThaiDeXuat.HOI_DONG.value),
        (15,       TrangThaiDeXuat.HOI_DONG.value),
        (15,       TrangThaiDeXuat.HOI_DONG.value),
        ('reject', TrangThaiDeXuat.TU_CHOI.value),
        ('reject', TrangThaiDeXuat.TU_CHOI.value),
    ]

    # Admin dept stored as 'Tuyen huan' — get exact value
    from app.models.approval import PhongDuyet
    ADMIN_PHONG = PhongDuyet.ADMIN_TUYENHUAN.value

    t_now = datetime.utcnow()

    for dx_id, (n_depts, status_val) in zip(dx_ids, scenarios):
        dx = DeXuat.query.filter_by(id=dx_id).first()
        if not dx:
            continue
        chi_tiets = DeXuatChiTiet.query.filter_by(de_xuat_id=dx_id).all()
        if not chi_tiets:
            continue

        depts = REAL_DEPTS[:]
        random.shuffle(depts)

        if n_depts == 'reject':
            use_depts = depts[:random.randint(2, 5)]
        else:
            use_depts = depts[:n_depts] if n_depts < len(depts) else depts

        for dept_i, dept in enumerate(use_depts):
            is_reject = (n_depts == 'reject' and dept == use_depts[-1])
            pd_kq = KetQuaDuyet.TU_CHOI.value if is_reject else KetQuaDuyet.DONG_Y.value
            pd = PheDuyet(
                de_xuat_id=dx_id,
                phong_duyet=dept,
                ket_qua=pd_kq,
                ly_do=('Khong dat yeu cau' if is_reject else None),
                ngay_duyet=t_now - timedelta(days=random.randint(1, 15), hours=dept_i),
            )
            db.session.add(pd)
            db.session.flush()

            for ct in chi_tiets:
                kq = KetQuaDuyetChiTiet(
                    phe_duyet_id=pd.id,
                    chi_tiet_id=ct.id,
                    ket_qua=(KetQuaDuyet.TU_CHOI.value if is_reject else KetQuaDuyet.DONG_Y.value),
                    ly_do=('Khong dat' if is_reject else None),
                )
                db.session.add(kq)

        # For HOI_DONG: add admin PheDuyet pending
        if status_val == TrangThaiDeXuat.HOI_DONG.value:
            admin_pd = PheDuyet(
                de_xuat_id=dx_id,
                phong_duyet=ADMIN_PHONG,
                ket_qua=KetQuaDuyet.CHO_DUYET.value,
            )
            db.session.add(admin_pd)

        dx.trang_thai = status_val
        db.session.flush()

    db.session.commit()

    pd_count = PheDuyet.query.count()
    kq_count = KetQuaDuyetChiTiet.query.count()
    print('Done! PheDuyet=%d KetQua=%d' % (pd_count, kq_count))
    for dx in DeXuat.query.filter(DeXuat.trang_thai != TrangThaiDeXuat.NHAP.value).all():
        print('  DX %d -> %s' % (dx.id, dx.trang_thai))
