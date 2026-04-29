"""
Reset data and generate sample dataset with 20 nominations covering the full 3-table flow:
  Bảng 1  : HOI_DONG status, all depts approved, admin_approved=False  → 6 nominations
  Bảng 2  : PHE_DUYET_CUOI status, admin_approved=True, Hội đồng votes  → 4 nominations
  Bảng 3  : KhenThuong confirmed                                         → 4 nominations
  Other   : Nháp (2), Chờ duyệt (2), Đang duyệt (2)

Usage:
    python seed_sample.py
"""

import random
from datetime import date, datetime, timedelta

from app import create_app
from app.extensions import db
from app.models.user import User, Role
from app.models.unit import DonVi
from app.models.personnel import QuanNhan, DoiTuong, MucDoHoanThanh
from app.models.certificate import ChungChi
from app.models.nomination import DeXuat, DeXuatChiTiet, TrangThaiDeXuat, DanhHieu
from app.models.approval import PheDuyet, PhongDuyet, KetQuaDuyet, KetQuaDuyetChiTiet
from app.models.reward import KhenThuong
from app.models.notification import ThongBao
from app.models.evaluation import DanhGiaHangNam
from app.models.hoi_dong import HoiDongBieuQuyet, HOI_DONG_VAI_TRO, KET_QUA_DONG_Y, KET_QUA_KHONG_DONG_Y

TARGET_PERSONNEL = 60
TARGET_NOMINATIONS = 20

HO = ['Nguyen', 'Tran', 'Le', 'Pham', 'Hoang', 'Phan', 'Vo', 'Dang', 'Bui', 'Do']
DEM = ['Van', 'Thi', 'Duc', 'Minh', 'Quoc', 'Xuan', 'Ngoc', 'Thanh', 'Huu', 'Duy']
TEN = ['An', 'Binh', 'Cuong', 'Dung', 'Hung', 'Hieu', 'Khanh', 'Long',
       'Nam', 'Phong', 'Quan', 'Son', 'Tai', 'Tung', 'Vinh', 'Dat',
       'Hoa', 'Lam', 'Manh', 'Nghia']

CAP_BAC_MAIN = ['Thieu uy', 'Trung uy', 'Thuong uy', 'Dai uy', 'Thieu ta', 'Trung ta']
CAP_BAC_STUDENT = ['Binh nhi', 'Binh nhat', 'Ha si', 'Trung si']
CHUC_VU_GV = ['Giang vien', 'Giang vien chinh', 'Pho Truong bo mon', 'Truong bo mon']
CHUC_VU_CB = ['Tro ly', 'Nhan vien', 'Pho truong ban', 'Truong ban']

DEPT_FLOW = [
    PhongDuyet.PHONG_KHOAHOC,
    PhongDuyet.PHONG_DAOTAO,
    PhongDuyet.BAN_CANBO,
    PhongDuyet.BAN_TOCHUC,
    PhongDuyet.BAN_TUYENHUAN,
    PhongDuyet.BAN_CTCQ,
    PhongDuyet.BAN_CNTT,
    PhongDuyet.BAN_TAC_HUAN,
    PhongDuyet.BAN_KHAOTHI,
    PhongDuyet.UY_BAN_KIEMTRA,
    PhongDuyet.BAN_QUANLUC,
    PhongDuyet.THU_TRUONG_PHONG_CHINHTRI,
    PhongDuyet.THU_TRUONG_PHONG_TMHC,
]

BAN_QUANLUC_SCOPE = {DoiTuong.CNV.value, DoiTuong.QNCN.value}


def random_name():
    return f"{random.choice(HO)} {random.choice(DEM)} {random.choice(TEN)}"


def random_date(start_year, end_year):
    start = date(start_year, 1, 1)
    end = date(end_year, 12, 31)
    return start + timedelta(days=random.randint(0, (end - start).days))


def random_cccd(index):
    return f"{920000000000 + index:012d}"


def in_scope(dept: PhongDuyet, doi_tuong: str):
    if dept == PhongDuyet.BAN_QUANLUC:
        return doi_tuong in BAN_QUANLUC_SCOPE
    if dept == PhongDuyet.BAN_CANBO:
        return doi_tuong not in BAN_QUANLUC_SCOPE
    return True


def ensure_department_users():
    required = {
        'admin':             (Role.ADMIN,                   'Ban thu ky Hoi dong'),
        'phong_khoahoc':     (Role.PHONG_KHOAHOC,           'Phong Khoa hoc'),
        'phong_daotao':      (Role.PHONG_DAOTAO,            'Phong Dao tao'),
        'tt_phong_chinhtri': (Role.THU_TRUONG_PHONG_CHINHTRI, 'Thu truong Phong Chinh tri'),
        'tt_phong_tmhc':     (Role.THU_TRUONG_PHONG_TMHC,  'Thu truong Phong TM-HC'),
        'ban_canbo':         (Role.BAN_CANBO,               'Ban Can bo'),
        'ban_tochuc':        (Role.BAN_TOCHUC,              'Ban To chuc'),
        'ban_tuyenhuan':     (Role.BAN_TUYENHUAN,           'Ban Tuyen huan'),
        'ban_ctcq':          (Role.BAN_CTCQ,                'Ban Cong tac quan chung'),
        'ban_cntt':          (Role.BAN_CNTT,                'Ban Cong nghe thong tin'),
        'ban_tachuan':       (Role.BAN_TAC_HUAN,            'Ban Tac huan'),
        'ban_khaothi':       (Role.BAN_KHAOTHI,             'Ban Khao thi'),
        'uyban_kiemtra':     (Role.UY_BAN_KIEMTRA,          'Uy ban Kiem tra'),
        'ban_quanluc':       (Role.BAN_QUANLUC,             'Ban Quan luc'),
    }
    for username, (role, full_name) in required.items():
        if User.query.filter_by(username=username).first():
            continue
        u = User(username=username, ho_ten=full_name, role=role, don_vi_id=None)
        u.set_password('123456')
        db.session.add(u)


def ensure_unit_users(units):
    for unit in units:
        username = unit.ma_don_vi.lower()
        existing = User.query.filter_by(username=username).first()
        if existing:
            if existing.role == Role.UNIT_USER and existing.don_vi_id != unit.id:
                existing.don_vi_id = unit.id
            continue
        u = User(username=username, ho_ten=unit.ten_don_vi,
                 role=Role.UNIT_USER, don_vi_id=unit.id)
        u.set_password('123456')
        db.session.add(u)


def ensure_award_titles():
    defaults = [
        ('Chiến sĩ thi đua',  'CSTD', 'Cá nhân', 1),
        ('Chiến sĩ tiên tiến', 'CSTT', 'Cá nhân', 2),
        ('Đơn vị quyết thắng', 'DVQT', 'Đơn vị',  3),
        ('Đơn vị tiên tiến',   'DVTT', 'Đơn vị',  4),
    ]
    for ten, ma, pham_vi, thu_tu in defaults:
        if DanhHieu.query.filter_by(ten_danh_hieu=ten).first():
            continue
        dh = DanhHieu(ten_danh_hieu=ten, ma_danh_hieu=ma, pham_vi=pham_vi,
                      thu_tu=thu_tu, is_active=True)
        dh.tieu_chi = ['muc_do_hoan_thanh', 'phieu_tin_nhiem', 'xep_loai_dang_vien', 'diem_nckh']
        db.session.add(dh)


def reset_data():
    ThongBao.query.delete()
    KhenThuong.query.delete()
    HoiDongBieuQuyet.query.delete()
    DanhGiaHangNam.query.delete()
    KetQuaDuyetChiTiet.query.delete()
    PheDuyet.query.delete()
    from app.models.nomination import MinhChung
    MinhChung.query.delete()
    DeXuatChiTiet.query.delete()
    DeXuat.query.delete()
    ChungChi.query.delete()
    QuanNhan.query.delete()
    User.query.filter(User.role == Role.UNIT_USER).delete()
    db.session.commit()


def seed_personnel(units):
    all_qn = []
    khoa_units  = [u for u in units if u.ma_don_vi.startswith('K')]
    td_units    = [u for u in units if u.ma_don_vi.startswith('TD')]
    other_units = [u for u in units if u not in khoa_units + td_units]

    for idx in range(TARGET_PERSONNEL):
        if idx < 25 and khoa_units:
            unit = random.choice(khoa_units)
            doi_tuong = random.choice([DoiTuong.GV.value, DoiTuong.CB.value, DoiTuong.CCQP.value])
        elif idx < 50 and td_units:
            unit = random.choice(td_units)
            doi_tuong = random.choice([
                DoiTuong.SV_NAM1.value, DoiTuong.SV_NAM2.value,
                DoiTuong.SV_NAM3.value, DoiTuong.SV_NAM4.value,
                DoiTuong.QNCN.value, DoiTuong.CNV.value,
            ])
        else:
            unit = random.choice(other_units or units)
            doi_tuong = random.choice([e.value for e in DoiTuong])

        is_student = doi_tuong.startswith('Học viên')
        cap_bac = random.choice(CAP_BAC_STUDENT if is_student else CAP_BAC_MAIN)

        if doi_tuong == DoiTuong.GV.value:
            chuc_vu = random.choice(CHUC_VU_GV)
            hoc_vi = random.choice(['Thạc sĩ', 'Tiến sĩ'])
        elif is_student:
            chuc_vu = random.choice(['Hoc vien', 'Lop truong'])
            hoc_vi = 'Không'
        else:
            chuc_vu = random.choice(CHUC_VU_CB)
            hoc_vi = random.choice(['Không', 'Thạc sĩ'])

        qn = QuanNhan(
            don_vi_id=unit.id,
            ho_ten=random_name(),
            cap_bac=cap_bac,
            chuc_danh=doi_tuong,
            chuc_vu=chuc_vu,
            don_vi_truc_thuoc=f"{unit.ma_don_vi}-D{random.randint(1, 4)}",
            can_cuoc_cong_dan=random_cccd(idx + 1),
            ngay_sinh=random_date(1970, 2004),
            ngay_nhap_ngu=f"{random.randint(1, 12):02d}/{random.randint(1990, 2024)}",
            doi_tuong=doi_tuong,
            hoc_ham='Không', hoc_vi=hoc_vi,
            trinh_do_hoc_van='12/12',
            ngoai_ngu=random.choice(['Anh B1', 'Anh B2', 'Anh C1']),
            la_chi_huy=False, la_bi_thu=False, is_active=True,
        )
        db.session.add(qn)
        all_qn.append(qn)
        if (idx + 1) % 10 == 0:
            db.session.commit()

    db.session.commit()

    # Use direct SQL to mark chi_huy / bi_thu to avoid stale object issues
    all_qn_fresh = QuanNhan.query.order_by(QuanNhan.id).all()
    by_unit = {}
    for qn in all_qn_fresh:
        by_unit.setdefault(qn.don_vi_id, []).append(qn)
    for _, rows in by_unit.items():
        db.session.execute(
            QuanNhan.__table__.update()
            .where(QuanNhan.__table__.c.id == rows[0].id)
            .values(la_chi_huy=True)
        )
        if len(rows) > 1:
            db.session.execute(
                QuanNhan.__table__.update()
                .where(QuanNhan.__table__.c.id == rows[1].id)
                .values(la_bi_thu=True)
            )
    db.session.commit()
    return all_qn_fresh


def make_chi_tiet(de_xuat_id, qn, nam_hoc, loai_danh_hieu, admin_approved=False):
    ct = DeXuatChiTiet(
        de_xuat_id=de_xuat_id,
        quan_nhan_id=qn.id,
        loai_danh_hieu=loai_danh_hieu,
        doi_tuong=qn.doi_tuong,
        nam_hoc=nam_hoc,
        muc_do_hoan_thanh=random.choice([
            MucDoHoanThanh.HTXSNV.value,
            MucDoHoanThanh.HTTNV.value,
            MucDoHoanThanh.HTNV.value,
        ]),
        phieu_tin_nhiem=f"{random.randint(70, 100)}%",
        xep_loai_dang_vien=random.choice([
            'Hoan thanh xuat sac nhiem vu',
            'Hoan thanh tot nhiem vu',
            'Hoan thanh nhiem vu',
        ]),
        kiem_tra_tin_hoc=random.choice(['Gioi', 'Kha', 'Dat']),
        kiem_tra_dieu_lenh=random.choice(['Gioi', 'Kha', 'Dat']),
        dia_ly_quan_su=random.choice(['Gioi', 'Kha', 'Dat']),
        ban_sung=random.choice(['Gioi', 'Kha', 'Dat']),
        the_luc=random.choice(['Gioi', 'Kha', 'Dat']),
        kiem_tra_chinh_tri=random.choice(['Gioi', 'Kha', 'Dat']),
        diem_nckh=round(random.uniform(5.5, 9.8), 1),
        nckh_noi_dung='Tham gia de tai NCKH cap truong',
        thanh_tich_ca_nhan_khac='Co thanh tich tieu bieu trong phong trao don vi',
        admin_approved=admin_approved,
    )
    return ct


def mark_all_depts_approved(dx, items, dept_reviewer, sent_at):
    """Create PheDuyet records for all depts, all DONG_Y."""
    pd_map = {}
    for dept in DEPT_FLOW:
        pd = PheDuyet(
            de_xuat_id=dx.id,
            phong_duyet=dept.value,
            ket_qua=KetQuaDuyet.DONG_Y.value,
            nguoi_duyet_id=(dept_reviewer.get(dept.value).id
                            if dept_reviewer.get(dept.value) else None),
            ngay_duyet=sent_at + timedelta(days=random.randint(1, 10)) if sent_at else datetime.utcnow(),
        )
        db.session.add(pd)
        pd_map[dept.value] = pd
    db.session.flush()

    for dept in DEPT_FLOW:
        pd = pd_map[dept.value]
        for ct in items:
            if not in_scope(dept, ct.doi_tuong):
                continue
            db.session.add(KetQuaDuyetChiTiet(
                phe_duyet_id=pd.id,
                chi_tiet_id=ct.id,
                ket_qua=KetQuaDuyet.DONG_Y.value,
            ))
    return pd_map


def seed_nominations(units, personnel):
    unit_users = {
        u.don_vi_id: u
        for u in User.query.filter_by(role=Role.UNIT_USER).all()
        if u.don_vi_id
    }
    dept_reviewer = {
        PhongDuyet.PHONG_KHOAHOC.value:             User.query.filter_by(role=Role.PHONG_KHOAHOC).first(),
        PhongDuyet.PHONG_DAOTAO.value:              User.query.filter_by(role=Role.PHONG_DAOTAO).first(),
        PhongDuyet.BAN_CANBO.value:                 User.query.filter_by(role=Role.BAN_CANBO).first(),
        PhongDuyet.BAN_TOCHUC.value:                User.query.filter_by(role=Role.BAN_TOCHUC).first(),
        PhongDuyet.BAN_TUYENHUAN.value:             User.query.filter_by(role=Role.BAN_TUYENHUAN).first(),
        PhongDuyet.BAN_CTCQ.value:                  User.query.filter_by(role=Role.BAN_CTCQ).first(),
        PhongDuyet.BAN_CNTT.value:                  User.query.filter_by(role=Role.BAN_CNTT).first(),
        PhongDuyet.BAN_TAC_HUAN.value:              User.query.filter_by(role=Role.BAN_TAC_HUAN).first(),
        PhongDuyet.BAN_KHAOTHI.value:               User.query.filter_by(role=Role.BAN_KHAOTHI).first(),
        PhongDuyet.UY_BAN_KIEMTRA.value:            User.query.filter_by(role=Role.UY_BAN_KIEMTRA).first(),
        PhongDuyet.BAN_QUANLUC.value:               User.query.filter_by(role=Role.BAN_QUANLUC).first(),
        PhongDuyet.THU_TRUONG_PHONG_CHINHTRI.value: User.query.filter_by(role=Role.THU_TRUONG_PHONG_CHINHTRI).first(),
        PhongDuyet.THU_TRUONG_PHONG_TMHC.value:     User.query.filter_by(role=Role.THU_TRUONG_PHONG_TMHC).first(),
    }
    admin = User.query.filter_by(role=Role.ADMIN).first()

    by_unit = {}
    for qn in personnel:
        by_unit.setdefault(qn.don_vi_id, []).append(qn)

    units_pool = [
        u for u in units
        if u.id in by_unit and u.id in unit_users and len(by_unit[u.id]) >= 2
    ]
    if not units_pool:
        raise RuntimeError('Not enough units with personnel and unit_user accounts.')

    awards_personal = ['Chiến sĩ thi đua', 'Chiến sĩ tiên tiến']

    # ── Plan: 20 nominations ──────────────────────────────────────────────────
    # stage: ('key', count)
    PLAN = [
        ('nhap',         2),   # Nháp
        ('cho_duyet',    2),   # Chờ duyệt
        ('dang_duyet',   2),   # Đang duyệt
        ('hoi_dong',     6),   # HOI_DONG → Bảng 1 (admin_approved=False)
        ('phe_duyet_b2', 4),   # PHE_DUYET_CUOI with partial Hội đồng votes → Bảng 2
        ('confirmed',    4),   # PHE_DUYET_CUOI + KhenThuong confirmed → Bảng 3
    ]

    unit_cycle = units_pool * 10  # ensure enough
    unit_idx = 0
    nom_count = 0

    for stage, count in PLAN:
        for _ in range(count):
            unit = unit_cycle[unit_idx % len(unit_cycle)]
            unit_idx += 1
            creator = unit_users[unit.id]
            unit_personnel = by_unit[unit.id]
            nam_hoc = random.choice(['2024-2025', '2025-2026'])
            created_at = datetime.utcnow() - timedelta(days=random.randint(10, 90))
            sent_at = created_at + timedelta(days=random.randint(1, 3))

            # ── Create DeXuat ──────────────────────────────────────────────
            if stage == 'nhap':
                trang_thai = TrangThaiDeXuat.NHAP.value
            elif stage == 'cho_duyet':
                trang_thai = TrangThaiDeXuat.CHO_DUYET.value
            elif stage == 'dang_duyet':
                trang_thai = TrangThaiDeXuat.DANG_DUYET.value
            else:
                # hoi_dong / phe_duyet_b2 / confirmed
                trang_thai = TrangThaiDeXuat.HOI_DONG.value

            dx = DeXuat(
                don_vi_id=unit.id,
                nam_hoc=nam_hoc,
                nguoi_tao_id=creator.id,
                trang_thai=trang_thai,
                ngay_gui=None if stage == 'nhap' else sent_at,
                ghi_chu='Du lieu seed',
            )
            db.session.add(dx)
            db.session.flush()

            # ── Chi tiets ─────────────────────────────────────────────────
            chosen = random.sample(unit_personnel, k=min(len(unit_personnel), random.randint(2, 3)))
            admin_approved_flag = stage in ('phe_duyet_b2', 'confirmed')
            items = []
            for qn in chosen:
                ct = make_chi_tiet(dx.id, qn, nam_hoc,
                                   random.choice(awards_personal),
                                   admin_approved=admin_approved_flag)
                db.session.add(ct)
                items.append(ct)
            db.session.flush()

            if stage == 'nhap':
                nom_count += 1
                continue

            # ── PheDuyet for stages that need it ─────────────────────────
            if stage == 'cho_duyet':
                for dept in DEPT_FLOW:
                    db.session.add(PheDuyet(
                        de_xuat_id=dx.id,
                        phong_duyet=dept.value,
                        ket_qua=KetQuaDuyet.CHO_DUYET.value,
                    ))
                nom_count += 1
                continue

            if stage == 'dang_duyet':
                reviewed = random.sample(DEPT_FLOW, k=random.randint(3, 7))
                for dept in DEPT_FLOW:
                    ket_qua = (KetQuaDuyet.DONG_Y.value if dept in reviewed
                               else KetQuaDuyet.CHO_DUYET.value)
                    reviewer_user = dept_reviewer.get(dept.value)
                    pd = PheDuyet(
                        de_xuat_id=dx.id,
                        phong_duyet=dept.value,
                        ket_qua=ket_qua,
                        nguoi_duyet_id=reviewer_user.id if (reviewer_user and ket_qua == KetQuaDuyet.DONG_Y.value) else None,
                        ngay_duyet=sent_at + timedelta(days=2) if ket_qua == KetQuaDuyet.DONG_Y.value else None,
                    )
                    db.session.add(pd)
                    db.session.flush()
                    if ket_qua == KetQuaDuyet.DONG_Y.value:
                        for ct in items:
                            if in_scope(dept, ct.doi_tuong):
                                db.session.add(KetQuaDuyetChiTiet(
                                    phe_duyet_id=pd.id,
                                    chi_tiet_id=ct.id,
                                    ket_qua=KetQuaDuyet.DONG_Y.value,
                                ))
                nom_count += 1
                continue

            # ── HOI_DONG, PHE_DUYET_B2, CONFIRMED: all depts DONG_Y ──────
            mark_all_depts_approved(dx, items, dept_reviewer, sent_at)

            if stage == 'hoi_dong':
                # Admin PheDuyet pending (not yet decided)
                db.session.add(PheDuyet(
                    de_xuat_id=dx.id,
                    phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value,
                    ket_qua=KetQuaDuyet.CHO_DUYET.value,
                ))
                nom_count += 1
                continue

            # phe_duyet_b2 / confirmed: Admin PheDuyet DONG_Y, nomination → PHE_DUYET_CUOI
            admin_time = sent_at + timedelta(days=random.randint(13, 20))
            admin_pd = PheDuyet(
                de_xuat_id=dx.id,
                phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value,
                ket_qua=KetQuaDuyet.DONG_Y.value,
                nguoi_duyet_id=admin.id if admin else None,
                ngay_duyet=admin_time,
            )
            db.session.add(admin_pd)
            dx.trang_thai = TrangThaiDeXuat.PHE_DUYET_CUOI.value
            db.session.flush()

            if stage == 'phe_duyet_b2':
                # Add partial Hội đồng votes (some roles voted, some not)
                voted_roles = random.sample(HOI_DONG_VAI_TRO, k=random.randint(2, 5))
                for ct in items:
                    for vai_tro in voted_roles:
                        voter = admin  # reuse admin user as proxy voter for seed
                        db.session.add(HoiDongBieuQuyet(
                            de_xuat_id=dx.id,
                            chi_tiet_id=ct.id,
                            nguoi_bieu_quyet_id=voter.id if voter else (admin.id if admin else 1),
                            vai_tro=vai_tro,
                            ket_qua=KET_QUA_DONG_Y,
                        ))
                nom_count += 1
                continue

            # confirmed: all 6 roles voted DONG_Y + KhenThuong created
            for ct in items:
                for vai_tro in HOI_DONG_VAI_TRO:
                    db.session.add(HoiDongBieuQuyet(
                        de_xuat_id=dx.id,
                        chi_tiet_id=ct.id,
                        nguoi_bieu_quyet_id=admin.id if admin else 1,
                        vai_tro=vai_tro,
                        ket_qua=KET_QUA_DONG_Y,
                    ))
                db.session.add(KhenThuong(
                    de_xuat_id=dx.id,
                    chi_tiet_id=ct.id,
                    quan_nhan_id=ct.quan_nhan_id,
                    don_vi_id=unit.id,
                    ho_ten=ct.quan_nhan.ho_ten if ct.quan_nhan else 'Don vi',
                    cap_bac=ct.quan_nhan.cap_bac if ct.quan_nhan else None,
                    chuc_vu=ct.quan_nhan.chuc_vu if ct.quan_nhan else None,
                    doi_tuong=ct.doi_tuong,
                    loai_danh_hieu=ct.loai_danh_hieu,
                    nam_hoc=nam_hoc,
                    nguoi_duyet_id=admin.id if admin else None,
                    ngay_duyet=admin_time + timedelta(days=3),
                ))
            nom_count += 1

    db.session.commit()
    print(f'  Created {nom_count} nominations.')


def validate():
    checks = {
        'quan_nhan':     QuanNhan.query.count(),
        'de_xuat':       DeXuat.query.count(),
        'nhap':          DeXuat.query.filter_by(trang_thai=TrangThaiDeXuat.NHAP.value).count(),
        'cho_duyet':     DeXuat.query.filter_by(trang_thai=TrangThaiDeXuat.CHO_DUYET.value).count(),
        'dang_duyet':    DeXuat.query.filter_by(trang_thai=TrangThaiDeXuat.DANG_DUYET.value).count(),
        'hoi_dong':      DeXuat.query.filter_by(trang_thai=TrangThaiDeXuat.HOI_DONG.value).count(),
        'phe_duyet_cuoi':DeXuat.query.filter_by(trang_thai=TrangThaiDeXuat.PHE_DUYET_CUOI.value).count(),
        'khen_thuong':   KhenThuong.query.count(),
        'hoi_dong_votes':HoiDongBieuQuyet.query.count(),
        'admin_approved_ct': DeXuatChiTiet.query.filter_by(admin_approved=True).count(),
    }
    return checks


def main():
    app = create_app()
    with app.app_context():
        units = DonVi.query.filter_by(is_active=True).order_by(DonVi.thu_tu).all()
        if not units:
            raise RuntimeError('No units found. Run init_db.py first.')

        print('1) Resetting data...')
        reset_data()

        # Re-query units after reset to get fresh session-bound objects
        units = DonVi.query.filter_by(is_active=True).order_by(DonVi.thu_tu).all()

        print('2) Creating accounts...')
        ensure_department_users()
        ensure_unit_users(units)
        ensure_award_titles()
        db.session.commit()

        print(f'3) Creating {TARGET_PERSONNEL} personnel...')
        all_qn = seed_personnel(units)

        print('4) Creating 20 nominations (full 3-table flow)...')
        seed_nominations(units, all_qn)

        checks = validate()
        print('--- RESULTS ---')
        for k, v in checks.items():
            print(f'  {k}: {v}')

        assert checks['de_xuat'] == TARGET_NOMINATIONS, \
            f'Wrong nomination count: {checks["de_xuat"]} != {TARGET_NOMINATIONS}'
        assert checks['hoi_dong'] == 6, f'Bang 1 should be 6, got {checks["hoi_dong"]}'
        assert checks['phe_duyet_cuoi'] == 8, \
            f'PHE_DUYET_CUOI should be 8 (4 b2 + 4 confirmed), got {checks["phe_duyet_cuoi"]}'
        assert checks['khen_thuong'] > 0, 'Must have KhenThuong records'
        print('Done.')


if __name__ == '__main__':
    main()
