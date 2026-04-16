"""
Reset data and generate sample dataset:
- 100 personnel
- 50 nominations
- 10 final approvals

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


TARGET_PERSONNEL = 100
TARGET_NOMINATIONS = 50
TARGET_FINAL_APPROVAL = 10


HO = [
    'Nguyen', 'Tran', 'Le', 'Pham', 'Hoang', 'Phan', 'Vo', 'Dang',
    'Bui', 'Do', 'Ngo', 'Duong', 'Ly', 'Trinh', 'Dinh',
]
DEM = [
    'Van', 'Thi', 'Duc', 'Minh', 'Quoc', 'Xuan', 'Ngoc', 'Thanh',
    'Huu', 'Duy', 'Anh', 'Tuan', 'Quang', 'Trong', 'Tien',
]
TEN = [
    'An', 'Binh', 'Cuong', 'Dung', 'Hung', 'Hieu', 'Khanh', 'Long',
    'Nam', 'Phong', 'Quan', 'Son', 'Tai', 'Tung', 'Vinh', 'Dat',
    'Hoa', 'Lam', 'Manh', 'Nghia', 'Phu', 'Thinh', 'Toan', 'Trung',
]

CAP_BAC_MAIN = ['Thieu uy', 'Trung uy', 'Thuong uy', 'Dai uy', 'Thieu ta', 'Trung ta']
CAP_BAC_STUDENT = ['Binh nhi', 'Binh nhat', 'Ha si', 'Trung si']

CHUC_VU_GV = ['Giang vien', 'Giang vien chinh', 'Pho Truong bo mon', 'Truong bo mon']
CHUC_VU_CB = ['Tro ly', 'Nhan vien', 'Pho truong ban', 'Truong ban', 'Pho truong phong']
CHUC_VU_TD = ['Tieu doi truong', 'Trung doi truong', 'Pho dai doi truong', 'Dai doi truong']


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
        'admin': (Role.ADMIN, 'Co quan Tuyen huan'),
        'phong_khoahoc': (Role.PHONG_KHOAHOC, 'Phong Khoa hoc'),
        'phong_daotao': (Role.PHONG_DAOTAO, 'Phong Dao tao'),
        'tt_phong_chinhtri': (Role.THU_TRUONG_PHONG_CHINHTRI, 'Thu truong Phong Chinh tri'),
        'tt_phong_tmhc': (Role.THU_TRUONG_PHONG_TMHC, 'Thu truong Phong TM-HC'),
        'ban_canbo': (Role.BAN_CANBO, 'Ban Can bo'),
        'ban_tochuc': (Role.BAN_TOCHUC, 'Ban To chuc'),
        'ban_tuyenhuan': (Role.BAN_TUYENHUAN, 'Ban Tuyen huan'),
        'ban_ctcq': (Role.BAN_CTCQ, 'Ban Cong tac quan chung'),
        'ban_cntt': (Role.BAN_CNTT, 'Ban Cong nghe thong tin'),
        'ban_tachuan': (Role.BAN_TAC_HUAN, 'Ban Tac huan'),
        'ban_khaothi': (Role.BAN_KHAOTHI, 'Ban Khao thi'),
        'uyban_kiemtra': (Role.UY_BAN_KIEMTRA, 'Uy ban Kiem tra'),
        'ban_quanluc': (Role.BAN_QUANLUC, 'Ban Quan luc'),
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
            if existing.role != Role.UNIT_USER:
                continue
            if existing.don_vi_id != unit.id:
                existing.don_vi_id = unit.id
            continue

        u = User(
            username=username,
            ho_ten=unit.ten_don_vi,
            role=Role.UNIT_USER,
            don_vi_id=unit.id,
        )
        u.set_password('123456')
        db.session.add(u)


def ensure_award_titles():
    defaults = [
        ('Chiến sĩ thi đua', 'CSTD', 'Cá nhân', 1),
        ('Chiến sĩ tiên tiến', 'CSTT', 'Cá nhân', 2),
        ('Đơn vị quyết thắng', 'DVQT', 'Đơn vị', 3),
        ('Đơn vị tiên tiến', 'DVTT', 'Đơn vị', 4),
    ]
    for ten, ma, pham_vi, thu_tu in defaults:
        if DanhHieu.query.filter_by(ten_danh_hieu=ten).first():
            continue
        dh = DanhHieu(ten_danh_hieu=ten, ma_danh_hieu=ma, pham_vi=pham_vi, thu_tu=thu_tu, is_active=True)
        dh.tieu_chi = ['muc_do_hoan_thanh', 'phieu_tin_nhiem', 'xep_loai_dang_vien', 'diem_nckh']
        db.session.add(dh)


def reset_data():
    ThongBao.query.delete()
    KhenThuong.query.delete()
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

    ensure_department_users()
    db.session.commit()


def seed_personnel(units):
    all_qn = []
    all_doi_tuong = [e.value for e in DoiTuong]
    khoa_units = [u for u in units if u.ma_don_vi.startswith('K')]
    td_units = [u for u in units if u.ma_don_vi.startswith('TD')]
    other_units = [u for u in units if u not in khoa_units + td_units]

    for idx in range(TARGET_PERSONNEL):
        if idx < 45 and khoa_units:
            unit = random.choice(khoa_units)
            doi_tuong = random.choice([DoiTuong.GV.value, DoiTuong.CB.value, DoiTuong.CCQP.value])
        elif idx < 80 and td_units:
            unit = random.choice(td_units)
            doi_tuong = random.choice([
                DoiTuong.SV_NAM1.value, DoiTuong.SV_NAM2.value, DoiTuong.SV_NAM3.value,
                DoiTuong.SV_NAM4.value, DoiTuong.QNCN.value, DoiTuong.CNV.value,
            ])
        else:
            unit = random.choice(other_units or units)
            doi_tuong = random.choice(all_doi_tuong)

        is_student = doi_tuong.startswith('Học viên')
        cap_bac = random.choice(CAP_BAC_STUDENT if is_student else CAP_BAC_MAIN)

        if doi_tuong == DoiTuong.GV.value:
            chuc_vu = random.choice(CHUC_VU_GV)
            hoc_vi = random.choice(['Thạc sĩ', 'Tiến sĩ'])
        elif is_student:
            chuc_vu = random.choice(['Hoc vien', 'Lop truong', 'Tieu doi truong'])
            hoc_vi = 'Không'
        else:
            chuc_vu = random.choice(CHUC_VU_CB + CHUC_VU_TD)
            hoc_vi = random.choice(['Không', 'Thạc sĩ'])

        qn = QuanNhan(
            don_vi_id=unit.id,
            ho_ten=random_name(),
            cap_bac=cap_bac,
            chuc_danh=doi_tuong,
            chuc_vu=chuc_vu,
            don_vi_truc_thuoc=f"{unit.ma_don_vi}-D{random.randint(1, 6)}",
            can_cuoc_cong_dan=random_cccd(idx + 1),
            ngay_sinh=random_date(1970, 2004),
            ngay_nhap_ngu=f"{random.randint(1, 12):02d}/{random.randint(1990, 2024)}",
            doi_tuong=doi_tuong,
            hoc_ham='Không',
            hoc_vi=hoc_vi,
            trinh_do_hoc_van='12/12',
            ngoai_ngu=random.choice(['Anh B1', 'Anh B2', 'Anh C1']),
            la_chi_huy=False,
            la_bi_thu=False,
            is_active=True,
        )
        db.session.add(qn)
        all_qn.append(qn)

        # Commit in batches to avoid long-lived transactions/connection drops
        if (idx + 1) % 20 == 0:
            db.session.commit()

    db.session.commit()

    # set one commander + one secretary per unit if possible
    by_unit = {}
    for qn in all_qn:
        by_unit.setdefault(qn.don_vi_id, []).append(qn)

    for _, rows in by_unit.items():
        if not rows:
            continue
        rows[0].la_chi_huy = True
        if len(rows) > 1:
            rows[1].la_bi_thu = True

    db.session.commit()
    return all_qn


def seed_certificates(personnel):
    cert_names = [
        'Giay khen hoan thanh tot nhiem vu',
        'Bang khen cap co so',
        'Thanh tich NCKH cap truong',
    ]
    sample = random.sample(personnel, min(45, len(personnel)))
    for qn in sample:
        for _ in range(random.randint(1, 2)):
            db.session.add(ChungChi(
                quan_nhan_id=qn.id,
                loai=random.choice(['Bang khen', 'Giay khen', 'Thanh tich khac']),
                ten_chung_chi=random.choice(cert_names),
                so_hieu=f"QD-{random.randint(100, 999)}/{random.randint(2022, 2026)}",
                ngay_cap=random_date(2022, 2026),
                co_quan_cap=random.choice(['Nha truong', 'Bo Quoc phong', 'Tong cuc Chinh tri']),
            ))
    db.session.commit()


def create_de_xuat_item(de_xuat_id, qn, nam_hoc, loai_danh_hieu):
    return DeXuatChiTiet(
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
        nckh_noi_dung='Tham gia de tai NCKH cap truong, nghiem thu kha tro len',
        thanh_tich_ca_nhan_khac='Co thanh tich tieu bieu trong phong trao don vi',
    )


def mark_department_review(pd, items, result, reviewer_id, decided_at):
    pd.ket_qua = result
    pd.nguoi_duyet_id = reviewer_id
    pd.ngay_duyet = decided_at
    if result == KetQuaDuyet.TU_CHOI.value:
        pd.ly_do = 'Chua dat tieu chuan theo quy dinh'

    for idx, ct in enumerate(items):
        if not in_scope(PhongDuyet(pd.phong_duyet), ct.doi_tuong):
            continue
        row_result = result
        row_reason = None
        if result == KetQuaDuyet.TU_CHOI.value and idx == 0:
            row_result = KetQuaDuyet.TU_CHOI.value
            row_reason = 'Chua dat tieu chuan theo doi tuong'
        elif result == KetQuaDuyet.TU_CHOI.value:
            row_result = KetQuaDuyet.DONG_Y.value

        db.session.add(KetQuaDuyetChiTiet(
            phe_duyet_id=pd.id,
            chi_tiet_id=ct.id,
            ket_qua=row_result,
            ly_do=row_reason,
        ))


def seed_nominations(units, personnel):
    unit_users = {
        u.don_vi_id: u for u in User.query.filter_by(role=Role.UNIT_USER).all() if u.don_vi_id
    }

    dept_reviewer = {
        PhongDuyet.PHONG_KHOAHOC.value: User.query.filter_by(role=Role.PHONG_KHOAHOC).first(),
        PhongDuyet.PHONG_DAOTAO.value: User.query.filter_by(role=Role.PHONG_DAOTAO).first(),
        PhongDuyet.BAN_CANBO.value: User.query.filter_by(role=Role.BAN_CANBO).first(),
        PhongDuyet.BAN_TOCHUC.value: User.query.filter_by(role=Role.BAN_TOCHUC).first(),
        PhongDuyet.BAN_TUYENHUAN.value: User.query.filter_by(role=Role.BAN_TUYENHUAN).first(),
        PhongDuyet.BAN_CTCQ.value: User.query.filter_by(role=Role.BAN_CTCQ).first(),
        PhongDuyet.BAN_CNTT.value: User.query.filter_by(role=Role.BAN_CNTT).first(),
        PhongDuyet.BAN_TAC_HUAN.value: User.query.filter_by(role=Role.BAN_TAC_HUAN).first(),
        PhongDuyet.BAN_KHAOTHI.value: User.query.filter_by(role=Role.BAN_KHAOTHI).first(),
        PhongDuyet.UY_BAN_KIEMTRA.value: User.query.filter_by(role=Role.UY_BAN_KIEMTRA).first(),
        PhongDuyet.BAN_QUANLUC.value: User.query.filter_by(role=Role.BAN_QUANLUC).first(),
        PhongDuyet.THU_TRUONG_PHONG_CHINHTRI.value: User.query.filter_by(role=Role.THU_TRUONG_PHONG_CHINHTRI).first(),
        PhongDuyet.THU_TRUONG_PHONG_TMHC.value: User.query.filter_by(role=Role.THU_TRUONG_PHONG_TMHC).first(),
    }
    admin = User.query.filter_by(role=Role.ADMIN).first()

    by_unit = {}
    for qn in personnel:
        by_unit.setdefault(qn.don_vi_id, []).append(qn)

    nom_plan = (
        [('Nháp', 8)] +
        [('Chờ duyệt', 12)] +
        [('Đang duyệt', 10)] +
        [('Đã duyệt', 10)] +
        [('Phê duyệt cuối', TARGET_FINAL_APPROVAL)]
    )

    status_list = []
    for name, c in nom_plan:
        status_list.extend([name] * c)
    random.shuffle(status_list)

    units_pool = [u for u in units if u.id in by_unit and u.id in unit_users and len(by_unit[u.id]) >= 2]
    if not units_pool:
        return

    awards_personal = ['Chiến sĩ thi đua', 'Chiến sĩ tiên tiến']

    created_nom = 0
    final_created = 0
    idx = 0
    max_attempts = TARGET_NOMINATIONS * 5

    while created_nom < TARGET_NOMINATIONS and idx < max_attempts:
        target_status = status_list[created_nom]
        unit = units_pool[idx % len(units_pool)]
        idx += 1
        creator = unit_users[unit.id]
        unit_personnel = by_unit[unit.id]

        nam_hoc = random.choice(['2024-2025', '2025-2026'])
        created_at = datetime.utcnow() - timedelta(days=random.randint(10, 90))
        sent_at = None if target_status == 'Nháp' else (created_at + timedelta(days=random.randint(1, 3)))

        dx = DeXuat(
            don_vi_id=unit.id,
            nam_hoc=nam_hoc,
            nguoi_tao_id=creator.id,
            trang_thai=target_status,
            ngay_gui=sent_at,
            ghi_chu='Du lieu seed tu dong',
        )
        db.session.add(dx)
        db.session.flush()

        chosen = random.sample(unit_personnel, k=min(len(unit_personnel), random.randint(2, 4)))
        items = []
        for qn in chosen:
            ct = create_de_xuat_item(dx.id, qn, nam_hoc, random.choice(awards_personal))
            db.session.add(ct)
            items.append(ct)

        db.session.flush()

        if target_status == 'Nháp':
            created_nom += 1
            continue

        pd_map = {}
        for dept in DEPT_FLOW:
            pd = PheDuyet(
                de_xuat_id=dx.id,
                phong_duyet=dept.value,
                ket_qua=KetQuaDuyet.CHO_DUYET.value,
            )
            db.session.add(pd)
            pd_map[dept.value] = pd
        db.session.flush()

        if target_status == 'Chờ duyệt':
            created_nom += 1
            continue

        if target_status == 'Đang duyệt':
            reviewed = random.sample(DEPT_FLOW, k=random.randint(3, 6))
            for dept in reviewed:
                reviewer = dept_reviewer.get(dept.value)
                mark_department_review(
                    pd_map[dept.value],
                    items,
                    KetQuaDuyet.DONG_Y.value,
                    reviewer.id if reviewer else None,
                    sent_at + timedelta(days=random.randint(1, 8)) if sent_at else datetime.utcnow(),
                )
            created_nom += 1
            continue

        # DA_DUYET + PHE_DUYET_CUOI
        for dept in DEPT_FLOW:
            reviewer = dept_reviewer.get(dept.value)
            mark_department_review(
                pd_map[dept.value],
                items,
                KetQuaDuyet.DONG_Y.value,
                reviewer.id if reviewer else None,
                sent_at + timedelta(days=random.randint(1, 12)) if sent_at else datetime.utcnow(),
            )

        if target_status == 'Đã duyệt':
            admin_pd = PheDuyet(
                de_xuat_id=dx.id,
                phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value,
                ket_qua=KetQuaDuyet.CHO_DUYET.value,
            )
            db.session.add(admin_pd)
            created_nom += 1
            continue

        # Phê duyệt cuối
        admin_time = (sent_at + timedelta(days=random.randint(13, 20))) if sent_at else datetime.utcnow()
        admin_pd = PheDuyet(
            de_xuat_id=dx.id,
            phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value,
            ket_qua=KetQuaDuyet.DONG_Y.value,
            nguoi_duyet_id=admin.id if admin else None,
            ngay_duyet=admin_time,
        )
        db.session.add(admin_pd)
        db.session.flush()

        for ct in items:
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
                ngay_duyet=admin_time,
            ))

        created_nom += 1
        final_created += 1

    if created_nom < TARGET_NOMINATIONS:
        raise RuntimeError(f'Khong the tao du {TARGET_NOMINATIONS} de xuat, chi tao duoc {created_nom}')
    if final_created < TARGET_FINAL_APPROVAL:
        raise RuntimeError(f'Khong the tao du {TARGET_FINAL_APPROVAL} de xuat phe duyet cuoi, chi tao duoc {final_created}')

    db.session.commit()


def validate_counts():
    checks = {
        'quan_nhan': QuanNhan.query.count(),
        'de_xuat': DeXuat.query.count(),
        'phe_duyet_cuoi': DeXuat.query.filter_by(trang_thai=TrangThaiDeXuat.PHE_DUYET_CUOI.value).count(),
        'khen_thuong': KhenThuong.query.count(),
        'unit_users': User.query.filter_by(role=Role.UNIT_USER).count(),
    }
    return checks


def main():
    app = create_app()
    with app.app_context():
        units = DonVi.query.filter_by(is_active=True).order_by(DonVi.thu_tu).all()
        if not units:
            raise RuntimeError('Khong tim thay don vi. Hay chay init_db.py truoc.')

        print('1) Reset du lieu...')
        reset_data()

        print('2) Tao tai khoan...')
        ensure_department_users()
        ensure_unit_users(units)
        ensure_award_titles()
        db.session.commit()

        print('3) Tao 100 quan nhan...')
        all_qn = seed_personnel(units)

        print('4) Tao chung chi mau...')
        seed_certificates(all_qn)

        print('5) Tao 50 de xuat + 10 phe duyet cuoi...')
        seed_nominations(units, all_qn)

        checks = validate_counts()
        print('--- KET QUA ---')
        for k, v in checks.items():
            print(f'{k}: {v}')

        if checks['quan_nhan'] != TARGET_PERSONNEL:
            raise RuntimeError(f'Sai so luong quan nhan: {checks["quan_nhan"]} != {TARGET_PERSONNEL}')
        if checks['de_xuat'] != TARGET_NOMINATIONS:
            raise RuntimeError(f'Sai so luong de xuat: {checks["de_xuat"]} != {TARGET_NOMINATIONS}')
        if checks['phe_duyet_cuoi'] != TARGET_FINAL_APPROVAL:
            raise RuntimeError(f'Sai so luong phe duyet cuoi: {checks["phe_duyet_cuoi"]} != {TARGET_FINAL_APPROVAL}')

        print('Done.')


if __name__ == '__main__':
    main()
