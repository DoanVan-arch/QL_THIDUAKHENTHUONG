"""
Xóa toàn bộ dữ liệu cũ (trừ đơn vị và tài khoản hệ thống),
tạo 200 quân nhân ngẫu nhiên cùng đề xuất, phê duyệt, khen thưởng.
Chạy: python seed_sample.py
"""
import sys
import io
import random
from datetime import date, datetime, timedelta

# Fix Windows console encoding
if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

from app import create_app
from app.extensions import db
from app.models.user import User, Role
from app.models.unit import DonVi
from app.models.personnel import QuanNhan
from app.models.certificate import ChungChi
from app.models.nomination import DeXuat, DeXuatChiTiet, TrangThaiDeXuat, LoaiDanhHieu
from app.models.approval import PheDuyet, PhongDuyet, KetQuaDuyet, KetQuaDuyetChiTiet
from app.models.reward import KhenThuong
from app.models.notification import ThongBao
from app.models.personnel import MucDoHoanThanh

app = create_app()

# ---------------------------------------------------------------------------
# Vietnamese name components (with diacritics)
# ---------------------------------------------------------------------------
HO = [
    'Nguyễn', 'Trần', 'Lê', 'Phạm', 'Hoàng', 'Huỳnh', 'Phan', 'Vũ',
    'Võ', 'Đặng', 'Bùi', 'Đỗ', 'Hồ', 'Ngô', 'Dương', 'Lý',
    'Trịnh', 'Đinh', 'Lưu', 'Tô',
]
DEM = [
    'Văn', 'Thị', 'Đức', 'Minh', 'Quốc', 'Đình', 'Xuân', 'Ngọc',
    'Thanh', 'Hữu', 'Duy', 'Anh', 'Hồng', 'Tuấn', 'Quang', 'Bá',
    'Công', 'Trọng', 'Tiến', 'Hải',
]
TEN = [
    'An', 'Bình', 'Cường', 'Dũng', 'Hùng', 'Hiếu', 'Khánh', 'Long',
    'Minh', 'Nam', 'Phong', 'Quân', 'Sơn', 'Tài', 'Tùng', 'Vinh',
    'Đạt', 'Hải', 'Hòa', 'Lâm', 'Mạnh', 'Nghĩa', 'Phú', 'Thịnh',
    'Toàn', 'Trung', 'Trường', 'Uy', 'Vũ', 'Huy', 'Kiên', 'Thắng',
    'Đông', 'Tú', 'Hoàng', 'Quý', 'Tuấn', 'Thành', 'Khải', 'Nhật',
]

CAP_BAC_LIST = [
    'Thiếu úy', 'Trung úy', 'Thượng úy', 'Đại úy',
    'Thiếu tá', 'Trung tá', 'Thượng tá',
]

CAP_BAC_SV = ['Binh nhì', 'Binh nhất', 'Hạ sĩ', 'Trung sĩ']

DOI_TUONG_LIST = [
    'Giảng viên', 'Cán bộ', 'Công chức quốc phòng',
    'Quân nhân chuyên nghiệp', 'Công nhân viên',
    'Học viên năm thứ I', 'Học viên năm thứ II',
    'Học viên năm thứ III', 'Học viên năm thứ IV',
    'Học viên sau đại học', 'Học viên VB2',
    'Học viên tiến sĩ', 'Học viên quốc tế',
]

DOI_TUONG_SV = [
    'Học viên năm thứ I', 'Học viên năm thứ II',
    'Học viên năm thứ III', 'Học viên năm thứ IV',
    'Học viên sau đại học', 'Học viên VB2',
    'Học viên tiến sĩ', 'Học viên quốc tế',
]

DOI_TUONG_CADRE = [
    'Giảng viên', 'Cán bộ', 'Công chức quốc phòng',
    'Quân nhân chuyên nghiệp', 'Công nhân viên',
]

# BAN_QUANLUC scope
BAN_QUANLUC_DOI_TUONG = {'Công nhân viên', 'Quân nhân chuyên nghiệp', 'Công chức quốc phòng'}

CHUC_VU_GV = [
    'Giảng viên', 'Giảng viên chính', 'Phó Trưởng bộ môn',
    'Trưởng bộ môn', 'Phó Trưởng khoa', 'Trưởng khoa',
]

CHUC_VU_CB = [
    'Trợ lý', 'Nhân viên', 'Phó trưởng ban', 'Trưởng ban',
    'Phó trưởng phòng', 'Trưởng phòng',
]

CHUC_VU_TD = [
    'Tiểu đội trưởng', 'Trung đội trưởng', 'Phó đại đội trưởng',
    'Đại đội trưởng', 'Phó tiểu đoàn trưởng', 'Tiểu đoàn trưởng',
]

HOC_VI = ['Không', 'Không', 'Không', 'Thạc sĩ', 'Thạc sĩ', 'Tiến sĩ']
HOC_HAM = ['Không', 'Không', 'Không', 'Không', 'Phó Giáo sư']
NGOAI_NGU = ['Anh B1', 'Anh B2', 'Anh C1', 'Trung B1', 'Nga B1', 'Pháp B1']

LOAI_CHUNG_CHI = ['Bằng khen', 'Giấy khen', 'Thành tích khác']

TEN_CHUNG_CHI = [
    'Chiến sĩ thi đua cấp cơ sở năm 2024',
    'Giấy khen hoàn thành xuất sắc nhiệm vụ',
    'Bằng khen Bộ Quốc phòng',
    'Giấy khen phong trào thi đua',
    'Chiến sĩ tiên tiến năm 2023',
    'Bằng khen Tổng cục Chính trị',
    'Giấy khen NCKH xuất sắc',
    'Bằng khen phong trào thể dục thể thao',
    'Giấy khen đạt thành tích trong huấn luyện',
    'Giấy khen công tác Đảng, công tác chính trị',
    'Chiến sĩ thi đua cấp cơ sở năm 2023',
    'Giấy khen hoàn thành tốt nhiệm vụ năm 2022',
    'Bằng khen Bộ trưởng Bộ Quốc phòng năm 2024',
]

KET_QUA = ['Giỏi', 'Khá', 'Xuất sắc', 'Đạt']

NCKH_NOI_DUNG = [
    'Tham gia nghiên cứu đề tài NCKH cấp Trường, nghiệm thu đạt Khá',
    'Bài đăng trên Website Nhà trường',
    'Sáng kiến ứng dụng hiệu quả trong giảng dạy',
    'Nghiên cứu chuyên đề khoa học đạt giải Nhì cấp hệ',
    'Chủ nhiệm đề tài NCKH cấp khoa, nghiệm thu đạt Giỏi',
    'Viết bài tạp chí Giáo dục Lý luận chính trị quân sự',
    'Tham gia hội thảo khoa học cấp Trường',
    'Đề tài NCKH sinh viên đạt giải Ba cấp Trường',
]


def random_name():
    return f'{random.choice(HO)} {random.choice(DEM)} {random.choice(TEN)}'


def random_date(start_year, end_year):
    start = date(start_year, 1, 1)
    end = date(end_year, 12, 31)
    delta = (end - start).days
    return start + timedelta(days=random.randint(0, delta))


def _is_in_dept_scope(role_value, doi_tuong):
    """Check if doi_tuong is in department's review scope."""
    if role_value == PhongDuyet.BAN_QUANLUC.value:
        return doi_tuong in BAN_QUANLUC_DOI_TUONG
    elif role_value == PhongDuyet.BAN_CANBO.value:
        return doi_tuong not in BAN_QUANLUC_DOI_TUONG
    return True  # All other depts review everything


with app.app_context():
    # ===================================================================
    # STEP 0: Delete all old data (preserve DonVi and system users)
    # ===================================================================
    print("=" * 60)
    print("BƯỚC 0: Xóa toàn bộ dữ liệu cũ...")

    # Delete in correct order (respect foreign keys)
    ThongBao.query.delete()
    KhenThuong.query.delete()
    KetQuaDuyetChiTiet.query.delete()
    PheDuyet.query.delete()
    # Delete MinhChung before DeXuatChiTiet
    from app.models.nomination import MinhChung
    MinhChung.query.delete()
    DeXuatChiTiet.query.delete()
    DeXuat.query.delete()
    ChungChi.query.delete()
    QuanNhan.query.delete()
    # Delete unit user accounts (keep department + admin)
    User.query.filter(User.role == Role.UNIT_USER).delete()
    db.session.commit()
    print("  Đã xóa: ThongBao, KhenThuong, KetQuaDuyetChiTiet, PheDuyet,")
    print("           MinhChung, DeXuatChiTiet, DeXuat, ChungChi, QuanNhan,")
    print("           và tất cả tài khoản đơn vị.")

    # ===================================================================
    # STEP 1: Recreate unit accounts
    # ===================================================================
    print("\nBƯỚC 1: Tạo tài khoản đơn vị...")
    units = DonVi.query.filter_by(is_active=True).order_by(DonVi.thu_tu).all()
    for unit in units:
        u = User(
            username=unit.ma_don_vi.lower(),
            ho_ten=unit.ten_don_vi,
            role=Role.UNIT_USER,
            don_vi_id=unit.id,
        )
        u.set_password('123456')
        db.session.add(u)
    db.session.commit()
    print(f"  Tạo {len(units)} tài khoản đơn vị (password: 123456).")

    # Categorize units
    khoa_units = [u for u in units if u.ma_don_vi.startswith('K')]
    td_units = [u for u in units if u.ma_don_vi.startswith('TD')]
    cq_units = [u for u in units if u.ma_don_vi.startswith('B')]
    he_units = [u for u in units if u.ma_don_vi.startswith('H')]

    # ===================================================================
    # STEP 2: Create 200 random personnel records
    # ===================================================================
    print("\nBƯỚC 2: Tạo 200 quân nhân ngẫu nhiên...")
    created_count = 0
    all_qn = []
    target_total = 200

    # Distribution plan:
    # Khoa (14 units): ~7-8 per unit = ~100-112 (70% GV, 30% CB)
    # Tiểu đoàn (13 units): ~5-6 per unit = ~65-78 (60% SV, 40% CB)
    # Cơ quan (11 units): ~2-3 per unit = ~22-33 (cadres only)
    # Hệ (3 units): ~2-3 per unit = ~6-9

    # Khoa units - lecturers + cadres
    for unit in khoa_units:
        n = random.randint(6, 9)
        for i in range(n):
            is_gv = random.random() < 0.7
            doi_tuong = 'Giảng viên' if is_gv else random.choice(DOI_TUONG_CADRE)
            hoc_vi = random.choice(HOC_VI) if is_gv else random.choice(['Không', 'Thạc sĩ'])
            hoc_ham = random.choice(HOC_HAM) if hoc_vi == 'Tiến sĩ' else 'Không'
            chuc_vu = random.choice(CHUC_VU_GV) if is_gv else random.choice(CHUC_VU_CB)

            qn = QuanNhan(
                don_vi_id=unit.id,
                ho_ten=random_name(),
                cap_bac=random.choice(CAP_BAC_LIST),
                chuc_danh=doi_tuong,
                chuc_vu=chuc_vu,
                ngay_sinh=random_date(1970, 1998),
                ngay_nhap_ngu=f'{random.randint(1, 12):02d}/{random.randint(1990, 2015)}',
                doi_tuong=doi_tuong,
                hoc_ham=hoc_ham,
                hoc_vi=hoc_vi,
                trinh_do_hoc_van='12/12',
                ngoai_ngu=random.choice(NGOAI_NGU),
                la_chi_huy=(i == 0),
                la_bi_thu=(i == 1),
            )
            db.session.add(qn)
            all_qn.append(qn)
            created_count += 1

    # Tiểu đoàn units - students + cadres
    for unit in td_units:
        n = random.randint(4, 6)
        for i in range(n):
            is_sv = random.random() < 0.6
            if is_sv:
                doi_tuong = random.choice(DOI_TUONG_SV)
                cap_bac = random.choice(CAP_BAC_SV)
                chuc_vu = random.choice(['Học viên', 'Tiểu đội trưởng', 'Lớp trưởng'])
                hoc_vi = 'Không'
            else:
                doi_tuong = random.choice(DOI_TUONG_CADRE)
                cap_bac = random.choice(CAP_BAC_LIST)
                chuc_vu = random.choice(CHUC_VU_TD)
                hoc_vi = random.choice(['Không', 'Thạc sĩ'])

            qn = QuanNhan(
                don_vi_id=unit.id,
                ho_ten=random_name(),
                cap_bac=cap_bac,
                chuc_danh=doi_tuong,
                chuc_vu=chuc_vu,
                ngay_sinh=random_date(1975, 2004),
                ngay_nhap_ngu=f'{random.randint(1, 12):02d}/{random.randint(2000, 2023)}',
                doi_tuong=doi_tuong,
                hoc_ham='Không',
                hoc_vi=hoc_vi,
                trinh_do_hoc_van='12/12',
                ngoai_ngu=random.choice(NGOAI_NGU),
                la_chi_huy=(i == 0 and not is_sv),
                la_bi_thu=(i == 1 and not is_sv),
            )
            db.session.add(qn)
            all_qn.append(qn)
            created_count += 1

    # Cơ quan units - cadres only
    for unit in cq_units:
        n = random.randint(2, 3)
        for i in range(n):
            doi_tuong = random.choice(DOI_TUONG_CADRE)
            qn = QuanNhan(
                don_vi_id=unit.id,
                ho_ten=random_name(),
                cap_bac=random.choice(CAP_BAC_LIST),
                chuc_danh=doi_tuong,
                chuc_vu=random.choice(CHUC_VU_CB),
                ngay_sinh=random_date(1970, 1995),
                ngay_nhap_ngu=f'{random.randint(1, 12):02d}/{random.randint(1990, 2010)}',
                doi_tuong=doi_tuong,
                hoc_ham='Không',
                hoc_vi=random.choice(['Không', 'Thạc sĩ']),
                trinh_do_hoc_van='12/12',
                ngoai_ngu=random.choice(NGOAI_NGU),
                la_chi_huy=(i == 0),
            )
            db.session.add(qn)
            all_qn.append(qn)
            created_count += 1

    # Hệ units - cadres + students
    for unit in he_units:
        n = random.randint(2, 3)
        for i in range(n):
            doi_tuong = random.choice(DOI_TUONG_CADRE)
            qn = QuanNhan(
                don_vi_id=unit.id,
                ho_ten=random_name(),
                cap_bac=random.choice(CAP_BAC_LIST),
                chuc_danh=doi_tuong,
                chuc_vu=random.choice(CHUC_VU_CB),
                ngay_sinh=random_date(1975, 1998),
                ngay_nhap_ngu=f'{random.randint(1, 12):02d}/{random.randint(1995, 2015)}',
                doi_tuong=doi_tuong,
                hoc_ham='Không',
                hoc_vi=random.choice(['Không', 'Thạc sĩ']),
                trinh_do_hoc_van='12/12',
                ngoai_ngu=random.choice(NGOAI_NGU),
                la_chi_huy=(i == 0),
            )
            db.session.add(qn)
            all_qn.append(qn)
            created_count += 1

    db.session.commit()
    print(f"  Tạo {created_count} quân nhân.")

    # If we haven't reached 200, add more to random khoa units
    while created_count < target_total:
        unit = random.choice(khoa_units + td_units)
        is_khoa = unit in khoa_units
        if is_khoa:
            is_gv = random.random() < 0.7
            doi_tuong = 'Giảng viên' if is_gv else random.choice(DOI_TUONG_CADRE)
            hoc_vi = random.choice(HOC_VI) if is_gv else 'Không'
            hoc_ham = random.choice(HOC_HAM) if hoc_vi == 'Tiến sĩ' else 'Không'
            chuc_vu = random.choice(CHUC_VU_GV) if is_gv else random.choice(CHUC_VU_CB)
            cap_bac = random.choice(CAP_BAC_LIST)
        else:
            is_sv = random.random() < 0.6
            if is_sv:
                doi_tuong = random.choice(DOI_TUONG_SV)
                cap_bac = random.choice(CAP_BAC_SV)
                chuc_vu = random.choice(['Học viên', 'Tiểu đội trưởng', 'Lớp trưởng'])
                hoc_vi = 'Không'
            else:
                doi_tuong = random.choice(DOI_TUONG_CADRE)
                cap_bac = random.choice(CAP_BAC_LIST)
                chuc_vu = random.choice(CHUC_VU_TD)
                hoc_vi = random.choice(['Không', 'Thạc sĩ'])
            hoc_ham = 'Không'

        qn = QuanNhan(
            don_vi_id=unit.id,
            ho_ten=random_name(),
            cap_bac=cap_bac,
            chuc_danh=doi_tuong,
            chuc_vu=chuc_vu,
            ngay_sinh=random_date(1975, 2004),
            ngay_nhap_ngu=f'{random.randint(1, 12):02d}/{random.randint(1995, 2020)}',
            doi_tuong=doi_tuong,
            hoc_ham=hoc_ham,
            hoc_vi=hoc_vi,
            trinh_do_hoc_van='12/12',
            ngoai_ngu=random.choice(NGOAI_NGU),
        )
        db.session.add(qn)
        all_qn.append(qn)
        created_count += 1

    db.session.commit()
    print(f"  Tổng cộng: {created_count} quân nhân (mục tiêu: {target_total}).")

    # ===================================================================
    # STEP 3: Add certificates to ~80 random personnel
    # ===================================================================
    print("\nBƯỚC 3: Tạo chứng chỉ, bằng khen...")
    cert_count = 0
    for qn in random.sample(all_qn, min(80, len(all_qn))):
        for _ in range(random.randint(1, 3)):
            cc = ChungChi(
                quan_nhan_id=qn.id,
                loai=random.choice(LOAI_CHUNG_CHI),
                ten_chung_chi=random.choice(TEN_CHUNG_CHI),
                so_hieu=f'QĐ-{random.randint(100, 999)}/{random.randint(2020, 2025)}',
                ngay_cap=random_date(2020, 2025),
                co_quan_cap=random.choice([
                    'Nhà trường', 'Bộ Quốc phòng', 'Tổng cục Chính trị',
                    'Quân ủy Trung ương', 'Đảng ủy Nhà trường',
                ]),
            )
            db.session.add(cc)
            cert_count += 1
    db.session.commit()
    print(f"  Tạo {cert_count} chứng chỉ/bằng khen.")

    # ===================================================================
    # STEP 4: Create nominations with various statuses
    # ===================================================================
    print("\nBƯỚC 4: Tạo đề xuất khen thưởng...")

    # Group personnel by don_vi_id
    qn_by_unit = {}
    for qn in all_qn:
        qn_by_unit.setdefault(qn.don_vi_id, []).append(qn)

    units_with_personnel = [u for u in units if u.id in qn_by_unit]

    # We'll create nominations with different statuses for realism:
    # - 3 NHAP (draft)
    # - 5 CHO_DUYET (submitted, waiting)
    # - 4 DANG_DUYET (some depts reviewed)
    # - 4 DA_DUYET (all 6 depts approved)
    # - 3 PHE_DUYET_CUOI (final approved -> KhenThuong records)
    # - 2 TU_CHOI (rejected by some dept)
    # Total: ~21 nominations

    nom_configs = (
        [('NHAP', 3)] +
        [('CHO_DUYET', 5)] +
        [('DANG_DUYET', 4)] +
        [('DA_DUYET', 4)] +
        [('PHE_DUYET_CUOI', 3)] +
        [('TU_CHOI', 2)]
    )

    random.shuffle(units_with_personnel)
    nom_count = 0
    nom_idx = 0  # index into units_with_personnel

    DEPT_PHONGS = [
        PhongDuyet.PHONG_CHINHTRI, PhongDuyet.PHONG_THAMMUU,
        PhongDuyet.PHONG_KHOAHOC, PhongDuyet.PHONG_DAOTAO,
        PhongDuyet.BAN_CANBO, PhongDuyet.BAN_TOCHUC,
        PhongDuyet.BAN_TUYENHUAN, PhongDuyet.BAN_CTCQ,
        PhongDuyet.BAN_CNTT, PhongDuyet.BAN_TAC_HUAN,
        PhongDuyet.BAN_QUANLUC,
    ]

    # Map PhongDuyet -> Role for finding department user accounts
    PHONG_TO_ROLE = {
        PhongDuyet.PHONG_CHINHTRI: Role.PHONG_CHINHTRI,
        PhongDuyet.PHONG_THAMMUU: Role.PHONG_THAMMUU,
        PhongDuyet.PHONG_KHOAHOC: Role.PHONG_KHOAHOC,
        PhongDuyet.PHONG_DAOTAO: Role.PHONG_DAOTAO,
        PhongDuyet.BAN_CANBO: Role.BAN_CANBO,
        PhongDuyet.BAN_TOCHUC: Role.BAN_TOCHUC,
        PhongDuyet.BAN_TUYENHUAN: Role.BAN_TUYENHUAN,
        PhongDuyet.BAN_CTCQ: Role.BAN_CTCQ,
        PhongDuyet.BAN_CNTT: Role.BAN_CNTT,
        PhongDuyet.BAN_TAC_HUAN: Role.BAN_TAC_HUAN,
        PhongDuyet.BAN_QUANLUC: Role.BAN_QUANLUC,
    }

    # Pre-load department users
    dept_users = {}
    for phong, role in PHONG_TO_ROLE.items():
        dept_users[phong] = User.query.filter_by(role=role).first()

    admin_user = User.query.filter_by(role=Role.ADMIN).first()

    for target_status, count in nom_configs:
        for _ in range(count):
            if nom_idx >= len(units_with_personnel):
                nom_idx = 0
            unit = units_with_personnel[nom_idx]
            nom_idx += 1

            unit_user = User.query.filter_by(don_vi_id=unit.id, role=Role.UNIT_USER).first()
            if not unit_user:
                continue

            unit_personnel = qn_by_unit.get(unit.id, [])
            if not unit_personnel:
                continue

            nam_hoc = random.choice(['2024-2025', '2025-2026'])
            ngay_gui = None if target_status == 'NHAP' else (
                datetime.utcnow() - timedelta(days=random.randint(1, 60))
            )

            de_xuat = DeXuat(
                don_vi_id=unit.id,
                nam_hoc=nam_hoc,
                nguoi_tao_id=unit_user.id,
                trang_thai=TrangThaiDeXuat.NHAP.value if target_status == 'NHAP' else TrangThaiDeXuat.CHO_DUYET.value,
                ngay_gui=ngay_gui,
            )
            db.session.add(de_xuat)
            db.session.flush()

            # Add 2-5 individuals
            n_items = random.randint(2, min(5, len(unit_personnel)))
            selected_qn = random.sample(unit_personnel, n_items)
            chi_tiets = []

            for qn in selected_qn:
                is_sv = qn.doi_tuong and 'Học viên' in qn.doi_tuong
                danh_hieu = random.choice([
                    LoaiDanhHieu.CHIEN_SI_THI_DUA.value,
                    LoaiDanhHieu.CHIEN_SI_TIEN_TIEN.value,
                ])

                ct = DeXuatChiTiet(
                    de_xuat_id=de_xuat.id,
                    quan_nhan_id=qn.id,
                    loai_danh_hieu=danh_hieu,
                    doi_tuong=qn.doi_tuong,
                    nam_hoc=nam_hoc,
                    muc_do_hoan_thanh=random.choice([
                        MucDoHoanThanh.HTXSNV.value, MucDoHoanThanh.HTTNV.value,
                        MucDoHoanThanh.HTNV.value,
                    ]),
                    kiem_tra_tin_hoc=random.choice(KET_QUA),
                    kiem_tra_dieu_lenh=random.choice(KET_QUA),
                    dia_ly_quan_su=random.choice(KET_QUA),
                    ban_sung=random.choice(KET_QUA),
                    the_luc=random.choice(KET_QUA),
                    kiem_tra_chinh_tri=random.choice(KET_QUA),
                    phieu_tin_nhiem=f'{random.randint(70, 100)}%',
                    danh_hieu_gv_gioi='Giảng viên giỏi cấp Trường' if qn.doi_tuong == 'Giảng viên' else None,
                    dinh_muc_giang_day=f'{random.randint(200, 400)} giờ' if qn.doi_tuong == 'Giảng viên' else None,
                    ket_qua_kiem_tra_giang=random.choice(KET_QUA) if qn.doi_tuong == 'Giảng viên' else None,
                    thoi_gian_lao_dong_kh=f'{random.randint(50, 200)} giờ' if qn.doi_tuong == 'Giảng viên' else None,
                    danh_hieu_hv_gioi='Học viên giỏi' if is_sv else None,
                    diem_tong_ket=f'{random.uniform(7.0, 9.5):.1f}' if is_sv else None,
                    ket_qua_thuc_hanh=random.choice(KET_QUA) if is_sv else None,
                    diem_nckh=round(random.uniform(5.0, 10.0), 1) if random.random() < 0.4 else None,
                    nckh_noi_dung=random.choice(NCKH_NOI_DUNG) if random.random() < 0.5 else None,
                )
                db.session.add(ct)
                chi_tiets.append(ct)

            db.session.flush()

            # For NHAP status: no approval records needed
            if target_status == 'NHAP':
                nom_count += 1
                continue

            # Create 6 PheDuyet records
            phe_duyets = {}
            for phong in DEPT_PHONGS:
                pd = PheDuyet(
                    de_xuat_id=de_xuat.id,
                    phong_duyet=phong.value,
                    ket_qua=KetQuaDuyet.CHO_DUYET.value,
                )
                db.session.add(pd)
                phe_duyets[phong.value] = pd

            db.session.flush()

            # For CHO_DUYET: leave all as pending
            if target_status == 'CHO_DUYET':
                nom_count += 1
                continue

            # For DANG_DUYET: approve 2-4 departments
            if target_status == 'DANG_DUYET':
                de_xuat.trang_thai = TrangThaiDeXuat.DANG_DUYET.value
                n_depts = random.randint(2, 4)
                approved_depts = random.sample(DEPT_PHONGS, n_depts)

                for phong in approved_depts:
                    pd = phe_duyets[phong.value]
                    dept_user = dept_users.get(phong)
                    pd.ket_qua = KetQuaDuyet.DONG_Y.value
                    pd.nguoi_duyet_id = dept_user.id if dept_user else None
                    pd.ngay_duyet = ngay_gui + timedelta(days=random.randint(1, 10)) if ngay_gui else datetime.utcnow()

                    # Create per-item approvals
                    for ct in chi_tiets:
                        in_scope = _is_in_dept_scope(phong.value, ct.doi_tuong)
                        kq_ct = KetQuaDuyetChiTiet(
                            phe_duyet_id=pd.id,
                            chi_tiet_id=ct.id,
                            ket_qua=KetQuaDuyet.DONG_Y.value,
                        )
                        db.session.add(kq_ct)

                nom_count += 1
                continue

            # For DA_DUYET, PHE_DUYET_CUOI, TU_CHOI: all 6 depts must review
            if target_status in ('DA_DUYET', 'PHE_DUYET_CUOI'):
                # All 6 depts approve all individuals
                for phong in DEPT_PHONGS:
                    pd = phe_duyets[phong.value]
                    dept_user = dept_users.get(phong)

                    pd.ket_qua = KetQuaDuyet.DONG_Y.value
                    pd.nguoi_duyet_id = dept_user.id if dept_user else None
                    pd.ngay_duyet = ngay_gui + timedelta(days=random.randint(1, 14)) if ngay_gui else datetime.utcnow()

                    for ct in chi_tiets:
                        kq_ct = KetQuaDuyetChiTiet(
                            phe_duyet_id=pd.id,
                            chi_tiet_id=ct.id,
                            ket_qua=KetQuaDuyet.DONG_Y.value,
                        )
                        db.session.add(kq_ct)

                de_xuat.trang_thai = TrangThaiDeXuat.DA_DUYET.value

                # Add admin PheDuyet for DA_DUYET
                admin_pd = PheDuyet(
                    de_xuat_id=de_xuat.id,
                    phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value,
                    ket_qua=KetQuaDuyet.CHO_DUYET.value,
                )
                db.session.add(admin_pd)

                if target_status == 'PHE_DUYET_CUOI':
                    de_xuat.trang_thai = TrangThaiDeXuat.PHE_DUYET_CUOI.value
                    admin_pd.ket_qua = KetQuaDuyet.DONG_Y.value
                    admin_pd.nguoi_duyet_id = admin_user.id if admin_user else None
                    admin_pd.ngay_duyet = (ngay_gui + timedelta(days=random.randint(15, 25))
                                           if ngay_gui else datetime.utcnow())

                    db.session.flush()

                    # Create KhenThuong records
                    now = admin_pd.ngay_duyet or datetime.utcnow()
                    for ct in chi_tiets:
                        kt = KhenThuong(
                            de_xuat_id=de_xuat.id,
                            chi_tiet_id=ct.id,
                            quan_nhan_id=ct.quan_nhan_id,
                            don_vi_id=unit.id,
                            ho_ten=ct.quan_nhan.ho_ten if ct.quan_nhan else '',
                            cap_bac=ct.quan_nhan.cap_bac if ct.quan_nhan else None,
                            chuc_vu=ct.quan_nhan.chuc_vu if ct.quan_nhan else None,
                            doi_tuong=ct.doi_tuong,
                            loai_danh_hieu=ct.loai_danh_hieu,
                            nam_hoc=nam_hoc,
                            nguoi_duyet_id=admin_user.id if admin_user else None,
                            ngay_duyet=now,
                        )
                        db.session.add(kt)

                nom_count += 1
                continue

            # TU_CHOI: some depts approved, one rejects some individuals
            if target_status == 'TU_CHOI':
                de_xuat.trang_thai = TrangThaiDeXuat.TU_CHOI.value

                # First 4-5 depts approve
                approving_depts = random.sample(DEPT_PHONGS, random.randint(4, 5))
                rejecting_dept = [p for p in DEPT_PHONGS if p not in approving_depts][0] if len(approving_depts) < 6 else random.choice(DEPT_PHONGS)

                for phong in DEPT_PHONGS:
                    pd = phe_duyets[phong.value]
                    dept_user = dept_users.get(phong)

                    if phong == rejecting_dept:
                        # This dept rejects at least one individual
                        pd.ket_qua = KetQuaDuyet.TU_CHOI.value
                        pd.nguoi_duyet_id = dept_user.id if dept_user else None
                        pd.ngay_duyet = ngay_gui + timedelta(days=random.randint(3, 10)) if ngay_gui else datetime.utcnow()
                        pd.ly_do = 'Không đủ điều kiện theo quy định'

                        for idx, ct in enumerate(chi_tiets):
                            in_scope = _is_in_dept_scope(phong.value, ct.doi_tuong)
                            # Reject at least 1 in-scope individual
                            if in_scope and idx == 0:
                                kq_ct = KetQuaDuyetChiTiet(
                                    phe_duyet_id=pd.id,
                                    chi_tiet_id=ct.id,
                                    ket_qua=KetQuaDuyet.TU_CHOI.value,
                                    ly_do='Không đạt tiêu chuẩn xét thi đua',
                                )
                            else:
                                kq_ct = KetQuaDuyetChiTiet(
                                    phe_duyet_id=pd.id,
                                    chi_tiet_id=ct.id,
                                    ket_qua=KetQuaDuyet.DONG_Y.value if in_scope else KetQuaDuyet.DONG_Y.value,
                                )
                            db.session.add(kq_ct)
                    elif phong in approving_depts:
                        pd.ket_qua = KetQuaDuyet.DONG_Y.value
                        pd.nguoi_duyet_id = dept_user.id if dept_user else None
                        pd.ngay_duyet = ngay_gui + timedelta(days=random.randint(1, 8)) if ngay_gui else datetime.utcnow()

                        for ct in chi_tiets:
                            kq_ct = KetQuaDuyetChiTiet(
                                phe_duyet_id=pd.id,
                                chi_tiet_id=ct.id,
                                ket_qua=KetQuaDuyet.DONG_Y.value,
                            )
                            db.session.add(kq_ct)
                    # Else: leave as CHO_DUYET (no review yet)

                # Create notifications for rejections
                rejected_unit_user = unit_user
                for ct in chi_tiets:
                    pd_rej = phe_duyets[rejecting_dept.value]
                    kq = KetQuaDuyetChiTiet.query.filter_by(
                        phe_duyet_id=pd_rej.id, chi_tiet_id=ct.id
                    ).first()
                    if kq and kq.ket_qua == KetQuaDuyet.TU_CHOI.value:
                        tb = ThongBao(
                            user_id=rejected_unit_user.id,
                            de_xuat_id=de_xuat.id,
                            chi_tiet_id=ct.id,
                            loai='tu_choi',
                            tieu_de=f'{rejecting_dept.value} từ chối: {ct.quan_nhan.ho_ten}',
                            noi_dung=f'Lý do: Không đạt tiêu chuẩn xét thi đua',
                        )
                        db.session.add(tb)

                nom_count += 1
                continue

    db.session.commit()
    print(f"  Tạo {nom_count} đề xuất khen thưởng.")

    # ===================================================================
    # Summary
    # ===================================================================
    total_qn = QuanNhan.query.count()
    total_cc = ChungChi.query.count()
    total_dx = DeXuat.query.count()
    total_kt = KhenThuong.query.count()
    total_tb = ThongBao.query.count()

    print("\n" + "=" * 60)
    print("KẾT QUẢ SEED DỮ LIỆU:")
    print(f"  Quân nhân:         {total_qn}")
    print(f"  Chứng chỉ:         {total_cc}")
    print(f"  Đề xuất:           {total_dx}")
    print(f"  Khen thưởng:       {total_kt}")
    print(f"  Thông báo:         {total_tb}")
    print("=" * 60)

    # Status breakdown
    for status in TrangThaiDeXuat:
        cnt = DeXuat.query.filter_by(trang_thai=status.value).count()
        if cnt > 0:
            print(f"  Đề xuất [{status.value}]: {cnt}")

    print("\nDone! Chạy 'python run.py' để khởi động server.")
