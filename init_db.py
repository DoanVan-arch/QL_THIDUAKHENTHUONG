"""
Khởi tạo database và seed dữ liệu ban đầu.
Chạy: python init_db.py
"""
import pymysql

# Tạo database nếu chưa có
conn = pymysql.connect(host='localhost', user='root', password='1111')
cursor = conn.cursor()
cursor.execute("CREATE DATABASE IF NOT EXISTS quanly_thidua_khenthuong CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci")
conn.close()
print("Database created/verified.")

from app import create_app
from app.extensions import db
from app.models.user import User, Role
from app.models.unit import DonVi, LoaiDonVi

app = create_app()

with app.app_context():
    db.create_all()
    print("Tables created.")

    # Seed units if empty
    if DonVi.query.count() == 0:
        units_data = [
            # Khối cơ quan
            ('B1', 'Phòng Tham mưu - Hành chính', LoaiDonVi.KHOI_COQUAN, 1),
            ('B2', 'Phòng Đào tạo', LoaiDonVi.KHOI_COQUAN, 2),
            ('B3', 'Phòng Chính trị', LoaiDonVi.KHOI_COQUAN, 3),
            ('B4', 'Phòng Khoa học quân sự', LoaiDonVi.KHOI_COQUAN, 4),
            ('B5', 'Phòng Hậu cần - Kỹ thuật', LoaiDonVi.KHOI_COQUAN, 5),
            ('B6', 'Ban Tài Chính', LoaiDonVi.KHOI_COQUAN, 6),
            ('B7', 'Ban Khảo thí và đảm bảo chất lượng giáo dục đào tạo', LoaiDonVi.KHOI_COQUAN, 7),
            ('B8', 'Ban Sau Đại học', LoaiDonVi.KHOI_COQUAN, 8),
            ('B9', 'Ban Thông tin khoa học quân sự', LoaiDonVi.KHOI_COQUAN, 9),
            ('B10', 'Ban Dự án', LoaiDonVi.KHOI_COQUAN, 10),
            ('B11', 'Ủy Ban kiểm tra Đảng', LoaiDonVi.KHOI_COQUAN, 11),
            # Các khoa
            ('K1', 'Khoa Triết học Mác - Lênin', LoaiDonVi.KHOA, 12),
            ('K2', 'Khoa Lịch sử Đảng Cộng sản Việt Nam', LoaiDonVi.KHOA, 13),
            ('K3', 'Khoa Công tác Đảng, Công tác Chính trị', LoaiDonVi.KHOA, 14),
            ('K4', 'Khoa Chiến thuật', LoaiDonVi.KHOA, 15),
            ('K5', 'Khoa Văn hóa - Ngoại ngữ', LoaiDonVi.KHOA, 16),
            ('K6', 'Khoa Kinh tế chính trị Mác - Lênin', LoaiDonVi.KHOA, 17),
            ('K7', 'Khoa Chủ nghĩa xã hội khoa học', LoaiDonVi.KHOA, 18),
            ('K8', 'Khoa Tâm lý học quân sự', LoaiDonVi.KHOA, 19),
            ('K9', 'Khoa Bắn súng', LoaiDonVi.KHOA, 20),
            ('K10', 'Khoa Quân sự chung', LoaiDonVi.KHOA, 21),
            ('K11', 'Khoa Giáo dục thể chất', LoaiDonVi.KHOA, 22),
            ('K12', 'Khoa Sư phạm quân sự', LoaiDonVi.KHOA, 23),
            ('K13', 'Khoa Tư tưởng Hồ Chí Minh', LoaiDonVi.KHOA, 24),
            ('K14', 'Khoa Nhà nước & Pháp luật', LoaiDonVi.KHOA, 25),
            # Đơn vị trực thuộc - Hệ
            ('H1', 'Hệ 1 (Hệ Chuyển loại cán bộ chính trị, hoàn thiện đại học)', LoaiDonVi.HE, 26),
            ('H2', 'Hệ 2 (Hệ sau đại học)', LoaiDonVi.HE, 27),
            ('H3', 'Hệ 3 (Hệ quốc tế)', LoaiDonVi.HE, 28),
            # Tiểu đoàn
            ('TD1', 'Tiểu đoàn 1', LoaiDonVi.TIEU_DOAN, 29),
            ('TD2', 'Tiểu đoàn 2', LoaiDonVi.TIEU_DOAN, 30),
            ('TD3', 'Tiểu đoàn 3', LoaiDonVi.TIEU_DOAN, 31),
            ('TD4', 'Tiểu đoàn 4', LoaiDonVi.TIEU_DOAN, 32),
            ('TD5', 'Tiểu đoàn 5', LoaiDonVi.TIEU_DOAN, 33),
            ('TD6', 'Tiểu đoàn 6', LoaiDonVi.TIEU_DOAN, 34),
            ('TD7', 'Tiểu đoàn 7', LoaiDonVi.TIEU_DOAN, 35),
            ('TD8', 'Tiểu đoàn 8', LoaiDonVi.TIEU_DOAN, 36),
            ('TD9', 'Tiểu đoàn 9', LoaiDonVi.TIEU_DOAN, 37),
            ('TD10', 'Tiểu đoàn 10', LoaiDonVi.TIEU_DOAN, 38),
            ('TD11', 'Tiểu đoàn 11', LoaiDonVi.TIEU_DOAN, 39),
            ('TD12', 'Tiểu đoàn 12', LoaiDonVi.TIEU_DOAN, 40),
            ('TDPV', 'Tiểu đoàn Phục vụ huấn luyện dã ngoại', LoaiDonVi.TIEU_DOAN, 41),
        ]

        for ma, ten, loai, tt in units_data:
            dv = DonVi(ma_don_vi=ma, ten_don_vi=ten, loai_don_vi=loai, thu_tu=tt)
            db.session.add(dv)

        db.session.commit()
        print(f"Seeded {len(units_data)} units.")

    # Seed admin & department users if empty
    if User.query.count() == 0:
        users_data = [
            ('admin', 'Ban thư ký Hội đồng thi đua khen thưởng', Role.ADMIN, None),
            ('phong_khoahoc', 'Phòng Khoa học quân sự', Role.PHONG_KHOAHOC, None),
            ('phong_daotao', 'Phòng Đào tạo', Role.PHONG_DAOTAO, None),
            ('tt_phong_chinhtri', 'Thủ trưởng Phòng Chính trị', Role.THU_TRUONG_PHONG_CHINHTRI, None),
            ('tt_phong_tmhc', 'Thủ trưởng Phòng TM-HC', Role.THU_TRUONG_PHONG_TMHC, None),
            ('ban_canbo', 'Ban Cán bộ', Role.BAN_CANBO, None),
            ('ban_tochuc', 'Ban Tổ chức', Role.BAN_TOCHUC, None),
            ('ban_tuyenhuan', 'Ban Tuyên huấn', Role.BAN_TUYENHUAN, None),
            ('ban_ctcq', 'Ban Công tác quần chúng', Role.BAN_CTCQ, None),
            ('ban_cntt', 'Ban Công nghệ thông tin', Role.BAN_CNTT, None),
            ('ban_tachuan', 'Ban Tác huấn', Role.BAN_TAC_HUAN, None),
            ('ban_baove_anninh', 'Ban Bảo vệ an ninh', Role.BAN_BAOVE_ANNINH, None),
            ('ban_khth', 'Ban Kế hoạch tổng hợp', Role.BAN_KEHOACH_TONGHOP, None),
            ('uyban_kiemtra', 'Ủy ban Kiểm tra', Role.UY_BAN_KIEMTRA, None),
            ('ban_quanluc', 'Ban Quân lực', Role.BAN_QUANLUC, None),
        ]

        for username, ho_ten, role, don_vi_id in users_data:
            u = User(username=username, ho_ten=ho_ten, role=role, don_vi_id=don_vi_id)
            u.set_password('123456')
            db.session.add(u)

        # Create sample unit accounts for some units
        sample_units = ['K1', 'K2', 'K3', 'TD1', 'TD2', 'B1']
        for ma in sample_units:
            dv = DonVi.query.filter_by(ma_don_vi=ma).first()
            if dv:
                u = User(
                    username=ma.lower(),
                    ho_ten=dv.ten_don_vi,
                    role=Role.UNIT_USER,
                    don_vi_id=dv.id,
                )
                u.set_password('123456')
                db.session.add(u)

        db.session.commit()
        print("Seeded admin, department, and sample unit accounts.")
        print("=" * 50)
        print("SAMPLE ACCOUNTS (password: 123456):")
        print("  admin          - Admin (Tuyen huan)")
        print("  phong_khoahoc  - Phong Khoa hoc")
        print("  phong_daotao   - Phong Dao tao")
        print("  tt_phong_chinhtri - Thu truong Phong Chinh tri")
        print("  tt_phong_tmhc  - Thu truong Phong TM-HC")
        print("  ban_canbo      - Ban Can bo")
        print("  ban_tochuc     - Ban To chuc")
        print("  ban_tuyenhuan  - Ban Tuyen huan")
        print("  ban_ctcq       - Ban Cong tac quan chung")
        print("  ban_cntt       - Ban Cong nghe thong tin")
        print("  ban_tachuan    - Ban Tac huan")
        print("  ban_baove_anninh - Ban Bao ve an ninh")
        print("  ban_khth       - Ban Ke hoach tong hop")
        print("  uyban_kiemtra  - Uy ban Kiem tra")
        print("  ban_quanluc    - Ban Quan luc")
        print("  k1, k2, k3     - Khoa (unit)")
        print("  td1, td2       - Tieu doan (unit)")
        print("  b1             - Phong Tham muu (unit)")
        print("=" * 50)
    else:
        print("Users already exist, skipping seed.")

    print("Done! Run 'python run.py' to start the server.")
