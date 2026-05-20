"""
Seed TieuChi (tiêu chí tập thể) for DVQT and DVTT danh hiệu.
Run: python scripts/seed_tieu_chi_tap_the.py
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import create_app
from app.extensions import db
from app.models.nomination import TieuChi, DanhHieu

app = create_app()

TIEU_CHI_TAP_THE = [
    # --- Ban Cán bộ ---
    {
        'ma_truong': 'cb_xeploai_canthu_pct_hoanttot',
        'ten': 'Ban CB: Tỷ lệ HTTNV cán bộ (%)',
        'nhom': 'ban_canbo',
        'loai_input': 'textbox',
        'huong_dan': 'Tỷ lệ % cán bộ xếp loại Hoàn thành tốt nhiệm vụ trở lên. ĐK QT: 100%, TT: 100%',
        'thu_tu': 10,
    },
    {
        'ma_truong': 'cb_xeploai_canthu_pct_xuatsac',
        'ten': 'Ban CB: Tỷ lệ HTT+XS cán bộ (%)',
        'nhom': 'ban_canbo',
        'loai_input': 'textbox',
        'huong_dan': 'Tỷ lệ % cán bộ xếp loại Hoàn thành tốt + Xuất sắc. ĐK QT: ≥95%, TT: ≥90%',
        'thu_tu': 11,
    },

    # --- Ban Tổ chức ---
    {
        'ma_truong': 'tc_xeploai_dangvien_pct_hoanttot',
        'ten': 'Ban TC: Tỷ lệ HTTNV đảng viên (%)',
        'nhom': 'ban_tochuc',
        'loai_input': 'textbox',
        'huong_dan': 'Tỷ lệ % đảng viên xếp loại Hoàn thành tốt nhiệm vụ trở lên.',
        'thu_tu': 20,
    },
    {
        'ma_truong': 'tc_xeploai_dangvien_pct_xuatsac',
        'ten': 'Ban TC: Tỷ lệ HTT+XS đảng viên (%)',
        'nhom': 'ban_tochuc',
        'loai_input': 'textbox',
        'huong_dan': 'Tỷ lệ % đảng viên xếp loại HTT+XS.',
        'thu_tu': 21,
    },
    {
        'ma_truong': 'tc_xeploai_tcdcs',
        'ten': 'Ban TC: Xếp loại TCĐCS',
        'nhom': 'ban_tochuc',
        'loai_input': 'combobox',
        'gia_tri_chon': ['Hoàn thành xuất sắc nhiệm vụ', 'Hoàn thành tốt nhiệm vụ', 'Hoàn thành nhiệm vụ', 'Không hoàn thành nhiệm vụ'],
        'huong_dan': 'ĐK QT: HTXSNV, TT: HTTNV',
        'thu_tu': 22,
    },

    # --- Ban Tuyên huấn ---
    {
        'ma_truong': 'th_diem_tdtx_quy1',
        'ten': 'Ban TH: Điểm TĐTX BQ Quý I',
        'nhom': 'ban_tuyenhuan',
        'loai_input': 'textbox',
        'huong_dan': 'Điểm tổng điểm tư tưởng xét bình quân Quý I. ĐK QT: ≥8.2, TT: ≥8.0',
        'thu_tu': 30,
    },
    {
        'ma_truong': 'th_diem_tdtx_quy2',
        'ten': 'Ban TH: Điểm TĐTX BQ Quý II',
        'nhom': 'ban_tuyenhuan',
        'loai_input': 'textbox',
        'huong_dan': 'Điểm tổng điểm tư tưởng xét bình quân Quý II.',
        'thu_tu': 31,
    },
    {
        'ma_truong': 'th_diem_tdtx_quy3',
        'ten': 'Ban TH: Điểm TĐTX BQ Quý III',
        'nhom': 'ban_tuyenhuan',
        'loai_input': 'textbox',
        'huong_dan': 'Điểm tổng điểm tư tưởng xét bình quân Quý III.',
        'thu_tu': 32,
    },
    {
        'ma_truong': 'th_diem_tdtx_quy4',
        'ten': 'Ban TH: Điểm TĐTX BQ Quý IV',
        'nhom': 'ban_tuyenhuan',
        'loai_input': 'textbox',
        'huong_dan': 'Điểm tổng điểm tư tưởng xét bình quân Quý IV.',
        'thu_tu': 33,
    },
    {
        'ma_truong': 'th_kq_ktra_ct_pct_dyc',
        'ten': 'Ban TH: Kết quả KT chính trị - % ĐYC',
        'nhom': 'ban_tuyenhuan',
        'loai_input': 'textbox',
        'huong_dan': 'Tỷ lệ % đạt yêu cầu trong kiểm tra chính trị.',
        'thu_tu': 34,
    },
    {
        'ma_truong': 'th_kq_ktra_ct_xeploai',
        'ten': 'Ban TH: Kết quả KT chính trị - Xếp loại',
        'nhom': 'ban_tuyenhuan',
        'loai_input': 'combobox',
        'gia_tri_chon': ['Giỏi', 'Khá', 'Đạt', 'Không đạt'],
        'huong_dan': 'Xếp loại chung kết quả kiểm tra chính trị.',
        'thu_tu': 35,
    },

    # --- Ban CTQC ---
    {
        'ma_truong': 'ctqc_xeploai_doanvien',
        'ten': 'Ban CTQC: Xếp loại đoàn viên',
        'nhom': 'ban_ctqc',
        'loai_input': 'combobox',
        'gia_tri_chon': ['Hoàn thành xuất sắc nhiệm vụ', 'Hoàn thành tốt nhiệm vụ', 'Hoàn thành nhiệm vụ', 'Không hoàn thành nhiệm vụ'],
        'huong_dan': 'Xếp loại tập thể đoàn viên.',
        'thu_tu': 40,
    },
    {
        'ma_truong': 'ctqc_xeploai_tcd',
        'ten': 'Ban CTQC: Xếp loại TCĐ',
        'nhom': 'ban_ctqc',
        'loai_input': 'combobox',
        'gia_tri_chon': ['Hoàn thành xuất sắc nhiệm vụ', 'Hoàn thành tốt nhiệm vụ', 'Hoàn thành nhiệm vụ', 'Không hoàn thành nhiệm vụ'],
        'huong_dan': 'Xếp loại tổ chức đoàn.',
        'thu_tu': 41,
    },
    {
        'ma_truong': 'ctqc_xeploai_hoivien_phunu',
        'ten': 'Ban CTQC: Xếp loại hội viên phụ nữ',
        'nhom': 'ban_ctqc',
        'loai_input': 'combobox',
        'gia_tri_chon': ['Hoàn thành xuất sắc nhiệm vụ', 'Hoàn thành tốt nhiệm vụ', 'Hoàn thành nhiệm vụ', 'Không hoàn thành nhiệm vụ'],
        'huong_dan': 'Xếp loại hội viên phụ nữ.',
        'thu_tu': 42,
    },
    {
        'ma_truong': 'ctqc_xeploai_hoi_phunu_coso',
        'ten': 'Ban CTQC: Xếp loại Hội PN cơ sở',
        'nhom': 'ban_ctqc',
        'loai_input': 'combobox',
        'gia_tri_chon': ['Xuất sắc', 'Vững mạnh', 'Khá', 'Trung bình', 'Yếu'],
        'huong_dan': 'Xếp loại Hội Phụ nữ cơ sở.',
        'thu_tu': 43,
    },

    # --- Ban CNTT ---
    {
        'ma_truong': 'cntt_chuyen_doi_so',
        'ten': 'Chuyển đổi số',
        'nhom': 'ban_cntt',
        'loai_input': 'combobox',
        'gia_tri_chon': ['Tốt', 'Chưa tốt'],
        'huong_dan': 'Đánh giá công tác chuyển đổi số của đơn vị.',
        'thu_tu': 50,
    },
    {
        'ma_truong': 'cntt_an_toan_thong_tin',
        'ten': 'An toàn thông tin',
        'nhom': 'ban_cntt',
        'loai_input': 'combobox',
        'gia_tri_chon': ['An toàn', 'Có vụ việc'],
        'huong_dan': 'Đánh giá an toàn thông tin trong năm.',
        'thu_tu': 51,
    },

    # --- Ban Tác huấn ---
    {
        'ma_truong': 'tachuan_an_toan_tuyet_doi',
        'ten': 'Ban Tác huấn: An toàn tuyệt đối',
        'nhom': 'ban_tachuan',
        'loai_input': 'combobox',
        'gia_tri_chon': ['Có', 'Không'],
        'huong_dan': 'Đơn vị có đảm bảo an toàn tuyệt đối không?',
        'thu_tu': 60,
    },
    {
        'ma_truong': 'tachuan_vmtd_mau_muc',
        'ten': 'Ban Tác huấn: VMTD Mẫu mực',
        'nhom': 'ban_tachuan',
        'loai_input': 'combobox',
        'gia_tri_chon': ['Đạt', 'Không đạt'],
        'huong_dan': 'Đánh giá Văn minh tác đồng mẫu mực.',
        'thu_tu': 61,
    },
    {
        'ma_truong': 'tachuan_dinh_luong',
        'ten': 'Ban Tác huấn: Định lượng - Xếp loại',
        'nhom': 'ban_tachuan',
        'loai_input': 'combobox',
        'gia_tri_chon': ['Giỏi', 'Khá', 'Đạt yêu cầu', 'Không đạt'],
        'huong_dan': 'Kết quả kiểm tra định lượng.',
        'thu_tu': 62,
    },
    {
        'ma_truong': 'tachuan_dinh_hinh',
        'ten': 'Ban Tác huấn: Định hình - Xếp loại',
        'nhom': 'ban_tachuan',
        'loai_input': 'combobox',
        'gia_tri_chon': ['Giỏi', 'Khá', 'Đạt yêu cầu', 'Không đạt'],
        'huong_dan': 'Kết quả kiểm tra định hình.',
        'thu_tu': 63,
    },
    {
        'ma_truong': 'tachuan_ban_sung_pct',
        'ten': 'Ban Tác huấn: Bắn súng - % đạt',
        'nhom': 'ban_tachuan',
        'loai_input': 'textbox',
        'huong_dan': 'Tỷ lệ % đạt yêu cầu bắn súng.',
        'thu_tu': 64,
    },
    {
        'ma_truong': 'tachuan_the_luc_pct',
        'ten': 'Ban Tác huấn: Thể lực - % đạt',
        'nhom': 'ban_tachuan',
        'loai_input': 'textbox',
        'huong_dan': 'Tỷ lệ % đạt yêu cầu thể lực.',
        'thu_tu': 65,
    },

    # --- Phòng Đào tạo ---
    {
        'ma_truong': 'dt_dinh_muc_ldsp_pct_vuot',
        'ten': 'P.ĐT: Định mức LĐSP - % vượt',
        'nhom': 'phong_daotao',
        'loai_input': 'textbox',
        'huong_dan': 'Tỷ lệ % vượt định mức lao động sư phạm.',
        'thu_tu': 70,
    },
    {
        'ma_truong': 'dt_cl_bai_giang_pct_dyc',
        'ten': 'P.ĐT: Chất lượng bài giảng - % ĐYC',
        'nhom': 'phong_daotao',
        'loai_input': 'textbox',
        'huong_dan': 'Tỷ lệ % bài giảng đạt yêu cầu.',
        'thu_tu': 71,
    },
    {
        'ma_truong': 'dt_cl_bai_giang_pct_kdat',
        'ten': 'P.ĐT: Chất lượng bài giảng - % không đạt',
        'nhom': 'phong_daotao',
        'loai_input': 'textbox',
        'huong_dan': 'Tỷ lệ % bài giảng không đạt.',
        'thu_tu': 72,
    },
    {
        'ma_truong': 'dt_kq_ktra_giang',
        'ten': 'P.ĐT: KQ kiểm tra giảng',
        'nhom': 'phong_daotao',
        'loai_input': 'combobox',
        'gia_tri_chon': ['Giỏi', 'Khá', 'Đạt yêu cầu', 'Không đạt'],
        'huong_dan': 'Kết quả kiểm tra giảng năm học.',
        'thu_tu': 73,
    },
    {
        'ma_truong': 'dt_gv_gioi_pct',
        'ten': 'P.ĐT: Giảng viên giỏi - % đạt',
        'nhom': 'phong_daotao',
        'loai_input': 'textbox',
        'huong_dan': 'Tỷ lệ % giảng viên đạt danh hiệu giảng viên giỏi.',
        'thu_tu': 74,
    },
    {
        'ma_truong': 'dt_kq_hoc_tap_pct_khagioi',
        'ten': 'P.ĐT: KQ học tập - % khá giỏi',
        'nhom': 'phong_daotao',
        'loai_input': 'textbox',
        'huong_dan': 'Tỷ lệ % học viên đạt khá giỏi.',
        'thu_tu': 75,
    },
    {
        'ma_truong': 'dt_kq_hoc_tap_pct_gioi',
        'ten': 'P.ĐT: KQ học tập - % giỏi',
        'nhom': 'phong_daotao',
        'loai_input': 'textbox',
        'huong_dan': 'Tỷ lệ % học viên đạt giỏi.',
        'thu_tu': 76,
    },
    {
        'ma_truong': 'dt_luan_van_sdh',
        'ten': 'P.ĐT: Luận văn SĐH - xếp loại',
        'nhom': 'phong_daotao',
        'loai_input': 'combobox',
        'gia_tri_chon': ['Tốt', 'Khá', 'Đạt yêu cầu', 'Không đạt'],
        'huong_dan': 'Kết quả luận văn sau đại học.',
        'thu_tu': 77,
    },
    {
        'ma_truong': 'dt_tieng_anh_sdh_pct',
        'ten': 'P.ĐT: Tiếng Anh SĐH - % đạt',
        'nhom': 'phong_daotao',
        'loai_input': 'textbox',
        'huong_dan': 'Tỷ lệ % học viên SĐH đạt yêu cầu tiếng Anh.',
        'thu_tu': 78,
    },
    {
        'ma_truong': 'dt_ren_luyen_hv_xeploai',
        'ten': 'P.ĐT: Rèn luyện HV - xếp loại',
        'nhom': 'phong_daotao',
        'loai_input': 'combobox',
        'gia_tri_chon': ['Tốt', 'Khá', 'Trung bình', 'Yếu'],
        'huong_dan': 'Xếp loại rèn luyện học viên.',
        'thu_tu': 79,
    },

    # --- Phòng KHQS ---
    {
        'ma_truong': 'kh_dinh_muc_ldkh_pct_vuot',
        'ten': 'P.KHQS: Định mức LĐKH - % vượt',
        'nhom': 'phong_khqs',
        'loai_input': 'textbox',
        'huong_dan': 'Tỷ lệ % vượt định mức lao động khoa học.',
        'thu_tu': 80,
    },
    {
        'ma_truong': 'kh_vuot_chi_tieu_nckh',
        'ten': 'P.KHQS: Vượt chỉ tiêu NCKH',
        'nhom': 'phong_khqs',
        'loai_input': 'combobox',
        'gia_tri_chon': ['Có', 'Không'],
        'huong_dan': 'Đơn vị có vượt chỉ tiêu NCKH không?',
        'thu_tu': 81,
    },
    {
        'ma_truong': 'kh_kq_nghiem_thu',
        'ten': 'P.KHQS: KQ nghiệm thu đề tài',
        'nhom': 'phong_khqs',
        'loai_input': 'combobox',
        'gia_tri_chon': ['Xuất sắc', 'Tốt', 'Khá', 'Đạt', 'Không đạt'],
        'huong_dan': 'Kết quả nghiệm thu đề tài khoa học.',
        'thu_tu': 82,
    },
    {
        'ma_truong': 'kh_giao_trinh_so_luong',
        'ten': 'P.KHQS: Giáo trình - số lượng',
        'nhom': 'phong_khqs',
        'loai_input': 'textbox',
        'huong_dan': 'Số lượng giáo trình biên soạn/nghiệm thu trong năm.',
        'thu_tu': 83,
    },
    {
        'ma_truong': 'kh_nckh_ca_nhan_so_luong',
        'ten': 'P.KHQS: NCKH cá nhân - số lượng',
        'nhom': 'phong_khqs',
        'loai_input': 'textbox',
        'huong_dan': 'Số lượng đề tài NCKH cá nhân.',
        'thu_tu': 84,
    },
    {
        'ma_truong': 'kh_bai_bao_so_luong',
        'ten': 'P.KHQS: Bài báo ThS/TS/PGS - số lượng',
        'nhom': 'phong_khqs',
        'loai_input': 'textbox',
        'huong_dan': 'Số lượng bài báo khoa học của ThS/TS/PGS.',
        'thu_tu': 85,
    },
    {
        'ma_truong': 'kh_sang_kien_hieu_qua',
        'ten': 'P.KHQS: Sáng kiến hiệu quả cao',
        'nhom': 'phong_khqs',
        'loai_input': 'combobox',
        'gia_tri_chon': ['Có', 'Không'],
        'huong_dan': 'Đơn vị có sáng kiến hiệu quả cao được ghi nhận?',
        'thu_tu': 86,
    },

    # --- Phòng HC-KT ---
    {
        'ma_truong': 'hckt_bao_dam_tieu_chuan',
        'ten': 'P.HC-KT: Bảo đảm tiêu chuẩn',
        'nhom': 'phong_hckt',
        'loai_input': 'combobox',
        'gia_tri_chon': ['Đảm bảo', 'Không đảm bảo'],
        'huong_dan': 'Đơn vị có đảm bảo các tiêu chuẩn hậu cần kỹ thuật?',
        'thu_tu': 90,
    },
    {
        'ma_truong': 'hckt_tgsx_tieu_doan',
        'ten': 'P.HC-KT: TGSX (chỉ tiểu đoàn)',
        'nhom': 'phong_hckt',
        'loai_input': 'combobox',
        'gia_tri_chon': ['Đạt', 'Không đạt'],
        'huong_dan': 'Kết quả tham gia sản xuất (áp dụng cho đơn vị cấp tiểu đoàn).',
        'thu_tu': 91,
    },
    {
        'ma_truong': 'hckt_phong_chong_dich',
        'ten': 'P.HC-KT: Phòng chống dịch',
        'nhom': 'phong_hckt',
        'loai_input': 'combobox',
        'gia_tri_chon': ['Tốt', 'Đạt yêu cầu', 'Chưa đạt'],
        'huong_dan': 'Đánh giá công tác phòng chống dịch bệnh.',
        'thu_tu': 92,
    },
    {
        'ma_truong': 'hckt_quan_so_khoe_pct',
        'ten': 'P.HC-KT: Quân số khỏe (%)',
        'nhom': 'phong_hckt',
        'loai_input': 'textbox',
        'huong_dan': 'Tỷ lệ % quân số khỏe. ĐK QT: >99.8%, TT: >99.7%',
        'thu_tu': 93,
    },
    {
        'ma_truong': 'hckt_csvc_xeploai',
        'ten': 'P.HC-KT: CSVC - xếp loại',
        'nhom': 'phong_hckt',
        'loai_input': 'combobox',
        'gia_tri_chon': ['Tốt', 'Khá', 'Đạt yêu cầu', 'Không đạt'],
        'huong_dan': 'Xếp loại cơ sở vật chất hậu cần kỹ thuật.',
        'thu_tu': 94,
    },
    {
        'ma_truong': 'hckt_an_toan_gt',
        'ten': 'P.HC-KT: An toàn giao thông',
        'nhom': 'phong_hckt',
        'loai_input': 'combobox',
        'gia_tri_chon': ['An toàn', 'Có vi phạm nhỏ', 'Có tai nạn'],
        'huong_dan': 'Đánh giá an toàn giao thông trong năm.',
        'thu_tu': 95,
    },

    # --- Ban Khảo thí ---
    {
        'ma_truong': 'kt_kq_ktra_giang_pct_khatot',
        'ten': 'Ban KT: KQ kiểm tra giảng - % khá tốt',
        'nhom': 'ban_khaothi',
        'loai_input': 'textbox',
        'huong_dan': 'Tỷ lệ % kết quả kiểm tra giảng đạt khá tốt.',
        'thu_tu': 100,
    },
    {
        'ma_truong': 'kt_kq_ktra_giang_pct_tot',
        'ten': 'Ban KT: KQ kiểm tra giảng - % tốt',
        'nhom': 'ban_khaothi',
        'loai_input': 'textbox',
        'huong_dan': 'Tỷ lệ % kết quả kiểm tra giảng đạt tốt.',
        'thu_tu': 101,
    },
]

# TieuChi ma_truong list for DVQT and DVTT
DVQT_TIEU_CHI = [tc['ma_truong'] for tc in TIEU_CHI_TAP_THE]
DVTT_TIEU_CHI = [tc['ma_truong'] for tc in TIEU_CHI_TAP_THE]


def seed():
    with app.app_context():
        created = 0
        updated = 0
        for tc_data in TIEU_CHI_TAP_THE:
            gia_tri = tc_data.pop('gia_tri_chon', None)
            existing = TieuChi.query.filter_by(ma_truong=tc_data['ma_truong']).first()
            if existing:
                for k, v in tc_data.items():
                    setattr(existing, k, v)
                if gia_tri:
                    existing.gia_tri_chon = gia_tri
                existing.is_active = True
                updated += 1
            else:
                tc = TieuChi(**tc_data)
                if gia_tri:
                    tc.gia_tri_chon = gia_tri
                tc.is_active = True
                db.session.add(tc)
                created += 1

        # Assign to DVQT and DVTT
        for ma in ['DVQT', 'DVTT']:
            dh = DanhHieu.query.filter_by(ma_danh_hieu=ma).first()
            if dh:
                dh.tieu_chi = DVQT_TIEU_CHI
                print(f'  Assigned {len(DVQT_TIEU_CHI)} tieu_chi to {ma}')
            else:
                print(f'  WARNING: DanhHieu {ma} not found!')

        db.session.commit()
        print(f'Done: {created} created, {updated} updated.')


if __name__ == '__main__':
    seed()
