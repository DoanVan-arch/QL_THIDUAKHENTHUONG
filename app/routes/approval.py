from flask import Blueprint, render_template, redirect, url_for, flash, request, jsonify, send_file
from flask_login import login_required, current_user
from app.extensions import db
from app.models.user import User, Role
from app.models.unit import DonVi
from app.models.nomination import DeXuat, DeXuatChiTiet, TrangThaiDeXuat, TrangThaiChiTiet, TieuChi
from app.models.approval import PheDuyet, PhongDuyet, KetQuaDuyet, KetQuaDuyetChiTiet
from app.models.notification import ThongBao
from app.models.edit_request import YeuCauChinhSua, TrangThaiYeuCauSua
from app.utils.decorators import department_required
from app.utils.activity_logger import log_action
from datetime import datetime
from io import BytesIO
from sqlalchemy.orm import joinedload, subqueryload
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

#thaydoi import theo nhu cau
approval_bp = Blueprint('approval', __name__)

# Scope-limited depts that need auto-finalize when all their items are out-of-scope
_SCOPE_LIMITED_PHONGS = [
    PhongDuyet.BAN_QUANLUC.value,
    PhongDuyet.BAN_CANBO.value,
    PhongDuyet.PHONG_HAUCANKYTHUAT.value,
    PhongDuyet.BAN_SAUDAIHOC.value,
]

def _auto_finalize_scope_dept(de_xuat_id):
    """For BAN_QUANLUC and BAN_CANBO: if a PheDuyet has ALL chi_tiets out-of-scope,
    auto-create KetQuaDuyetChiTiet = DONG_Y for each and finalize PheDuyet.ket_qua = DONG_Y.
    Safe to call multiple times (idempotent).
    Returns list of auto-finalized phong_duyet names.
    """
    from app.models.nomination import DeXuat as _DX
    de_xuat = _DX.query.get(de_xuat_id)
    if not de_xuat:
        return []
    
    finalized = []
    for phong_val in _SCOPE_LIMITED_PHONGS:
        pd = PheDuyet.query.filter_by(
            de_xuat_id=de_xuat_id,
            phong_duyet=phong_val,
        ).first()
        
        if not pd or pd.ket_qua == KetQuaDuyet.DONG_Y.value:
            continue  # already done or not created yet
            
        # Determine scope role for this phong
        scope_role = _PHONG_TO_ROLE.get(phong_val)
        if scope_role is None:
            continue
            
        # Ensure KetQuaDuyetChiTiet records exist for all chi_tiets
        existing = {kq.chi_tiet_id for kq in pd.chi_tiet_duyet}
        
        # 1. Lấy toàn bộ ID chi tiết để truy vấn một lần duy nhất (Tránh N+1 Query)
        chi_tiet_ids = [ct.id for ct in de_xuat.chi_tiets]

        # 2. Lấy ra tất cả các kết quả duyệt hiện có của phe_duyet_id này dưới dạng Dictionary
        existing_records_list = KetQuaDuyetChiTiet.query.filter(
            KetQuaDuyetChiTiet.phe_duyet_id == pd.id,
            KetQuaDuyetChiTiet.chi_tiet_id.in_(chi_tiet_ids)
        ).all()
        existing_dict = {kq.chi_tiet_id: kq for kq in existing_records_list}

        # 3. Xử lý logic
        for ct in de_xuat.chi_tiets:
            if ct.bi_loai:
                continue
                
            in_scope = _is_in_dept_scope(scope_role, ct.doi_tuong)
            target_ket_qua = None
            skip_record = False

            # --- XÁC ĐỊNH TRẠNG THÁI (KẾT QUẢ) MONG MUỐN ---
            if ct.doi_tuong is None or ct.quan_nhan_id is None:
                if phong_val == PhongDuyet.BAN_SAUDAIHOC.value:
                    target_ket_qua = KetQuaDuyet.CHO_DUYET.value if in_scope else KetQuaDuyet.DONG_Y.value
                elif phong_val == PhongDuyet.PHONG_HAUCANKYTHUAT.value:
                    # Gán chờ duyệt và KHÔNG bỏ qua (skip_record vẫn là False)
                    target_ket_qua = KetQuaDuyet.CHO_DUYET.value
                else:
                    # Chỉ bỏ qua (skip) tập thể cho các phòng khác
                    skip_record = True
                    
            else:
                # Nếu không phải Phòng HCKT và không phải Ban SĐH
                if phong_val not in (PhongDuyet.PHONG_HAUCANKYTHUAT.value, PhongDuyet.BAN_SAUDAIHOC.value):
                    target_ket_qua = KetQuaDuyet.CHO_DUYET.value if in_scope else KetQuaDuyet.DONG_Y.value
                else:
                    # Nếu là Phòng HCKT hoặc Ban SĐH
                    if ct.doi_tuong == 'Học viên sau đại học':
                        target_ket_qua = KetQuaDuyet.CHO_DUYET.value
                    else:
                        target_ket_qua = KetQuaDuyet.DONG_Y.value

            # Nếu thuộc diện bỏ qua (tập thể của phòng khác) thì chuyển sang chi tiết tiếp theo
            if skip_record:
                continue

            # --- ÁP DỤNG THAY ĐỔI VÀO DATABASE (KHÔNG GỌI QUERY NỮA) ---
            if ct.id in existing_dict:
                # Đã tồn tại -> Cập nhật nếu giá trị khác với kết quả mong muốn
                ket_qua_hien_tai = existing_dict[ct.id]
                if ket_qua_hien_tai.ket_qua != target_ket_qua:
                    ket_qua_hien_tai.ket_qua = target_ket_qua
                    # KHÔNG commit ở đây, chờ làm xong hết mới flush
            else:
                # Chưa tồn tại -> Thêm mới
                ket_qua_moi = KetQuaDuyetChiTiet(
                    phe_duyet_id=pd.id,
                    chi_tiet_id=ct.id,
                    ket_qua=target_ket_qua
                )
                db.session.add(ket_qua_moi)

        # 4. Lưu tất cả thay đổi xuống database trong 1 lần duy nhất
        db.session.flush() 
        
        # Re-check: if no CHO_DUYET remains among ACTIVE items → auto-finalize
        pending = KetQuaDuyetChiTiet.query.filter_by(
            phe_duyet_id=pd.id,
            ket_qua=KetQuaDuyet.CHO_DUYET.value,
        ).join(DeXuatChiTiet, KetQuaDuyetChiTiet.chi_tiet_id == DeXuatChiTiet.id).filter(
            DeXuatChiTiet.bi_loai == False
        ).count()
        
        if pending == 0:
            pd.ket_qua = KetQuaDuyet.DONG_Y.value
            pd.ngay_duyet = datetime.utcnow()
            pd.ghi_chu = 'Tự động duyệt (không có đối tượng thuộc phạm vi)'
            finalized.append(phong_val)
            
    db.session.commit()
    return finalized

ROLE_TO_PHONG = {
    Role.PHONG_KHOAHOC: PhongDuyet.PHONG_KHOAHOC.value,
    Role.PHONG_DAOTAO: PhongDuyet.PHONG_DAOTAO.value,
    Role.THU_TRUONG_PHONG_TMHC: PhongDuyet.THU_TRUONG_PHONG_TMHC.value,
    Role.BAN_CANBO: PhongDuyet.BAN_CANBO.value,
    Role.BAN_TOCHUC: PhongDuyet.BAN_TOCHUC.value,
    Role.BAN_TUYENHUAN: PhongDuyet.BAN_TUYENHUAN.value,
    Role.BAN_CTCQ: PhongDuyet.BAN_CTCQ.value,
    Role.BAN_CNTT: PhongDuyet.BAN_CNTT.value,
    Role.BAN_TAC_HUAN: PhongDuyet.BAN_TAC_HUAN.value,
    Role.BAN_KHAOTHI: PhongDuyet.BAN_KHAOTHI.value,
    Role.BAN_BAOVE_ANNINH: PhongDuyet.BAN_BAOVE_ANNINH.value,
    Role.UY_BAN_KIEMTRA: PhongDuyet.UY_BAN_KIEMTRA.value,
    Role.BAN_QUANLUC: PhongDuyet.BAN_QUANLUC.value,
    Role.PHONG_HAUCANKYTHUAT: PhongDuyet.PHONG_HAUCANKYTHUAT.value,
    Role.BAN_SAUDAIHOC: PhongDuyet.BAN_SAUDAIHOC.value,
}


def _managed_gate_columns(role):
    if role == Role.THU_TRUONG_PHONG_TMHC:
        return [
            PhongDuyet.BAN_CNTT.value,
            PhongDuyet.BAN_TAC_HUAN.value,
            PhongDuyet.BAN_QUANLUC.value,
        ]
    return []

_GROUP_CONFIRMATION = {
    Role.THU_TRUONG_PHONG_TMHC: {
        PhongDuyet.BAN_CNTT.value,
        PhongDuyet.BAN_TAC_HUAN.value,
        PhongDuyet.BAN_QUANLUC.value,
    },
}


def _get_group_gate_for_pd(role, de_xuat_id):
    """Return gate status for group-confirmation roles.
    Output: {
      can_review: bool,
      required: [dept_names],
      approved: [dept_names],
      pending: [dept_names],
      rejected: [dept_names],
      results: {dept_name: ket_qua}
    }
    """
    if role not in _GROUP_CONFIRMATION:
        return {
            'can_review': True,
            'required': [],
            'approved': [],
            'pending': [],
            'rejected': [],
            'results': {},
        }

    required_groups = sorted(list(_GROUP_CONFIRMATION[role]))
    group_reviews = PheDuyet.query.filter(
        PheDuyet.de_xuat_id == de_xuat_id,
        PheDuyet.phong_duyet.in_(required_groups)
    ).all()
    result_map = {g.phong_duyet: g.ket_qua for g in group_reviews}

    approved = [d for d in required_groups if result_map.get(d) == KetQuaDuyet.DONG_Y.value]
    rejected = [d for d in required_groups if result_map.get(d) == KetQuaDuyet.TU_CHOI.value]
    pending = [d for d in required_groups if result_map.get(d) != KetQuaDuyet.DONG_Y.value]

    return {
        'can_review': len(pending) == 0,
        'required': required_groups,
        'approved': approved,
        'pending': pending,
        'rejected': rejected,
        'results': result_map,
    }


def _get_group_gate_for_ct(role, de_xuat_id, ct_id):
    """Return per-individual gate status for group-confirmation roles."""
    if role not in _GROUP_CONFIRMATION:
        return {
            'can_review': True,
            'required': [],
            'approved': [],
            'pending': [],
            'rejected': [],
            'results': {},
        }

    required_groups = sorted(list(_GROUP_CONFIRMATION[role]))

    # Query both the dept-level ket_qua and per-ct ket_qua in one shot
    rows = db.session.query(
        PheDuyet.phong_duyet,
        PheDuyet.ket_qua.label('pd_ket_qua'),
        KetQuaDuyetChiTiet.ket_qua.label('ct_ket_qua')
    ).outerjoin(
        KetQuaDuyetChiTiet,
        (KetQuaDuyetChiTiet.phe_duyet_id == PheDuyet.id) & (KetQuaDuyetChiTiet.chi_tiet_id == ct_id)
    ).filter(
        PheDuyet.de_xuat_id == de_xuat_id,
        PheDuyet.phong_duyet.in_(required_groups)
    ).all()

    # A dept is "approved" for this ct if:
    # 1. Per-ct KetQuaDuyetChiTiet.ket_qua == DONG_Y, OR
    # 2. PheDuyet.ket_qua == DONG_Y (dept fully finalized — all cts processed/auto-approved)
    result_map = {}
    for phong, pd_ket_qua, ct_ket_qua in rows:
        if ct_ket_qua == KetQuaDuyet.DONG_Y.value:
            result_map[phong] = KetQuaDuyet.DONG_Y.value
        elif pd_ket_qua == KetQuaDuyet.DONG_Y.value:
            # Dept fully approved → treat this ct as approved by that dept
            result_map[phong] = KetQuaDuyet.DONG_Y.value
        elif ct_ket_qua == KetQuaDuyet.TU_CHOI.value or pd_ket_qua == KetQuaDuyet.TU_CHOI.value:
            result_map[phong] = KetQuaDuyet.TU_CHOI.value
        else:
            result_map[phong] = ct_ket_qua  # None or CHO_DUYET

    approved = [d for d in required_groups if result_map.get(d) == KetQuaDuyet.DONG_Y.value]
    rejected = [d for d in required_groups if result_map.get(d) == KetQuaDuyet.TU_CHOI.value]
    pending = [d for d in required_groups if result_map.get(d) != KetQuaDuyet.DONG_Y.value]

    return {
        'can_review': len(pending) == 0,
        'required': required_groups,
        'approved': approved,
        'pending': pending,
        'rejected': rejected,
        'results': result_map,
    }

# Reverse map: PhongDuyet display name -> Role
_PHONG_TO_ROLE = {v: k for k, v in ROLE_TO_PHONG.items()}

# --- Hardcoded fallbacks (used when DB has no tieu_chi rows) ---
_FALLBACK_PHONG_FIELDS = {
    Role.PHONG_DAOTAO: [
        'danh_hieu_gv_gioi', 'tien_do_pgs', 'dinh_muc_giang_day',
        'ket_qua_kiem_tra_giang', 'danh_hieu_hv_gioi', 'diem_tong_ket', 'ket_qua_thuc_hanh',
    ],
    Role.PHONG_KHOAHOC: ['thoi_gian_lao_dong_kh', 'diem_nckh', 'nckh_noi_dung', 'nckh_minh_chung'],
    Role.PHONG_THAMMUU: [
        'kiem_tra_tin_hoc', 'kiem_tra_dieu_lenh', 'dia_ly_quan_su', 'ban_sung', 'the_luc', 'ket_qua_doan_the',
    ],
    Role.PHONG_CHINHTRI: ['kiem_tra_chinh_tri', 'ket_qua_doan_the', 'xep_loai_dang_vien'],
    Role.BAN_CANBO: ['muc_do_hoan_thanh'],
    Role.BAN_QUANLUC: ['muc_do_hoan_thanh'],
    Role.PHONG_HAUCANKYTHUAT: ['muc_do_hoan_thanh'],
    Role.BAN_SAUDAIHOC: ['muc_do_hoan_thanh'],
}

_FALLBACK_FIELD_LABELS = {
    'muc_do_hoan_thanh': 'Hoàn thành NV', 'phieu_tin_nhiem': 'Tín nhiệm',
    'kiem_tra_dieu_lenh': 'Điều lệnh', 'ban_sung': 'Bắn súng', 'the_luc': 'Thể lực',
    'kiem_tra_chinh_tri': 'Chính trị', 'kiem_tra_tin_hoc': 'Kỹ năng số',
    'dia_ly_quan_su': 'Địa hình QS', 'danh_hieu_gv_gioi': 'GV giỏi',
    'xep_loai_dang_vien': 'Xếp loại ĐV', 'xep_loai_doan_vien': 'Xếp loại đoàn viên',
    'hinh_thuc_khen_thuong_qc': 'KT quần chúng', 'ket_qua_phu_nu': 'XL phụ nữ',
    'hinh_thuc_khen_thuong_pn': 'KT phụ nữ',
    'dinh_muc_giang_day': 'Định mức GD', 'ket_qua_kiem_tra_giang': 'KT giảng',
    'thoi_gian_lao_dong_kh': 'LĐ KH', 'tien_do_pgs': 'Tiến độ PGS',
    'danh_hieu_hv_gioi': 'HV giỏi', 'diem_tong_ket': 'Điểm TK',
    'ket_qua_thuc_hanh': 'Thực hành', 'ket_qua_ren_luyen': 'KQ rèn luyện',
    'ket_qua_doan_the': 'Đoàn thể', 'hinh_thuc_tot_nghiep': 'HT thi TN',
    'diem_tn_ctd': 'Điểm CTĐ (TN)', 'diem_tn_ct': 'Điểm CT (TN)',
    'diem_tn_ta': 'Điểm TA (TN)', 'diem_tn_mon4': 'Điểm môn 4 (TN)',
    'diem_tn_chuyennganh': 'Điểm CN (TN)', 'diem_tn_baove': 'Điểm BV KL (TN)',
    'chu_tri_don_vi_danh_hieu': 'Chủ trì ĐV', 'diem_nckh': 'Điểm KH',
    'nckh_noi_dung': 'NCKH', 'nckh_minh_chung': 'MC NCKH',
    'mo_ta_khoa_hoc': 'Mô tả TT KH',
    'diem_tot_nghiep': 'Điểm TN (TB)',
    'minh_chung_thanh_tich_khac': 'MC TT khác',
    'thanh_tich_ca_nhan_khac': 'Thành tích khác',
}

# Long text / file fields excluded from table columns
_LONG_TEXT_FIELDS = {'nckh_noi_dung', 'nckh_minh_chung', 'tien_do_pgs', 'thanh_tich_ca_nhan_khac'}

# Roles that view ALL criteria (read-only oversight), regardless of their assigned phong_duyet mapping.
# They can still only approve/reject the items themselves; this only affects which columns are shown.
_VIEW_ALL_CRITERIA_ROLES = {
    Role.BAN_CANBO,
    Role.BAN_CTCQ,
    Role.BAN_BAOVE_ANNINH,
    Role.BAN_TOCHUC,
    Role.BAN_TUYENHUAN,
}


def _all_criteria_columns():
    """Return list of all ca_nhan criteria fields — lấy động từ bảng TieuChi,
    chỉ giữ các field thực sự tồn tại trên DeXuatChiTiet.
    Bao gồm cả long-text và file fields (không loại trừ nữa)."""
    from app.models.nomination import DeXuatChiTiet as _DX, TieuChi

    # Tập hợp tất cả cột thực tế trên bảng DeXuatChiTiet
    _all_cols = {c.name for c in _DX.__table__.columns}

    # Các field hệ thống — không phải tiêu chí, luôn loại trừ
    _SYSTEM_FIELDS = {
        'id', 'de_xuat_id', 'quan_nhan_id', 'loai_danh_hieu',
        'doi_tuong', 'ten_don_vi_de_xuat', 'ghi_chu',
        'bi_loai', 'trang_thai', 'ly_do_tu_choi',
        'created_at', 'updated_at', 'tap_the_data','ly_do_loai','diem_nckh','xep_loai_tong_ket','diem_tot_nghiep','xep_loai_doan_vien','diem_the_luc','ngay_loai','admin_approved','phong_loai','diem_kiem_tra_tin_hoc','nam_hoc','ly_do_loai',
    }

   

    # Fallback nếu TieuChi chưa có dữ liệu
   

    # Chỉ loại trừ field hệ thống, giữ lại tất cả tiêu chí kể cả long-text/file
    return [f for f  in _all_cols if f not in _SYSTEM_FIELDS]







def _load_phong_fields_from_db():
    """Build PHONG_FIELDS dict from TieuChi table. Returns None if table is empty."""
    rows = TieuChi.query.filter_by(is_active=True).all()
    if not rows:
        return None
    result = {}
    for tc in rows:
        for pd_name in tc.phong_duyet:
            role = _PHONG_TO_ROLE.get(pd_name)
            if role:
                result.setdefault(role, [])
                if tc.ma_truong not in result[role]:
                    result[role].append(tc.ma_truong)
    return result


def _load_field_labels_from_db():
    """Build FIELD_LABELS dict from TieuChi table. Returns None if table is empty."""
    rows = TieuChi.query.filter_by(is_active=True).all()
    if not rows:
        return None
    return {tc.ma_truong: tc.ten for tc in rows}


def get_phong_fields():
    """Get department -> fields mapping, preferring DB over hardcoded fallback."""
    result = _load_phong_fields_from_db()
    return result if result else _FALLBACK_PHONG_FIELDS


def get_field_labels():
    """Get field -> label mapping, preferring DB over hardcoded fallback."""
    result = _load_field_labels_from_db()
    return result if result else _FALLBACK_FIELD_LABELS


def get_phong_table_columns():
    """Get department -> table column fields (excluding long text/file fields and collective-only fields)."""
    from app.models.nomination import DeXuatChiTiet as _DX
    _ca_nhan_cols = {c.key for c in _DX.__table__.columns}
    phong_fields = get_phong_fields()
    return {role: [f for f in fields if f not in _LONG_TEXT_FIELDS and f in _ca_nhan_cols]
            for role, fields in phong_fields.items()}


# Conditional field visibility by doi_tuong (remains hardcoded — specific to business logic)
PHONG_FIELD_CONDITIONS = {
    Role.BAN_CANBO: {
        'muc_do_hoan_thanh': ['Giảng viên', 'Cán bộ'],
    },
    Role.BAN_QUANLUC: {
        'muc_do_hoan_thanh': ['Công nhân viên', 'Quân nhân chuyên nghiệp','Hạ sĩ quan chiến sĩ'],
    },
}

# doi_tuong scope: which doi_tuong values each department approves
# Departments not listed approve ALL doi_tuong values
BAN_QUANLUC_DOI_TUONG = ['Công nhân viên', 'Quân nhân chuyên nghiệp','Hạ sĩ quan chiến sĩ']

DEPT_DOI_TUONG_SCOPE = {
    Role.BAN_QUANLUC: BAN_QUANLUC_DOI_TUONG,
    # BAN_CANBO approves all EXCEPT BAN_QUANLUC's scope
}


def _is_in_dept_scope(role, doi_tuong):
    """Check if an individual's doi_tuong falls within the department's approval scope."""
    # Tập thể (doi_tuong = None/'') → tất cả phòng ban đều xét duyệt
    if not doi_tuong:
        return True
    if role == Role.BAN_QUANLUC:
        return doi_tuong in BAN_QUANLUC_DOI_TUONG
    elif role == Role.BAN_CANBO:
        return doi_tuong not in BAN_QUANLUC_DOI_TUONG
    return True  # All other departments approve all doi_tuong


def _notify_rejections(phe_duyet):
    """Create ThongBao notifications for unit account when individuals are rejected."""
    de_xuat = phe_duyet.de_xuat
    # Find the unit user account for this don_vi
    unit_user = User.query.filter_by(
        don_vi_id=de_xuat.don_vi_id, role=Role.UNIT_USER
    ).first()
    if not unit_user:
        return

    phong_name = phe_duyet.phong_duyet
    for kq in phe_duyet.chi_tiet_duyet:
        if kq.ket_qua == KetQuaDuyet.TU_CHOI.value:
            ct = kq.chi_tiet
            name = ct.quan_nhan.ho_ten if ct.quan_nhan else de_xuat.don_vi.ten_don_vi
            thong_bao = ThongBao(
                user_id=unit_user.id,
                de_xuat_id=de_xuat.id,
                chi_tiet_id=ct.id,
                loai='tu_choi',
                tieu_de=f'{phong_name} không nhất trí: {name}',
                noi_dung=f'Lý do: {kq.ly_do or "Không rõ"}. Đề xuất năm học {de_xuat.nam_hoc}.',
            )
            db.session.add(thong_bao)


def _remove_chi_tiet_on_reject(phe_duyet, ct_id, ly_do=''):
    """When a department rejects a single cá nhân/tập thể, remove ONLY that item from the
    active approval process (soft-remove). The rest of the đề xuất continues unaffected.
    The unit is notified so it can see who was removed and why."""
    de_xuat = phe_duyet.de_xuat
    ct = DeXuatChiTiet.query.get(ct_id)
    if ct and not ct.bi_loai:
        ct.bi_loai = True
        ct.ly_do_loai = ly_do or 'Không đạt yêu cầu'
        ct.phong_loai = phe_duyet.phong_duyet
        ct.ngay_loai = datetime.utcnow()

    # Notify the unit account
    unit_user = User.query.filter_by(
        don_vi_id=de_xuat.don_vi_id, role=Role.UNIT_USER
    ).first()
    if unit_user and ct:
        name = (ct.quan_nhan.ho_ten if ct.quan_nhan else
                (ct.ten_don_vi_de_xuat or de_xuat.don_vi.ten_don_vi))
        thong_bao = ThongBao(
            user_id=unit_user.id,
            de_xuat_id=de_xuat.id,
            chi_tiet_id=ct_id,
            loai='tu_choi',
            tieu_de=f'{phe_duyet.phong_duyet} loại khỏi đề xuất: {name}',
            noi_dung=(f'Lý do: {ly_do or "Không đạt yêu cầu"}. '
                      f'{name} đã bị loại khỏi đề xuất năm học {de_xuat.nam_hoc} '
                      f'của {de_xuat.don_vi.ten_don_vi}. Các cá nhân/tập thể còn lại vẫn tiếp tục được xét duyệt.'),
        )
        db.session.add(thong_bao)

    # Recompute đề xuất status now that this item no longer participates
    _recompute_de_xuat_status(de_xuat)


def _recompute_de_xuat_status(de_xuat):
    """Recompute each department's finalization and the overall đề xuất status,
    counting ONLY active (non-removed) chi_tiets. Removed items never block or reject.
    A department finalizes as DONG_Y once none of its active in-scope items are still pending.
    """
    active_ct_ids = {ct.id for ct in de_xuat.chi_tiets if not ct.bi_loai}

    dept_pds = PheDuyet.query.filter_by(de_xuat_id=de_xuat.id).filter(
        PheDuyet.phong_duyet != PhongDuyet.ADMIN_TUYENHUAN.value
    ).all()

    # If there are no active items left at all, reject the whole đề xuất.
    if not active_ct_ids:
        de_xuat.trang_thai = TrangThaiDeXuat.TU_CHOI.value
        return

    for pd in dept_pds:
        if pd.ket_qua == KetQuaDuyet.DONG_Y.value:
            continue
        # KetQuaDuyetChiTiet records are created lazily (when a dept first visits review_nomination).
        # If there are no records at all, this dept has never started reviewing → do NOT auto-promote.
        all_records = [kq for kq in pd.chi_tiet_duyet if kq.chi_tiet_id in active_ct_ids]
        if not all_records:
            continue  # dept hasn't reviewed yet; skip
        # Pending = active items in this dept still waiting for a decision
        pending = [kq for kq in all_records if kq.ket_qua == KetQuaDuyet.CHO_DUYET.value]
        if not pending:
            pd.ket_qua = KetQuaDuyet.DONG_Y.value
            if pd.ngay_duyet is None:
                pd.ngay_duyet = datetime.utcnow()

    db.session.flush()

    # All departments approved → advance to Hội đồng
    all_done = all(p.ket_qua == KetQuaDuyet.DONG_Y.value for p in dept_pds)
    if all_done and dept_pds:
        if de_xuat.trang_thai not in (TrangThaiDeXuat.HOI_DONG.value,
                                      TrangThaiDeXuat.PHE_DUYET_CUOI.value):
            de_xuat.trang_thai = TrangThaiDeXuat.HOI_DONG.value
        existing_admin = PheDuyet.query.filter_by(
            de_xuat_id=de_xuat.id, phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value
        ).first()
        if not existing_admin:
            db.session.add(PheDuyet(
                de_xuat_id=de_xuat.id,
                phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value,
                ket_qua=KetQuaDuyet.CHO_DUYET.value,
            ))
    else:
        if de_xuat.trang_thai == TrangThaiDeXuat.TU_CHOI.value:
            de_xuat.trang_thai = TrangThaiDeXuat.DANG_DUYET.value

    # Update per-item trang_thai to match parent de_xuat stage
    _recompute_chi_tiet_status(de_xuat)


def _recompute_chi_tiet_status(de_xuat):
    """Sync each active DeXuatChiTiet.trang_thai with the parent đề xuất stage.

    Mapping:
      de_xuat NHAP/CHO_DUYET/DANG_DUYET → chi_tiet DANG_DUYET  (submitted, under review)
      de_xuat HOI_DONG                   → chi_tiet DA_DUYET    (all depts approved, Bảng 1)
      de_xuat PHE_DUYET_CUOI             → chi_tiet HOI_DONG    (admin_approved, Bảng 2/3)
      bi_loai = True                     → chi_tiet TU_CHOI
    Each individual item's admin_approved flag further promotes it to PHE_DUYET_CUOI.
    """
    dx_tt = de_xuat.trang_thai
    for ct in de_xuat.chi_tiets:
        if ct.bi_loai:
            ct.trang_thai = TrangThaiChiTiet.TU_CHOI.value
         
            continue
        if dx_tt in (TrangThaiDeXuat.NHAP.value, TrangThaiDeXuat.CHO_DUYET.value,
                     TrangThaiDeXuat.DANG_DUYET.value, TrangThaiDeXuat.TU_CHOI.value):
            ct.trang_thai = TrangThaiChiTiet.DANG_DUYET.value
            ct.ly_do_loai = None
            ct.phong_loai = None
            ct.ngay_loai = None
            ct.bi_loai = False
        elif dx_tt == TrangThaiDeXuat.HOI_DONG.value:
            if ct.admin_approved:
                ct.trang_thai = TrangThaiChiTiet.HOI_DONG.value
            else:
                ct.trang_thai = TrangThaiChiTiet.DA_DUYET.value
        elif dx_tt == TrangThaiDeXuat.PHE_DUYET_CUOI.value:
            ct.trang_thai = TrangThaiChiTiet.HOI_DONG.value

def _auto_prepare_pending(phong_name, nam_hoc_filter):
    """Tạo KetQuaDuyetChiTiet còn thiếu và auto-finalize nếu đủ điều kiện.
    Commit 1 lần duy nhất. Không trả về gì."""
    from app.models.nomination import DeXuat as _DeXuat

    pds = PheDuyet.query.filter_by(
        phong_duyet=phong_name,
        ket_qua=KetQuaDuyet.CHO_DUYET.value
    ).options(
        subqueryload(PheDuyet.chi_tiet_duyet),
        joinedload(PheDuyet.de_xuat).subqueryload(_DeXuat.chi_tiets),
    ).join(_DeXuat, PheDuyet.de_xuat_id == _DeXuat.id)\
     .filter(_DeXuat.nam_hoc == nam_hoc_filter).all()

    new_kq_list = []
    is_auto_dept = phong_name in (
        PhongDuyet.PHONG_HAUCANKYTHUAT.value,
        PhongDuyet.BAN_SAUDAIHOC.value,
    )

    for pd in pds:
        existing_ct_ids = {kq.chi_tiet_id for kq in pd.chi_tiet_duyet}
        for ct in pd.de_xuat.chi_tiets:
            if ct.bi_loai or ct.id in existing_ct_ids:
                continue
            in_scope = _is_in_dept_scope(
                next(r for r, p in ROLE_TO_PHONG.items() if p == phong_name),
                ct.doi_tuong
            )
            ket_qua_val = (
                KetQuaDuyet.DONG_Y.value
                if (is_auto_dept or not in_scope)
                else KetQuaDuyet.CHO_DUYET.value
            )
            new_kq_list.append(KetQuaDuyetChiTiet(
                phe_duyet_id=pd.id,
                chi_tiet_id=ct.id,
                ket_qua=ket_qua_val,
            ))

    if new_kq_list:
        db.session.add_all(new_kq_list)
        db.session.flush()

    need_commit = bool(new_kq_list)

    for pd in pds:
        if pd.ket_qua != KetQuaDuyet.CHO_DUYET.value:
            continue
        active_ct_ids = {ct.id for ct in pd.de_xuat.chi_tiets if not ct.bi_loai}
        if not active_ct_ids:
            continue
        all_kq = list(pd.chi_tiet_duyet) + [
            kq for kq in new_kq_list if kq.phe_duyet_id == pd.id
        ]
        kq_map = {kq.chi_tiet_id: kq.ket_qua for kq in all_kq}
        if any(kq_map.get(ct_id) == KetQuaDuyet.CHO_DUYET.value for ct_id in active_ct_ids):
            continue
        pd.ket_qua        = KetQuaDuyet.DONG_Y.value
        pd.nguoi_duyet_id = None
        pd.ngay_duyet     = datetime.utcnow()
        pd.ghi_chu        = 'Tự động duyệt (không có đối tượng thuộc phạm vi)'
        need_commit       = True
        de_xuat  = pd.de_xuat
        all_dept = PheDuyet.query.filter_by(de_xuat_id=de_xuat.id).filter(
            PheDuyet.phong_duyet != PhongDuyet.ADMIN_TUYENHUAN.value
        ).all()
        if all(a.ket_qua == KetQuaDuyet.DONG_Y.value for a in all_dept):
            de_xuat.trang_thai = TrangThaiDeXuat.HOI_DONG.value
            if not PheDuyet.query.filter_by(
                de_xuat_id=de_xuat.id,
                phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value
            ).first():
                db.session.add(PheDuyet(
                    de_xuat_id=de_xuat.id,
                    phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value,
                    ket_qua=KetQuaDuyet.CHO_DUYET.value,
                ))

    if need_commit:
        db.session.commit()

@approval_bp.route('/pending')
@login_required
@department_required
def pending_list():
    phong_name = ROLE_TO_PHONG.get(current_user.role, '')
    nam_hoc_filter = request.args.get('nam_hoc', '')

    # Get available nam_hoc options for the dropdown — all years this phong has PheDuyet records
    from app.models.nomination import DeXuat as _DeXuat
    nam_hoc_list = [n[0] for n in db.session.query(_DeXuat.nam_hoc).join(
        PheDuyet, PheDuyet.de_xuat_id == _DeXuat.id
    ).filter(
        PheDuyet.phong_duyet == phong_name,
    ).distinct().order_by(_DeXuat.nam_hoc.desc()).all()]

    q = PheDuyet.query.filter_by(
        phong_duyet=phong_name
        # ket_qua=KetQuaDuyet.CHO_DUYET.value
    )
    if nam_hoc_filter:
        q = q.join(_DeXuat, PheDuyet.de_xuat_id == _DeXuat.id).filter(
            _DeXuat.nam_hoc == nam_hoc_filter
        )
    else:
        q = q.join(_DeXuat, PheDuyet.de_xuat_id == _DeXuat.id)
    
    # Sort by unit hierarchy (Phòng > Khoa > Đơn vị)
    from app.models.unit import DonVi
    q = q.join(DonVi, _DeXuat.don_vi_id == DonVi.id)
    pending_reviews = q.order_by(DonVi.thu_tu.asc(), _DeXuat.ngay_gui.desc()).all()
    # Filter out orphaned PheDuyet (de_xuat đã bị xóa khỏi DB) and đề xuất
    # whose cá nhân/tập thể have all been removed (bi_loai).
    pending_reviews = [
        pd for pd in pending_reviews
        if pd.de_xuat is not None and any(not ct.bi_loai for ct in pd.de_xuat.chi_tiets)
    ]

    # Ensure per-item records exist for all chi_tiets
    # For BAN_QUANLUC/BAN_CANBO: auto-approve out-of-scope items
    auto_finalized_ids = []
    for pd in pending_reviews:
        existing_ct_ids = {kq.chi_tiet_id for kq in pd.chi_tiet_duyet}
        for ct in pd.de_xuat.chi_tiets:
            if ct.bi_loai or ct.trang_thai == TrangThaiChiTiet.TU_CHOI.value:
                continue
            if ct.id not in existing_ct_ids:
                in_scope = _is_in_dept_scope(current_user.role, ct.doi_tuong)
                if phong_name != PhongDuyet.PHONG_HAUCANKYTHUAT.value and phong_name != PhongDuyet.BAN_SAUDAIHOC.value:
                    # For Ban Quan luc, only certain doi_tuong are in scope
                   
                    ket_qua_1 = KetQuaDuyetChiTiet(
                        phe_duyet_id=pd.id,
                        chi_tiet_id=ct.id,
                        ket_qua=KetQuaDuyet.CHO_DUYET.value if in_scope else KetQuaDuyet.DONG_Y.value,
                    )
                else:
                    ket_qua_1 = KetQuaDuyetChiTiet(
                        phe_duyet_id=pd.id,
                        chi_tiet_id=ct.id,
                        ket_qua=KetQuaDuyet.DONG_Y.value,
                    )

                # Thêm nhiều bản ghi cùng lúc
                db.session.add_all([ket_qua_1])
    db.session.commit()

    # Auto-finalize departments where ALL items are out-of-scope (all auto-approved)
    for pd in pending_reviews:
        db.session.refresh(pd)
        if pd.ket_qua != KetQuaDuyet.CHO_DUYET.value:
            continue
        active_ct_ids = {ct.id for ct in pd.de_xuat.chi_tiets if not ct.bi_loai}
        pending_in_scope = [
            kq for kq in pd.chi_tiet_duyet
            if kq.chi_tiet_id in active_ct_ids
            and kq.ket_qua == KetQuaDuyet.CHO_DUYET.value
        ]
        if not pending_in_scope and active_ct_ids:
            # All items are auto-approved (out-of-scope) -> auto-finalize
            pd.ket_qua = KetQuaDuyet.DONG_Y.value
            pd.nguoi_duyet_id = None
            pd.ngay_duyet = datetime.utcnow()
            pd.ghi_chu = 'Tự động duyệt (không có đối tượng thuộc phạm vi)'

            de_xuat = pd.de_xuat
            all_dept = PheDuyet.query.filter_by(de_xuat_id=de_xuat.id).filter(
                PheDuyet.phong_duyet != PhongDuyet.ADMIN_TUYENHUAN.value
            ).all()
            if all(a.ket_qua == KetQuaDuyet.DONG_Y.value for a in all_dept):
                de_xuat.trang_thai = TrangThaiDeXuat.HOI_DONG.value
                existing_admin = PheDuyet.query.filter_by(
                    de_xuat_id=de_xuat.id,
                    phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value
                ).first()
                if not existing_admin:
                    admin_pd = PheDuyet(
                        de_xuat_id=de_xuat.id,
                        phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value,
                        ket_qua=KetQuaDuyet.CHO_DUYET.value,
                    )
                    db.session.add(admin_pd)
            auto_finalized_ids.append(pd.id)
    db.session.commit()

    # Remove auto-finalized from pending list
    pending_reviews = [pd for pd in pending_reviews if pd.id not in auto_finalized_ids]

    # Build item results: pd.id -> {ct.id -> KetQuaDuyetChiTiet}
    all_item_results = {}
    for pd in pending_reviews:
        db.session.refresh(pd)
        all_item_results[pd.id] = {kq.chi_tiet_id: kq for kq in pd.chi_tiet_duyet}

    # Build out-of-scope set for template rendering
    out_of_scope_ct_ids = set()
    if current_user.role in (Role.BAN_QUANLUC, Role.BAN_CANBO):
        for pd in pending_reviews:
            for ct in pd.de_xuat.chi_tiets:
                if not _is_in_dept_scope(current_user.role, ct.doi_tuong):
                    out_of_scope_ct_ids.add(ct.id)

    allowed_fields = get_phong_fields().get(current_user.role, [])
    table_columns = get_phong_table_columns().get(current_user.role, [])
    field_conditions = PHONG_FIELD_CONDITIONS.get(current_user.role, {})
    managed_dept_columns = _managed_gate_columns(current_user.role)

    # Roles in _VIEW_ALL_CRITERIA_ROLES (BAN_CANBO/BAN_CTCQ/BAN_BAOVE_ANNINH/BAN_TOCHUC)
    # view ALL criteria like admin. Also fallback: if no mapping configured → show ALL.
    if current_user.role in _VIEW_ALL_CRITERIA_ROLES or not table_columns:
        table_columns = _all_criteria_columns()

    # For Thủ trưởng roles: build ordered list of {dept_name, fields} for gate sub-departments
    # so the template can show criteria columns grouped by department
    gate_dept_fields = []  # [{'dept': dept_name, 'fields': [field, ...]}, ...]
    if current_user.role in _GROUP_CONFIRMATION:
        # Eagerly auto-finalize scope-limited depts (BAN_QUANLUC/BAN_CANBO) for all pending de_xuats
        # This handles nominations submitted before the auto-finalize fix was in place
        _finalized_de_xuat_ids = set()
        for pd in pending_reviews:
            if pd.de_xuat_id not in _finalized_de_xuat_ids:
                _auto_finalize_scope_dept(pd.de_xuat_id)
                _finalized_de_xuat_ids.add(pd.de_xuat_id)

        phong_fields_all = get_phong_fields()
        field_labels_all = get_field_labels()
        for gate_dept_name in managed_dept_columns:
            gate_role = _PHONG_TO_ROLE.get(gate_dept_name)
            fields = []
            if gate_role:
                raw_fields = phong_fields_all.get(gate_role, [])
                # exclude long text/file fields and nckh_minh_chung
                fields = [f for f in raw_fields if f not in _LONG_TEXT_FIELDS]
            if fields:
                gate_dept_fields.append({'dept': gate_dept_name, 'fields': fields})


    # Group gate status for Thủ trưởng roles
    group_gate_by_pd = {}
    group_gate_by_ct = {}
    if current_user.role in _GROUP_CONFIRMATION:
        for pd in pending_reviews:
            group_gate_by_pd[pd.id] = _get_group_gate_for_pd(current_user.role, pd.de_xuat_id)
            ct_map = {}
            for ct in pd.de_xuat.chi_tiets:
                ct_map[ct.id] = _get_group_gate_for_ct(current_user.role, pd.de_xuat_id, ct.id)
            group_gate_by_ct[pd.id] = ct_map

    # Collect unique unit names for dropdown filter
    unit_names = []
    for pd in pending_reviews:
        name = pd.de_xuat.don_vi.ten_don_vi
        if name not in unit_names:
            unit_names.append(name)

    # Compute dynamic tap_the criteria columns from ALL collective DanhHieu definitions
    # so columns are always complete regardless of which items are on screen.
    from app.models.nomination import DanhHieu as _DanhHieu
    tt_all_keys = set()
    for dh in _DanhHieu.query.filter_by(pham_vi='Đơn vị', is_active=True).all():
        for ma_truong in (dh.tieu_chi or []):
            tt_all_keys.add(ma_truong)
    # Also pick up any keys actually stored in current items (legacy / custom data)
    for pd in pending_reviews:
        for ct in pd.de_xuat.chi_tiets:
            if ct.quan_nhan_id is None:
                td = ct.tap_the_dict or {}
                tt_all_keys.update(td.keys())
    if tt_all_keys:
        tt_tieu_chi_rows = TieuChi.query.filter(
            TieuChi.ma_truong.in_(list(tt_all_keys)), TieuChi.is_active == True
        ).order_by(TieuChi.thu_tu, TieuChi.ten).all()
        tt_criteria_fields = [tc.ma_truong for tc in tt_tieu_chi_rows]
        tt_field_labels_map = {tc.ma_truong: tc.ten for tc in tt_tieu_chi_rows}
        for k in tt_all_keys:
            if k not in tt_field_labels_map:
                tt_criteria_fields.append(k)
                tt_field_labels_map[k] = k
    else:
        tt_criteria_fields = []
        tt_field_labels_map = {}

    # Get all active edit requests for this phong
    edit_requests_by_ct = {}
    if pending_reviews:
        all_ct_ids = []
        for pd in pending_reviews:
            if pd.de_xuat:
                for ct in pd.de_xuat.chi_tiets:
                    if not ct.bi_loai:
                        all_ct_ids.append(ct.id)
        
        if all_ct_ids:
            active_edit_requests = YeuCauChinhSua.query.filter(
                YeuCauChinhSua.chi_tiet_id.in_(all_ct_ids),
                YeuCauChinhSua.trang_thai == TrangThaiYeuCauSua.CHO_SUA.value,
                YeuCauChinhSua.phong_yeu_cau == phong_name
            ).all()
            
            for req in active_edit_requests:
                edit_requests_by_ct[req.chi_tiet_id] = req

    return render_template('approval/pending_list.html',
                           pending_reviews=pending_reviews,
                           all_item_results=all_item_results,
                           phong_name=phong_name,
                           allowed_fields=allowed_fields,
                           table_columns=table_columns,
                           field_labels=get_field_labels(),
                           field_conditions=field_conditions,
                           unit_names=unit_names,
                           out_of_scope_ct_ids=out_of_scope_ct_ids,
                           group_gate_by_pd=group_gate_by_pd,
                           group_gate_by_ct=group_gate_by_ct,
                           managed_dept_columns=managed_dept_columns,
                           gate_dept_fields=gate_dept_fields,
                           nam_hoc_filter=nam_hoc_filter,
                           nam_hoc_list=nam_hoc_list,
                           tt_criteria_fields=tt_criteria_fields,
                           tt_field_labels=tt_field_labels_map,
                           edit_requests_by_ct=edit_requests_by_ct)

@approval_bp.route('/review/<int:id>', methods=['GET'])
@login_required
@department_required
def review_nomination(id):
    de_xuat = DeXuat.query.get_or_404(id)
    phong_name = ROLE_TO_PHONG.get(current_user.role, '')

    phe_duyet = PheDuyet.query.filter_by(
        de_xuat_id=id, phong_duyet=phong_name
    ).first_or_404()
   
    group_gate = _get_group_gate_for_pd(current_user.role, id)
    group_gate_by_ct = {}
    if current_user.role in _GROUP_CONFIRMATION:
        for ct in de_xuat.chi_tiets:
            group_gate_by_ct[ct.id] = _get_group_gate_for_ct(current_user.role, id, ct.id)

    # Ensure per-item records exist for all chi_tiets
    # For BAN_QUANLUC/BAN_CANBO: auto-approve out-of-scope items
    existing_ct_ids = {kq.chi_tiet_id for kq in phe_duyet.chi_tiet_duyet}
    for ct in de_xuat.chi_tiets:
        if ct.bi_loai:
            continue
        if ct.doi_tuong is None:
            if phong_name == PhongDuyet.PHONG_HAUCANKYTHUAT.value and phong_name == PhongDuyet.BAN_SAUDAIHOC.value:

                ket_qua_1 = KetQuaDuyetChiTiet(
                            phe_duyet_id=phe_duyet.id,
                            chi_tiet_id=ct.id,
                            ket_qua=KetQuaDuyet.DONG_Y.value,
                        )
        else:
            if ct.id not in existing_ct_ids:
                in_scope = _is_in_dept_scope(current_user.role, ct.doi_tuong)
                if phong_name != PhongDuyet.PHONG_HAUCANKYTHUAT.value and phong_name != PhongDuyet.BAN_SAUDAIHOC.value:

                    ket_qua_1 = KetQuaDuyetChiTiet(
                            phe_duyet_id=phe_duyet.id,
                            chi_tiet_id=ct.id,
                            ket_qua=KetQuaDuyet.CHO_DUYET.value if in_scope else KetQuaDuyet.DONG_Y.value,
                        )
            else:
                if ct.doi_tuong in ['Học viên sau đại học']:
                    ket_qua_1 = KetQuaDuyetChiTiet(
                        phe_duyet_id=phe_duyet.id,
                        chi_tiet_id=ct.id,
                        ket_qua=KetQuaDuyet.CHO_DUYET.value,
                    )
                else:
                    ket_qua_1 = KetQuaDuyetChiTiet(
                        phe_duyet_id=phe_duyet.id,
                        chi_tiet_id=ct.id,
                        ket_qua=KetQuaDuyet.DONG_Y.value,
                    )
        db.session.add(ket_qua_1)                                                                           
    db.session.commit()

    # Reload
    phe_duyet = PheDuyet.query.get(phe_duyet.id)

    # Build lookup: chi_tiet_id -> KetQuaDuyetChiTiet
    item_results = {kq.chi_tiet_id: kq for kq in phe_duyet.chi_tiet_duyet}

    # Build out-of-scope set for template
    out_of_scope_ct_ids = set()
    if current_user.role in (Role.BAN_QUANLUC, Role.BAN_CANBO):
        for ct in de_xuat.chi_tiets:
            if not _is_in_dept_scope(current_user.role, ct.doi_tuong):
                out_of_scope_ct_ids.add(ct.id)

    allowed_fields = get_phong_fields().get(current_user.role, [])
    table_columns = get_phong_table_columns().get(current_user.role, [])
    field_conditions = PHONG_FIELD_CONDITIONS.get(current_user.role, {})

    if current_user.role in _VIEW_ALL_CRITERIA_ROLES or not table_columns:
        table_columns = _all_criteria_columns()

    return render_template('approval/review.html',
                           de_xuat=de_xuat, phe_duyet=phe_duyet,
                           phong_name=phong_name, item_results=item_results,
                           allowed_fields=allowed_fields,
                           table_columns=table_columns,
                           field_labels=get_field_labels(),
                           field_conditions=field_conditions,
                           out_of_scope_ct_ids=out_of_scope_ct_ids,
                           group_gate=group_gate,
                           group_gate_by_ct=group_gate_by_ct)


@approval_bp.route('/review/<int:id>/item/<int:ct_id>/approve', methods=['POST'])
@login_required
@department_required
def approve_item(id, ct_id):
    phong_name = ROLE_TO_PHONG.get(current_user.role, '')
    phe_duyet = PheDuyet.query.filter_by(
        de_xuat_id=id, phong_duyet=phong_name
    ).first_or_404()

    if current_user.role in _GROUP_CONFIRMATION:
        group_gate = _get_group_gate_for_ct(current_user.role, id, ct_id)
        if not group_gate['can_review']:
            flash('Chưa đủ điều kiện phê duyệt của nhóm ban liên quan.', 'warning')
            return redirect(url_for('approval.review_nomination', id=id))

    # Block if out of scope for BAN_QUANLUC/BAN_CANBO
    ct = DeXuatChiTiet.query.get_or_404(ct_id)
    if not _is_in_dept_scope(current_user.role, ct.doi_tuong):
        flash('Cá nhân này không thuộc phạm vi duyệt của bạn.', 'warning')
        return redirect(url_for('approval.review_nomination', id=id))

    kq = KetQuaDuyetChiTiet.query.filter_by(
        phe_duyet_id=phe_duyet.id, chi_tiet_id=ct_id
    ).first_or_404()

    kq.ket_qua = KetQuaDuyet.DONG_Y.value
    kq.ly_do = None
    db.session.commit()

    name = ct.quan_nhan.ho_ten if ct.quan_nhan else 'Đơn vị'
    log_action('dept_approve_item', resource_type='chi_tiet', resource_id=ct_id,
               detail=f'{name} — {phong_name} nhất trí')
    db.session.commit()
    flash(f'Đã nhất trí: {name}', 'success')
    return redirect(url_for('approval.review_nomination', id=id))


@approval_bp.route('/review/<int:id>/item/<int:ct_id>/reject', methods=['POST'])
@login_required
@department_required
def reject_item(id, ct_id):
    phong_name = ROLE_TO_PHONG.get(current_user.role, '')
    phe_duyet = PheDuyet.query.filter_by(
        de_xuat_id=id, phong_duyet=phong_name
    ).first_or_404()

    if current_user.role in _GROUP_CONFIRMATION:
        group_gate = _get_group_gate_for_ct(current_user.role, id, ct_id)
        if not group_gate['can_review']:
            flash('Chưa đủ điều kiện phê duyệt của nhóm ban liên quan.', 'warning')
            return redirect(url_for('approval.review_nomination', id=id))

    # Block if out of scope for BAN_QUANLUC/BAN_CANBO
    ct_obj = DeXuatChiTiet.query.get_or_404(ct_id)
    if not _is_in_dept_scope(current_user.role, ct_obj.doi_tuong):
        flash('Cá nhân này không thuộc phạm vi duyệt của bạn.', 'warning')
        return redirect(url_for('approval.review_nomination', id=id))

    kq = KetQuaDuyetChiTiet.query.filter_by(
        phe_duyet_id=phe_duyet.id, chi_tiet_id=ct_id
    ).first_or_404()

    ly_do = request.form.get('ly_do', '').strip()
    if not ly_do:
        flash('Vui lòng nhập lý do không nhất trí.', 'danger')
        return redirect(url_for('approval.review_nomination', id=id))

    kq.ket_qua = KetQuaDuyet.TU_CHOI.value
    kq.ly_do = ly_do
    _remove_chi_tiet_on_reject(phe_duyet, ct_id, ly_do)
    db.session.commit()

    ct = DeXuatChiTiet.query.get(ct_id)
    name = ct.quan_nhan.ho_ten if ct and ct.quan_nhan else 'Tập thể'
    log_action('dept_reject_item', resource_type='chi_tiet', resource_id=ct_id,
               detail=f'{name} — {phong_name} không nhất trí: {ly_do}')
    db.session.commit()
    flash(f'Đã loại khỏi đề xuất: {name}. Các cá nhân/tập thể còn lại vẫn tiếp tục được xét duyệt. Đã gửi thông báo cho đơn vị.', 'warning')
    return redirect(url_for('approval.review_nomination', id=id))


@approval_bp.route('/review/<int:id>/submit', methods=['POST'])
@login_required
@department_required
def submit_review(id):
    phong_name = ROLE_TO_PHONG.get(current_user.role, '')
    phe_duyet = PheDuyet.query.filter_by(
        de_xuat_id=id, phong_duyet=phong_name
    ).first_or_404()

    if current_user.role in _GROUP_CONFIRMATION:
        blocked = []
        for kq in phe_duyet.chi_tiet_duyet:
            if kq.chi_tiet.bi_loai:
                continue
            ct_gate = _get_group_gate_for_ct(current_user.role, id, kq.chi_tiet_id)
            if not ct_gate['can_review']:
                blocked.append(kq.chi_tiet_id)
        if blocked:
            flash(f'Có {len(blocked)} cá nhân chưa đủ điều kiện theo nhóm ban liên quan.', 'warning')
            return redirect(url_for('approval.review_nomination', id=id))

    # Check all ACTIVE items have been reviewed (removed items are excluded)
    pending_items = [kq for kq in phe_duyet.chi_tiet_duyet
                     if not kq.chi_tiet.bi_loai and kq.ket_qua == KetQuaDuyet.CHO_DUYET.value]
    if pending_items:
        flash(f'Còn {len(pending_items)} cá nhân chưa được duyệt. Vui lòng duyệt tất cả trước khi hoàn tất.', 'danger')
        return redirect(url_for('approval.review_nomination', id=id))

    de_xuat = DeXuat.query.get(id)

    # Rejecting an item already removed it from the process, so finalize as DONG_Y
    # based on the remaining active items.
    phe_duyet.ket_qua = KetQuaDuyet.DONG_Y.value
    phe_duyet.nguoi_duyet_id = current_user.id
    phe_duyet.ngay_duyet = datetime.utcnow()
    phe_duyet.ghi_chu = request.form.get('ghi_chu', '').strip() or None

    _recompute_de_xuat_status(de_xuat)

    db.session.commit()
    log_action('dept_submit_review', resource_type='de_xuat', resource_id=id,
               detail=f'{phong_name} hoàn tất duyệt đề xuất #{id}')
    db.session.commit()
    flash(f'{phong_name} đã hoàn tất duyệt đề xuất.', 'success')
    return redirect(url_for('approval.pending_list'))


@approval_bp.route('/toggle/<int:pd_id>/<int:ct_id>', methods=['POST'])
@login_required
@department_required
def toggle_item(pd_id, ct_id):
    phong_name = ROLE_TO_PHONG.get(current_user.role, '')
    phe_duyet = PheDuyet.query.filter_by(
        id=pd_id, phong_duyet=phong_name
    ).first_or_404()

    if current_user.role in _GROUP_CONFIRMATION:
        group_gate = _get_group_gate_for_ct(current_user.role, phe_duyet.de_xuat_id, ct_id)
        if not group_gate['can_review']:
            return jsonify({'success': False, 'message': 'Chưa đủ điều kiện phê duyệt của nhóm ban liên quan.'}), 403

    # Block if out of scope for BAN_QUANLUC/BAN_CANBO
    ct_obj = DeXuatChiTiet.query.get_or_404(ct_id)
    if not _is_in_dept_scope(current_user.role, ct_obj.doi_tuong):
        return jsonify({'success': False, 'message': 'Cá nhân này không thuộc phạm vi duyệt của bạn.'}), 403

    kq = KetQuaDuyetChiTiet.query.filter_by(
        phe_duyet_id=phe_duyet.id, chi_tiet_id=ct_id
    ).first_or_404()

    data = request.get_json()
    approved = data.get('approved', True)
    ly_do = data.get('ly_do', '').strip()

    if approved:
        kq.ket_qua = KetQuaDuyet.DONG_Y.value
        kq.ly_do = None
        db.session.commit()
    else:
        if not ly_do:
            return jsonify({'success': False, 'message': 'Vui lòng nhập lý do'}), 400
        kq.ket_qua = KetQuaDuyet.TU_CHOI.value
        kq.ly_do = ly_do
        # Reject = remove ONLY this cá nhân/tập thể; the rest of the đề xuất continues.
        _remove_chi_tiet_on_reject(phe_duyet, ct_id, ly_do)
        db.session.commit()

    de_xuat = phe_duyet.de_xuat

    # Auto-finalize this department once none of its ACTIVE (non-removed) in-scope
    # items are still pending. Removed items never block or reject.
    active_ct_ids = {ct.id for ct in de_xuat.chi_tiets if not ct.bi_loai}
    pending_count = KetQuaDuyetChiTiet.query.filter_by(
        phe_duyet_id=phe_duyet.id,
        ket_qua=KetQuaDuyet.CHO_DUYET.value
    ).join(DeXuatChiTiet, KetQuaDuyetChiTiet.chi_tiet_id == DeXuatChiTiet.id).filter(
        DeXuatChiTiet.bi_loai == False
    ).count()

    auto_finalized = False
    if pending_count == 0 and active_ct_ids:
        if phe_duyet.ket_qua != KetQuaDuyet.DONG_Y.value:
            phe_duyet.ket_qua = KetQuaDuyet.DONG_Y.value
            phe_duyet.nguoi_duyet_id = current_user.id
            phe_duyet.ngay_duyet = datetime.utcnow()
        # Advance the whole đề xuất if every department has now approved.
        _recompute_de_xuat_status(de_xuat)
        db.session.commit()
        auto_finalized = True

    # Build stats over ACTIVE items only
    all_kq = [k for k in phe_duyet.chi_tiet_duyet if not k.chi_tiet.bi_loai]
    total = len(all_kq)
    approved_count = sum(1 for k in all_kq if k.ket_qua == KetQuaDuyet.DONG_Y.value)
    rejected_count = sum(1 for k in all_kq if k.ket_qua == KetQuaDuyet.TU_CHOI.value)

    return jsonify({
        'success': True,
        'ket_qua': kq.ket_qua,
        'auto_finalized': auto_finalized,
        'stats': {
            'total': total,
            'reviewed': approved_count + rejected_count,
            'approved': approved_count,
            'rejected': rejected_count,
        }
    })


def _reviewable_fields_for_role(role, ct):
    """Return the set of ma_truong this department may flag for editing on the
    given chi_tiet (cá nhân or tập thể)."""
    if ct.quan_nhan_id is None:
        # ★ FIX: Tập thể — trả về TẤT CẢ field tiêu chí tập thể theo config,
        # KHÔNG chỉ những field đã có giá trị trong tap_the_dict.
        # Điều này cho phép yêu cầu điền mới field đang rỗng.
        
        # ★ Tập thể: lấy từ config động + union với dict hiện có
        config_fields   = set(_all_tap_the_columns())
        existing_fields = set((ct.tap_the_dict or {}).keys())
        return config_fields | existing_fields

    # Cá nhân — giữ nguyên logic cũ
    if role in _VIEW_ALL_CRITERIA_ROLES:
        fields = set(_all_criteria_columns())
    else:
        fields = set(get_phong_table_columns().get(role, []))
        if not fields:
            fields = set(_all_criteria_columns())
    return fields
# ── Cache module-level để tránh query DB nhiều lần ──────────────────────────

# Python 3.9 trở xuống — dùng Optional từ typing
from typing import Optional

_criteria_cache: Optional[dict] = None

def _get_criteria_by_type() -> dict:
    """Query bảng TieuChi một lần, phân loại theo nhóm:
    - nhom bắt đầu bằng 'ban_' hoặc 'phong_' → tập thể
    - còn lại → cá nhân
    Trả về dict: {'ca_nhan': [...], 'tap_the': [...]}
    """
    global _criteria_cache
    if _criteria_cache is not None:
        return _criteria_cache

    from app.models.nomination import TieuChi

    all_tc = TieuChi.query.order_by(TieuChi.thu_tu.asc()).all()

    ca_nhan = []
    tap_the = []
    for tc in all_tc:
        if not tc.ma_truong:
            continue
        nhom = (tc.nhom or '').strip().lower()
        if nhom.startswith('ban_') or nhom.startswith('phong_'):
            tap_the.append(tc.ma_truong)
        else:
            ca_nhan.append(tc.ma_truong)

    _criteria_cache = {'ca_nhan': ca_nhan, 'tap_the': tap_the}
    return _criteria_cache

def _all_tap_the_columns():
    """Trả về list ma_truong tiêu chí tập thể — lấy động từ TieuChi."""
    fields = _get_criteria_by_type()['tap_the']
    return fields


@approval_bp.route('/request-edit/<int:pd_id>/<int:ct_id>', methods=['POST'])
@login_required
@department_required
def request_edit(pd_id, ct_id):
    """Approver flags one or more criteria of a single cá nhân/tập thể and asks the
    unit to fix them. Only the flagged criteria become editable by the unit; all other
    data stays locked. The flagging department's result for this item is reset to
    CHO_DUYET so it must re-review after the unit resubmits."""
    phong_name = ROLE_TO_PHONG.get(current_user.role, '')
    phe_duyet = PheDuyet.query.filter_by(
        id=pd_id, phong_duyet=phong_name
    ).first_or_404()

    ct = DeXuatChiTiet.query.get_or_404(ct_id)
    if ct.de_xuat_id != phe_duyet.de_xuat_id or ct.bi_loai:
        return jsonify({'success': False, 'message': 'Cá nhân/tập thể không hợp lệ.'}), 400

    if not _is_in_dept_scope(current_user.role, ct.doi_tuong):
        return jsonify({'success': False, 'message': 'Cá nhân này không thuộc phạm vi duyệt của bạn.'}), 403

    data = request.get_json(silent=True) or {}
    fields = data.get('fields') or []
    ly_do = (data.get('ly_do') or '').strip()
    if not isinstance(fields, list) or not fields:
        return jsonify({'success': False, 'message': 'Vui lòng chọn ít nhất một tiêu chí cần chỉnh sửa.'}), 400

    allowed = _reviewable_fields_for_role(current_user.role, ct)
    fields = [f for f in fields if f in allowed]
    if not fields and not phong_name == PhongDuyet.BAN_TUYENHUAN.value:
        return jsonify({'success': False, 'message': 'Các tiêu chí được chọn không thuộc phạm vi duyệt của bạn.'}), 400

    # Reuse an existing open request for the same item from the same department.
    yc = YeuCauChinhSua.query.filter_by(
        chi_tiet_id=ct_id,
        phong_yeu_cau=phong_name,
        trang_thai=TrangThaiYeuCauSua.CHO_SUA.value,
    ).first()
    if yc:
        merged = list(dict.fromkeys((yc.cac_truong or []) + fields))
        yc.cac_truong = merged
        yc.ly_do = ly_do or yc.ly_do
        yc.nguoi_yeu_cau_id = current_user.id
    else:
        yc = YeuCauChinhSua(
            de_xuat_id=phe_duyet.de_xuat_id,
            chi_tiet_id=ct_id,
            phong_yeu_cau=phong_name,
            nguoi_yeu_cau_id=current_user.id,
            ly_do=ly_do,
            trang_thai=TrangThaiYeuCauSua.CHO_SUA.value,
        )
        yc.cac_truong = fields
        db.session.add(yc)

    # Reset this department's result for the item so it must re-review after edit.
    kq = KetQuaDuyetChiTiet.query.filter_by(
        phe_duyet_id=phe_duyet.id, chi_tiet_id=ct_id
    ).first()
    if kq:
        kq.ket_qua = KetQuaDuyet.CHO_DUYET.value
        kq.ly_do = None
    # Keep this department open (not finalized) while the edit is pending.
    if phe_duyet.ket_qua == KetQuaDuyet.DONG_Y.value:
        phe_duyet.ket_qua = KetQuaDuyet.CHO_DUYET.value
        phe_duyet.ngay_duyet = None

    # If de_xuat had already advanced to HOI_DONG, revert it back to DANG_DUYET
    # so it clearly shows as "in departmental review" again, and the admin/hội đồng
    # PheDuyet (which was created prematurely) is removed.
    de_xuat = phe_duyet.de_xuat
    if de_xuat.trang_thai == TrangThaiDeXuat.HOI_DONG.value:
        de_xuat.trang_thai = TrangThaiDeXuat.DANG_DUYET.value
        admin_pd = PheDuyet.query.filter_by(
            de_xuat_id=de_xuat.id,
            phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value,
            ket_qua=KetQuaDuyet.CHO_DUYET.value,
        ).first()
        if admin_pd:
            db.session.delete(admin_pd)
    unit_user = User.query.filter_by(
        don_vi_id=de_xuat.don_vi_id, role=Role.UNIT_USER
    ).first()
    if unit_user:
        name = (ct.quan_nhan.ho_ten if ct.quan_nhan else
                (ct.ten_don_vi_de_xuat or de_xuat.don_vi.ten_don_vi))
        labels = get_field_labels()
        field_names = ', '.join(labels.get(f, f) for f in fields)
        db.session.add(ThongBao(
            user_id=unit_user.id,
            de_xuat_id=de_xuat.id,
            chi_tiet_id=ct_id,
            loai='yeu_cau_sua',
            tieu_de=f'{phong_name} yêu cầu chỉnh sửa: {name}',
            noi_dung=(f'Tiêu chí cần chỉnh sửa: {field_names}. '
                      f'Lý do: {ly_do or "Không rõ"}. '
                      f'Đề xuất năm học {de_xuat.nam_hoc} của {de_xuat.don_vi.ten_don_vi}.'),
        ))

    db.session.commit()
    return jsonify({'success': True, 'message': f'Đã gửi yêu cầu chỉnh sửa cho đơn vị ({len(fields)} tiêu chí).'})


@approval_bp.route('/revoke-item/<int:pd_id>/<int:ct_id>', methods=['POST'])
@login_required
@department_required
def revoke_item(pd_id, ct_id):
    """Thu hồi kết quả duyệt của 1 cá nhân/tập thể, đưa về CHO_DUYET."""
    phong_name = ROLE_TO_PHONG.get(current_user.role, '')
    phe_duyet = PheDuyet.query.filter_by(id=pd_id, phong_duyet=phong_name).first_or_404()
    de_xuat = phe_duyet.de_xuat

    if de_xuat.trang_thai in (TrangThaiDeXuat.PHE_DUYET_CUOI.value, TrangThaiDeXuat.HOI_DONG.value):
        return jsonify({'success': False, 'message': 'Không thể thu hồi - đề xuất đã qua giai đoạn duyệt của bộ phận.'}), 403

    kq = KetQuaDuyetChiTiet.query.filter_by(
        phe_duyet_id=phe_duyet.id, chi_tiet_id=ct_id
    ).first_or_404()

    if kq.ket_qua == KetQuaDuyet.CHO_DUYET.value:
        return jsonify({'success': False, 'message': 'Mục này chưa được duyệt.'}), 400

    # Reset this item
    kq.ket_qua = KetQuaDuyet.CHO_DUYET.value
    kq.ly_do = None

    # If PheDuyet was finalized, revert it back to pending
    if phe_duyet.ket_qua != KetQuaDuyet.CHO_DUYET.value:
        phe_duyet.ket_qua = KetQuaDuyet.CHO_DUYET.value
        phe_duyet.nguoi_duyet_id = None
        phe_duyet.ngay_duyet = None
        phe_duyet.ly_do = None
        # Revert DeXuat status to DANG_DUYET if it was HOI_DONG
        if de_xuat.trang_thai == TrangThaiDeXuat.HOI_DONG.value:
            de_xuat.trang_thai = TrangThaiDeXuat.DANG_DUYET.value

    db.session.commit()
    return jsonify({'success': True, 'message': 'Đã thu hồi kết quả duyệt cho cá nhân này.'})


@approval_bp.route('/batch-approve', methods=['POST'])
@login_required
@department_required
def batch_approve():
    """Batch approve multiple items at once (by unit or all)."""
    phong_name = ROLE_TO_PHONG.get(current_user.role, '')
    data = request.get_json()
    pd_id = data.get('pd_id')
    ct_ids = data.get('ct_ids', [])

    if not pd_id or not ct_ids:
        return jsonify({'success': False, 'message': 'Thiếu dữ liệu'}), 400

    phe_duyet = PheDuyet.query.filter_by(
        id=pd_id, phong_duyet=phong_name
    ).first_or_404()

    if current_user.role in _GROUP_CONFIRMATION:
        blocked_ids = []
        for ct_id in ct_ids:
            ct_gate = _get_group_gate_for_ct(current_user.role, phe_duyet.de_xuat_id, ct_id)
            if not ct_gate['can_review']:
                blocked_ids.append(ct_id)
        if blocked_ids:
            return jsonify({'success': False, 'message': f'Có {len(blocked_ids)} cá nhân chưa đủ điều kiện phê duyệt của nhóm ban liên quan.'}), 403

    if phe_duyet.ket_qua != KetQuaDuyet.CHO_DUYET.value:
        return jsonify({'success': False, 'message': 'Đã hoàn tất duyệt'}), 400

    approved_names = []
    for ct_id in ct_ids:
        # Skip out-of-scope items for BAN_QUANLUC/BAN_CANBO
        ct = DeXuatChiTiet.query.get(ct_id)
        if ct and not _is_in_dept_scope(current_user.role, ct.doi_tuong):
            continue
        kq = KetQuaDuyetChiTiet.query.filter_by(
            phe_duyet_id=phe_duyet.id, chi_tiet_id=ct_id
        ).first()
        if kq and kq.ket_qua == KetQuaDuyet.CHO_DUYET.value:
            kq.ket_qua = KetQuaDuyet.DONG_Y.value
            kq.ly_do = None
            name = ct.quan_nhan.ho_ten if ct and ct.quan_nhan else 'Đơn vị'
            approved_names.append(name)

    db.session.commit()

    de_xuat = phe_duyet.de_xuat

    # Auto-finalize once none of the ACTIVE (non-removed) items remain pending.
    active_ct_ids = {ct.id for ct in de_xuat.chi_tiets if not ct.bi_loai}
    pending_count = KetQuaDuyetChiTiet.query.filter_by(
        phe_duyet_id=phe_duyet.id,
        ket_qua=KetQuaDuyet.CHO_DUYET.value
    ).join(DeXuatChiTiet, KetQuaDuyetChiTiet.chi_tiet_id == DeXuatChiTiet.id).filter(
        DeXuatChiTiet.bi_loai == False
    ).count()

    auto_finalized = False
    if pending_count == 0 and active_ct_ids:
        if phe_duyet.ket_qua != KetQuaDuyet.DONG_Y.value:
            phe_duyet.ket_qua = KetQuaDuyet.DONG_Y.value
            phe_duyet.nguoi_duyet_id = current_user.id
            phe_duyet.ngay_duyet = datetime.utcnow()
        _recompute_de_xuat_status(de_xuat)
        db.session.commit()
        auto_finalized = True

    # Build stats over ACTIVE items only
    all_kq = [k for k in phe_duyet.chi_tiet_duyet if not k.chi_tiet.bi_loai]
    total = len(all_kq)
    approved_count = sum(1 for k in all_kq if k.ket_qua == KetQuaDuyet.DONG_Y.value)
    rejected_count = sum(1 for k in all_kq if k.ket_qua == KetQuaDuyet.TU_CHOI.value)

    return jsonify({
        'success': True,
        'approved_count': len(approved_names),
        'auto_finalized': auto_finalized,
        'stats': {
            'total': total,
            'reviewed': approved_count + rejected_count,
            'approved': approved_count,
            'rejected': rejected_count,
        }
    })


@approval_bp.route('/history')
@login_required
@department_required
def history():
    phong_name = ROLE_TO_PHONG.get(current_user.role, '')

    page = request.args.get('page', 1, type=int)
    unit_filter = request.args.get('unit', '')
    danh_hieu_filter = request.args.get('danh_hieu', '')
    ket_qua_filter = request.args.get('ket_qua', '')

    # Query per-individual results (KetQuaDuyetChiTiet) for this department
    query = db.session.query(KetQuaDuyetChiTiet).join(
        PheDuyet, KetQuaDuyetChiTiet.phe_duyet_id == PheDuyet.id
    ).filter(
        PheDuyet.phong_duyet == phong_name,
        KetQuaDuyetChiTiet.ket_qua != KetQuaDuyet.CHO_DUYET.value,
    )

    # Apply filters
    if ket_qua_filter:
        query = query.filter(KetQuaDuyetChiTiet.ket_qua == ket_qua_filter)

    if unit_filter:
        query = query.join(DeXuat, PheDuyet.de_xuat_id == DeXuat.id).join(
            DonVi, DeXuat.don_vi_id == DonVi.id
        ).filter(DonVi.ten_don_vi == unit_filter)
    else:
        query = query.join(DeXuat, PheDuyet.de_xuat_id == DeXuat.id)

    if danh_hieu_filter:
        query = query.join(
            DeXuatChiTiet, KetQuaDuyetChiTiet.chi_tiet_id == DeXuatChiTiet.id
        ).filter(DeXuatChiTiet.loai_danh_hieu == danh_hieu_filter)

    individual_results = query.order_by(
        PheDuyet.ngay_duyet.desc(), KetQuaDuyetChiTiet.id.desc()
    ).paginate(page=page, per_page=20, error_out=False)

    # Get filter options
    unit_names_q = db.session.query(DonVi.ten_don_vi).join(
        DeXuat, DeXuat.don_vi_id == DonVi.id
    ).join(PheDuyet, PheDuyet.de_xuat_id == DeXuat.id).join(
        KetQuaDuyetChiTiet, KetQuaDuyetChiTiet.phe_duyet_id == PheDuyet.id
    ).filter(
        PheDuyet.phong_duyet == phong_name,
        KetQuaDuyetChiTiet.ket_qua != KetQuaDuyet.CHO_DUYET.value,
    ).distinct().order_by(DonVi.ten_don_vi).all()
    unit_names = [u[0] for u in unit_names_q]

    # Summary stats (unfiltered)
    base_q = db.session.query(KetQuaDuyetChiTiet).join(
        PheDuyet, KetQuaDuyetChiTiet.phe_duyet_id == PheDuyet.id
    ).filter(
        PheDuyet.phong_duyet == phong_name,
        KetQuaDuyetChiTiet.ket_qua != KetQuaDuyet.CHO_DUYET.value,
    )
    stats = {
        'total': base_q.count(),
        'approved': base_q.filter(KetQuaDuyetChiTiet.ket_qua == KetQuaDuyet.DONG_Y.value).count(),
        'rejected': base_q.filter(KetQuaDuyetChiTiet.ket_qua == KetQuaDuyet.TU_CHOI.value).count(),
    }

    allowed_fields = get_phong_fields().get(current_user.role, [])
    table_columns = get_phong_table_columns().get(current_user.role, [])

    if current_user.role in _VIEW_ALL_CRITERIA_ROLES or not table_columns:
        table_columns = _all_criteria_columns()

    # Build tt_criteria_fields from ALL active collective DanhHieu definitions,
    # not from the current page items (which is paginated and would miss criteria
    # from items on other pages or in items not yet loaded).
    from app.models.nomination import DanhHieu
    tt_all_keys = set()
    collective_danh_hieus = DanhHieu.query.filter_by(pham_vi='Đơn vị', is_active=True).all()
    for dh in collective_danh_hieus:
        for ma_truong in (dh.tieu_chi or []):
            tt_all_keys.add(ma_truong)

    # Also collect any keys actually present in current-page items that may not
    # be in DanhHieu definitions (legacy data).
    for kq_item in individual_results.items:
        ct = kq_item.chi_tiet
        if ct and ct.quan_nhan_id is None:
            td = ct.tap_the_dict or {}
            tt_all_keys.update(td.keys())

    if tt_all_keys:
        tt_tieu_chi_rows = TieuChi.query.filter(
            TieuChi.ma_truong.in_(list(tt_all_keys)), TieuChi.is_active == True
        ).order_by(TieuChi.thu_tu, TieuChi.ten).all()
        tt_history_fields = [tc.ma_truong for tc in tt_tieu_chi_rows]
        tt_field_labels_h = {tc.ma_truong: tc.ten for tc in tt_tieu_chi_rows}
        for k in tt_all_keys:
            if k not in tt_field_labels_h:
                tt_history_fields.append(k)
                tt_field_labels_h[k] = k
    else:
        tt_history_fields = []
        tt_field_labels_h = {}

    return render_template('approval/history.html',
                           individual_results=individual_results,
                           phong_name=phong_name,
                           unit_filter=unit_filter,
                           danh_hieu_filter=danh_hieu_filter,
                           ket_qua_filter=ket_qua_filter,
                           unit_names=unit_names,
                           stats=stats,
                           allowed_fields=allowed_fields,
                           table_columns=table_columns,
                           field_labels=get_field_labels(),
                           tt_history_fields=tt_history_fields,
                           tt_field_labels=tt_field_labels_h)


@approval_bp.route('/history/chi-tiet/<int:ct_id>')
@login_required
@department_required
def history_detail(ct_id):
    """View detailed info for one individual in history, scoped to current department."""
    phong_name = ROLE_TO_PHONG.get(current_user.role, '')
    ct = DeXuatChiTiet.query.get_or_404(ct_id)
    de_xuat = ct.de_xuat

    # Get this department's PheDuyet and individual result
    phe_duyet = PheDuyet.query.filter_by(
        de_xuat_id=de_xuat.id, phong_duyet=phong_name
    ).first_or_404()

    kq = KetQuaDuyetChiTiet.query.filter_by(
        phe_duyet_id=phe_duyet.id, chi_tiet_id=ct.id
    ).first()

    allowed_fields = get_phong_fields().get(current_user.role, [])

    return render_template('approval/history_detail.html',
                           ct=ct,
                           de_xuat=de_xuat,
                           phe_duyet=phe_duyet,
                           kq=kq,
                           phong_name=phong_name,
                           allowed_fields=allowed_fields,
                           field_labels=get_field_labels())


@approval_bp.route('/revoke/<int:pd_id>', methods=['POST'])
@login_required
@department_required
def revoke_review(pd_id):
    """Revoke (thu hồi) a completed department review, resetting it to pending."""
    phong_name = ROLE_TO_PHONG.get(current_user.role, '')

    phe_duyet = PheDuyet.query.filter_by(
        id=pd_id, phong_duyet=phong_name
    ).first_or_404()

    de_xuat = phe_duyet.de_xuat

    # Cannot revoke after admin final approval
    if de_xuat.trang_thai == TrangThaiDeXuat.PHE_DUYET_CUOI.value:
        flash('Không thể thu hồi - đề xuất đã được phê duyệt cuối cùng.', 'danger')
        return redirect(url_for('approval.history'))

    # Can only revoke a completed review (not one that's still pending)
    if phe_duyet.ket_qua == KetQuaDuyet.CHO_DUYET.value:
        flash('Kết quả duyệt này vẫn đang chờ, không cần thu hồi.', 'warning')
        return redirect(url_for('approval.history'))

    # For paired scope logic, disallow revoke if this department is auto-approved-by-scope
    if current_user.role in (Role.BAN_CANBO, Role.BAN_QUANLUC):
        in_scope_items = [
            kq for kq in phe_duyet.chi_tiet_duyet
            if _is_in_dept_scope(current_user.role, kq.chi_tiet.doi_tuong)
        ]
        if not in_scope_items:
            flash('Kết quả tự động theo phạm vi, không thể thu hồi.', 'warning')
            return redirect(url_for('approval.history'))

    # 1. Reset the PheDuyet record
    phe_duyet.ket_qua = KetQuaDuyet.CHO_DUYET.value
    phe_duyet.nguoi_duyet_id = None
    phe_duyet.ngay_duyet = None
    phe_duyet.ly_do = None
    phe_duyet.ghi_chu = None

    # 2. Reset KetQuaDuyetChiTiet records back to pending
    #    For BAN_QUANLUC/BAN_CANBO: only reset in-scope items, keep out-of-scope as DONG_Y
    for kq in phe_duyet.chi_tiet_duyet:
        ct = kq.chi_tiet
        if ct and not _is_in_dept_scope(current_user.role, ct.doi_tuong):
            # Out-of-scope: keep auto-approved
            continue
        kq.ket_qua = KetQuaDuyet.CHO_DUYET.value
        kq.ly_do = None

    # 3. Handle DeXuat status changes
    old_status = de_xuat.trang_thai

    # If admin Tuyên huấn PheDuyet was already created (status was 'Đã duyệt'),
    # delete it since not all 6 depts approve anymore
    if old_status == TrangThaiDeXuat.HOI_DONG.value:
        admin_pd = PheDuyet.query.filter_by(
            de_xuat_id=de_xuat.id,
            phong_duyet=PhongDuyet.ADMIN_TUYENHUAN.value
        ).first()
        if admin_pd:
            # Delete admin's chi_tiet_duyet records if any
            KetQuaDuyetChiTiet.query.filter_by(phe_duyet_id=admin_pd.id).delete()
            db.session.delete(admin_pd)

    # Determine new status: check other departments
    other_depts = PheDuyet.query.filter_by(de_xuat_id=de_xuat.id).filter(
        PheDuyet.phong_duyet != PhongDuyet.ADMIN_TUYENHUAN.value,
        PheDuyet.id != phe_duyet.id  # exclude the one we just revoked
    ).all()

    has_other_completed = any(
        pd.ket_qua != KetQuaDuyet.CHO_DUYET.value for pd in other_depts
    )

    if has_other_completed:
        # At least one other dept has completed their review
        de_xuat.trang_thai = TrangThaiDeXuat.DANG_DUYET.value
    else:
        # No dept has completed review - back to waiting
        de_xuat.trang_thai = TrangThaiDeXuat.CHO_DUYET.value

    db.session.commit()
    flash(f'{phong_name} đã thu hồi kết quả duyệt cho đề xuất của {de_xuat.don_vi.ten_don_vi}.', 'success')
    return redirect(url_for('approval.history'))


@approval_bp.route('/export-excel')
@login_required
@department_required
def export_excel():
    """Export pending review list to Excel with timestamp."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    phong_name = ROLE_TO_PHONG.get(current_user.role, '')
    nam_hoc_filter = request.args.get('nam_hoc', '')

    q = PheDuyet.query.filter_by(phong_duyet=phong_name, ket_qua=KetQuaDuyet.CHO_DUYET.value)
    if nam_hoc_filter:
        from app.models.nomination import DeXuat as _DeXuat
        q = q.join(_DeXuat, PheDuyet.de_xuat_id == _DeXuat.id).filter(_DeXuat.nam_hoc == nam_hoc_filter)
    pending_reviews = q.order_by(PheDuyet.created_at.desc()).all()

    # Build out-of-scope set
    out_of_scope_ct_ids = set()
    if current_user.role in (Role.BAN_QUANLUC, Role.BAN_CANBO):
        for pd in pending_reviews:
            for ct in pd.de_xuat.chi_tiets:
                if not _is_in_dept_scope(current_user.role, ct.doi_tuong):
                    out_of_scope_ct_ids.add(ct.id)

    # Build item results
    all_item_results = {}
    for pd in pending_reviews:
        all_item_results[pd.id] = {kq.chi_tiet_id: kq for kq in pd.chi_tiet_duyet}

    # Get criteria fields for current department
    phong_fields_map = get_phong_fields()
    field_labels = get_field_labels()
    criteria_fields = phong_fields_map.get(current_user.role, [])

    wb = Workbook()
    ws = wb.active
    ws.title = 'Phê duyệt khen thưởng'

    # Timestamp header
    ts = datetime.now().strftime('%d/%m/%Y %H:%M')
    total_cols = 8 + len(criteria_fields)
    last_col = get_column_letter(total_cols)
    ws.merge_cells(f'A1:{last_col}1')
    ws['A1'] = f'DANH SÁCH PHÊ DUYỆT KHEN THƯỞNG - {phong_name}'
    ws['A1'].font = Font(bold=True, size=13)
    ws['A1'].alignment = Alignment(horizontal='center')

    ws.merge_cells(f'A2:{last_col}2')
    ws['A2'] = f'Năm học: {nam_hoc_filter or "Tất cả"} | Xuất lúc: {ts}'
    ws['A2'].alignment = Alignment(horizontal='center')
    ws['A2'].font = Font(italic=True, size=10)

    # Header row
    base_headers = ['STT', 'Đơn vị', 'Họ tên', 'Cấp bậc', 'Chức vụ', 'Đối tượng', 'Danh hiệu', 'Kết quả']
    criteria_headers = [field_labels.get(f, f) for f in criteria_fields]
    headers = base_headers + criteria_headers

    header_fill = PatternFill('solid', fgColor='1B3A6B')
    header_font = Font(bold=True, color='FFFFFF', size=10)
    thin = Side(style='thin', color='CCCCCC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border

    row_num = 5
    stt = 0
    for pd in pending_reviews:
        results = all_item_results.get(pd.id, {})
        don_vi = pd.de_xuat.don_vi.ten_don_vi
        for ct in pd.de_xuat.chi_tiets:
            if ct.id in out_of_scope_ct_ids:
                continue
            stt += 1
            kq = results.get(ct.id)
            ket_qua_str = ''
            if kq:
                if kq.ket_qua == 'Đồng ý':
                    ket_qua_str = 'Nhất trí'
                elif kq.ket_qua == 'Từ chối':
                    ket_qua_str = f'Không NT: {kq.ly_do or ""}'
                else:
                    ket_qua_str = 'Chờ duyệt'
            ho_ten = ct.quan_nhan.ho_ten if ct.quan_nhan else don_vi
            cap_bac = ct.quan_nhan.cap_bac if ct.quan_nhan else ''
            chuc_vu = ct.quan_nhan.chuc_vu if ct.quan_nhan else ''

            # Build criteria values (individual: getattr; tap_the: tap_the_dict)
            is_tap_the = ct.quan_nhan_id is None
            if is_tap_the:
                tt_dict = ct.tap_the_dict or {}
                criteria_vals = [tt_dict.get(f, '') for f in criteria_fields]
            else:
                criteria_vals = [getattr(ct, f, '') or '' for f in criteria_fields]

            row_data = [stt, don_vi, ho_ten, cap_bac or '', chuc_vu or '',
                        ct.doi_tuong or '', ct.loai_danh_hieu or '', ket_qua_str] + criteria_vals
            for col, val in enumerate(row_data, 1):
                cell = ws.cell(row=row_num, column=col, value=val)
                cell.border = border
                cell.alignment = Alignment(vertical='center', wrap_text=True)
                if ket_qua_str == 'Nhất trí':
                    cell.fill = PatternFill('solid', fgColor='D4EDDA')
                elif ket_qua_str.startswith('Không NT'):
                    cell.fill = PatternFill('solid', fgColor='F8D7DA')
            row_num += 1

    # Column widths
    base_widths = [6, 30, 25, 16, 20, 18, 22, 28]
    criteria_widths = [18] * len(criteria_fields)
    col_widths = base_widths + criteria_widths
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[4].height = 28

    # Page setup: A4 landscape, fit to 1 page wide
    ws.page_setup.paperSize = 9
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Sheet protection
    ws.protection.sheet = True
    ws.protection.password = 'hktd@2025'

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    ts_file = datetime.now().strftime('%Y%m%d_%H%M')
    filename = f'phe_duyet_{phong_name.replace(" ", "_")}_{ts_file}.xlsx'
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=filename)
@approval_bp.route('/export-word')
@login_required
@department_required
def export_word():
    """Export danh sách phê duyệt khen thưởng ra Word — 4 mục theo danh hiệu."""
    from io import BytesIO
    from datetime import date, datetime
    from docx import Document
    from docx.shared import Cm, Pt, RGBColor
    from docx.enum.text import WD_LINE_SPACING, WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    db.session.expire_all()

    phong_name     = ROLE_TO_PHONG.get(current_user.role, '')
    nam_hoc_filter = request.args.get('nam_hoc', '')

    # ── Query ─────────────────────────────────────────────────────────────────
    q = PheDuyet.query.filter_by(phong_duyet=phong_name, ket_qua=KetQuaDuyet.CHO_DUYET.value)
    if nam_hoc_filter:
        q = q.join(DeXuat, PheDuyet.de_xuat_id == DeXuat.id).filter(DeXuat.nam_hoc == nam_hoc_filter)
    pending_reviews = q.order_by(PheDuyet.created_at.desc()).all()

    # Out-of-scope
    out_of_scope_ct_ids = set()
    if current_user.role in (Role.BAN_QUANLUC, Role.BAN_CANBO):
        for pd in pending_reviews:
            for ct in pd.de_xuat.chi_tiets:
                if not _is_in_dept_scope(current_user.role, ct.doi_tuong):
                    out_of_scope_ct_ids.add(ct.id)

    all_item_results = {
        pd.id: {kq.chi_tiet_id: kq for kq in pd.chi_tiet_duyet}
        for pd in pending_reviews
    }

    phong_fields_map = get_phong_fields()
    field_labels     = get_field_labels()
    criteria_fields  = phong_fields_map.get(current_user.role, [])

    # ── Phân loại chi tiết theo danh hiệu ────────────────────────────────────
    # Mỗi item: dict {ct, pd, don_vi, kq, ket_qua_str, row_shading, kq_color}
    ds_don_vi_qt  = []   # Đơn vị quyết thắng
    ds_don_vi_tt  = []   # Đơn vị tiên tiến
    ds_ca_nhan_td = []   # Chiến sĩ thi đua
    ds_ca_nhan_tt = []   # Chiến sĩ tiên tiến
    seen_ids      = set()

    for pd in pending_reviews:
        results = all_item_results.get(pd.id, {})
        don_vi  = pd.de_xuat.don_vi.ten_don_vi if pd.de_xuat.don_vi else ''
        for ct in pd.de_xuat.chi_tiets:
            if ct.id in out_of_scope_ct_ids or ct.id in seen_ids:
                continue
            seen_ids.add(ct.id)

            kq = results.get(ct.id)
            if kq:
                if kq.ket_qua == KetQuaDuyet.DONG_Y.value:
                    ket_qua_str = 'Nhất trí'
                    row_shading = 'D4EDDA'
                    kq_color    = (0x15, 0x57, 0x24)
                elif kq.ket_qua == KetQuaDuyet.TU_CHOI.value:
                    ket_qua_str = f'Không NT: {kq.ly_do or ""}'
                    row_shading = 'F8D7DA'
                    kq_color    = (0x72, 0x1C, 0x24)
                else:
                    ket_qua_str = 'Chờ duyệt'
                    row_shading = None
                    kq_color    = None
            else:
                ket_qua_str = 'Chờ duyệt'
                row_shading = None
                kq_color    = None

            item = dict(ct=ct, pd=pd, don_vi=don_vi,
                        kq=kq, ket_qua_str=ket_qua_str,
                        row_shading=row_shading, kq_color=kq_color)

            dh = (ct.loai_danh_hieu or '').strip()
            if dh == 'Đơn vị quyết thắng':
                ds_don_vi_qt.append(item)
            elif dh == 'Đơn vị tiên tiến':
                ds_don_vi_tt.append(item)
            elif dh == 'Chiến sĩ thi đua':
                ds_ca_nhan_td.append(item)
            elif dh == 'Chiến sĩ tiên tiến':
                ds_ca_nhan_tt.append(item)
            # Bỏ qua danh hiệu khác (hoặc thêm ds_khac nếu cần)

    # ═══════════════════════════════════════════════════════════════════════════
    # HELPERS
    # ═══════════════════════════════════════════════════════════════════════════
    def set_font(run, bold=False, size=11, italic=False, color=None):
        run.bold       = bold
        run.italic     = italic
        run.font.size  = Pt(size)
        run.font.name  = 'Times New Roman'
        if color:
            run.font.color.rgb = RGBColor(*color)

    def para_font(para, text, bold=False, size=11,
                  align=WD_ALIGN_PARAGRAPH.LEFT, italic=False):
        para.alignment = align
        run = para.add_run(text)
        set_font(run, bold=bold, size=size, italic=italic)
        return run

    def add_cell(cell, text, bold=False, size=10,
                 align=WD_ALIGN_PARAGRAPH.LEFT, color=None):
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = align
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after  = Pt(1)
        run = p.add_run(str(text) if text is not None else '')
        set_font(run, bold=bold, size=size, color=color)
        return run

    def set_cell_shading(cell, hex_color):
        tcPr = cell._tc.get_or_add_tcPr()
        shd  = OxmlElement('w:shd')
        shd.set(qn('w:val'),   'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'),  hex_color)
        tcPr.append(shd)

    def set_repeat_table_header(row):
        tr   = row._tr
        trPr = tr.get_or_add_trPr()
        if trPr.find(qn('w:tblHeader')) is None:
            trPr.append(OxmlElement('w:tblHeader'))

    def _lock_cell_width(cell, width_cm):
        cell.width = Cm(width_cm)
        tcPr = cell._tc.get_or_add_tcPr()
        tcW  = tcPr.find(qn('w:tcW'))
        if tcW is None:
            tcW = OxmlElement('w:tcW'); tcPr.append(tcW)
        tcW.set(qn('w:w'), str(int(Cm(width_cm) / 635)))
        tcW.set(qn('w:type'), 'dxa')

    def _remove_cell_borders(cell):
        tcPr      = cell._tc.get_or_add_tcPr()
        tcBorders = OxmlElement('w:tcBorders')
        for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            tag = OxmlElement(f'w:{edge}')
            tag.set(qn('w:val'),   'none')
            tag.set(qn('w:sz'),    '0')
            tag.set(qn('w:space'), '0')
            tag.set(qn('w:color'), 'auto')
            tcBorders.append(tag)
        tcPr.append(tcBorders)

    def _clear_cell_margin(cell):
        tcPr  = cell._tc.get_or_add_tcPr()
        tcMar = tcPr.find(qn('w:tcMar'))
        if tcMar is None:
            tcMar = OxmlElement('w:tcMar'); tcPr.append(tcMar)
        for edge in ('top', 'left', 'bottom', 'right'):
            tag = tcMar.find(qn(f'w:{edge}'))
            if tag is None:
                tag = OxmlElement(f'w:{edge}'); tcMar.append(tag)
            tag.set(qn('w:w'), '0'); tag.set(qn('w:type'), 'dxa')

    def _set_cell_pad_left(cell, cm):
        tcPr  = cell._tc.get_or_add_tcPr()
        tcMar = tcPr.find(qn('w:tcMar'))
        if tcMar is None:
            tcMar = OxmlElement('w:tcMar'); tcPr.append(tcMar)
        tag = tcMar.find(qn('w:left'))
        if tag is None:
            tag = OxmlElement('w:left'); tcMar.append(tag)
        tag.set(qn('w:w'), str(int(Cm(cm) / 635)))
        tag.set(qn('w:type'), 'dxa')

    def _set_cell_pad_right(cell, cm):
        tcPr  = cell._tc.get_or_add_tcPr()
        tcMar = tcPr.find(qn('w:tcMar'))
        if tcMar is None:
            tcMar = OxmlElement('w:tcMar'); tcPr.append(tcMar)
        tag = tcMar.find(qn('w:right'))
        if tag is None:
            tag = OxmlElement('w:right'); tcMar.append(tag)
        tag.set(qn('w:w'), str(int(Cm(cm) / 635)))
        tag.set(qn('w:type'), 'dxa')

    def _para(cell, is_first=False, align=WD_ALIGN_PARAGRAPH.CENTER):
        p = cell.paragraphs[0] if is_first else cell.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before      = Pt(0)
        p.paragraph_format.space_after       = Pt(0)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        p.paragraph_format.left_indent       = Pt(0)
        p.paragraph_format.right_indent      = Pt(0)
        return p

    def build_tom_tat(ct):
        """Tóm tắt thành tích — trả về list, mỗi tiêu chí 1 dòng."""
        parts = []
        if ct.muc_do_hoan_thanh:
            parts.append(ct.muc_do_hoan_thanh)
        if ct.diem_tong_ket:
            parts.append(f'Kết quả học tập: {ct.diem_tong_ket}')
        if ct.ket_qua_ren_luyen:
            parts.append(f'Rèn luyện: {ct.ket_qua_ren_luyen}')
        if ct.hinh_thuc_tot_nghiep:
            tn = [f'TN: {ct.hinh_thuc_tot_nghiep}']
            for attr, label in (
                ('diem_tn_ctd',        'CTĐ-CT'),
                ('diem_tn_ct',         'CT'),
                ('diem_tn_ta',         'TA'),
                ('diem_tn_mon4',       'Môn 4'),
                ('diem_tn_chuyennganh','Chuyên ngành'),
                ('diem_tn_baove',      'Bảo vệ'),
            ):
                val = getattr(ct, attr, None)
                if val:
                    tn.append(f'{label}: {val}')
            parts.append(', '.join(tn))
        if ct.mo_ta_khoa_hoc:
            parts.append(f'NCKH: {ct.mo_ta_khoa_hoc}')
        if ct.thanh_tich_ca_nhan_khac:
            parts.append(ct.thanh_tich_ca_nhan_khac)
        # Tiêu chí động của ban
        for f in criteria_fields:
            val = getattr(ct, f, None) or ''
            if val:
                label = field_labels.get(f, f)
                parts.append(f'{label}: {val}')
        return parts

    # ── Hàm thêm bảng cá nhân ────────────────────────────────────────────────
    def add_personnel_table(doc, items, section_label, stt_start=1):
        """Bảng 7 cột: STT | Họ tên | Cấp bậc | Chức vụ | Đơn vị | Tóm tắt | Ghi chú"""
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after  = Pt(2)
        para_font(p, section_label, bold=True, size=11)

        if not items:
            p2 = doc.add_paragraph()
            para_font(p2, '(Không có)', size=10, italic=True)
            return stt_start

        # Widths: STT | Họ tên | Cấp bậc | Chức vụ | Đơn vị | Tóm tắt | Ghi chú
        widths = [0.7, 3.5, 1.8, 2.2, 2.5, 5.0, 1.5]
        
        tbl = doc.add_table(rows=1, cols=len(widths))
        set_fixed_table_widths(tbl, widths)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.style     = 'Table Grid'
        tbl.autofit   = True

        tblPr = tbl._tbl.tblPr
        tblLayout = tblPr.find(qn('w:tblLayout'))
        if tblLayout is None:
            tblLayout = OxmlElement('w:tblLayout'); tblPr.append(tblLayout)
        tblLayout.set(qn('w:type'), 'autofit')

        # tblGrid
        tblGrid = tbl._tbl.find(qn('w:tblGrid'))
        if tblGrid is not None:
            tbl._tbl.remove(tblGrid)
        tblGrid = OxmlElement('w:tblGrid')
        tbl._tbl.insert(1, tblGrid)
        for w in widths:
            gc = OxmlElement('w:gridCol')
            gc.set(qn('w:w'), str(int(Cm(w) / 635)))
            tblGrid.append(gc)

        # Chiều rộng từng cột
        for i, w in enumerate(widths):
            twips = int(Cm(w) / 635)
            tc    = tbl.rows[0].cells[i]._tc
            tcPr  = tc.get_or_add_tcPr()
            tcW   = tcPr.find(qn('w:tcW'))
            if tcW is None:
                tcW = OxmlElement('w:tcW'); tcPr.append(tcW)
            tcW.set(qn('w:w'), str(twips)); tcW.set(qn('w:type'), 'dxa')

        # Header row
        headers_txt = ['STT', 'Họ và tên', 'Cấp bậc', 'Chức vụ',
                       'Đơn vị', 'Tóm tắt thành tích', 'Ghi chú']
        
        hrow = tbl.rows[0]

        for i, h in enumerate(headers_txt):
            run = add_cell(hrow.cells[i], h, bold=True, size=10,
                           align=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_shading(hrow.cells[i], '1B3A6B')
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        set_repeat_table_header(hrow)

        # Data rows
        stt = stt_start
        for item in items:
            ct          = item['ct']
            don_vi      = item['don_vi']
            ket_qua_str = item['ket_qua_str']
            row_shading = item['row_shading']
            kq_color    = item['kq_color']
            if ct.bi_loai == True or ct.trang_thai == TrangThaiChiTiet.TU_CHOI:
                continue
            qn_obj  = ct.quan_nhan
            ho_ten  = qn_obj.ho_ten  if qn_obj else (ct.ten_don_vi_de_xuat or don_vi)
            cap_bac = qn_obj.cap_bac if qn_obj and qn_obj.cap_bac else ''
            chuc_vu = qn_obj.chuc_vu if qn_obj and qn_obj.chuc_vu else ''

            row = tbl.add_row()
            add_cell(row.cells[0], str(stt), align=WD_ALIGN_PARAGRAPH.CENTER)
            add_cell(row.cells[1], ho_ten)
            add_cell(row.cells[2], cap_bac)
            add_cell(row.cells[3], chuc_vu)
            add_cell(row.cells[4], don_vi)

            # Cột Tóm tắt — mỗi tiêu chí 1 dòng
            tom_tat_list = build_tom_tat(ct)
            cell_tt = row.cells[5]
            cell_tt.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p_tt = cell_tt.paragraphs[0]
            p_tt.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p_tt.paragraph_format.space_before = Pt(1)
            p_tt.paragraph_format.space_after  = Pt(1)
            if tom_tat_list:
                for idx_i, item_txt in enumerate(tom_tat_list):
                    if idx_i > 0:
                        p_tt = cell_tt.add_paragraph()
                        p_tt.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        p_tt.paragraph_format.space_before = Pt(1)
                        p_tt.paragraph_format.space_after  = Pt(1)
                    set_font(p_tt.add_run(f'- {item_txt}'), size=10)
            else:
                set_font(p_tt.add_run('-'), size=10)

            # Cột Ghi chú — hiện kết quả duyệt
            run_kq = add_cell(row.cells[6], ket_qua_str, size=9,
                              align=WD_ALIGN_PARAGRAPH.CENTER)
            if kq_color:
                run_kq.font.color.rgb = RGBColor(*kq_color)

            if row_shading:
                for c in row.cells:
                    set_cell_shading(c, row_shading)

            stt += 1

        # Dòng tổng
        sum_row = tbl.add_row()
        merged  = sum_row.cells[0]
        for i in range(1, len(widths)):
            merged = merged.merge(sum_row.cells[i])
        add_cell(merged, f'Tổng cộng: {stt - stt_start} người', bold=True, size=10)
        set_cell_shading(merged, 'EEF2FF')

        return stt  # trả về stt tiếp theo (nếu cần đánh số liên tục)

    # ── Hàm thêm bảng đơn vị ─────────────────────────────────────────────────
    def add_unit_table(doc, items, section_label):
        """Bảng 3 cột: STT | Tên đơn vị | Ghi chú"""
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after  = Pt(2)
        para_font(p, section_label, bold=True, size=11)

        if not items:
            p2 = doc.add_paragraph()
            para_font(p2, '(Không có)', size=10, italic=True)
            return

        widths = [0.7, 2.5,3, 10.0]  # STT | Tên đơn vị | Ghi chú
        
        tbl = doc.add_table(rows=1, cols=len(widths))
        set_fixed_table_widths(tbl, widths)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.style     = 'Table Grid'
        tbl.autofit   = True

        tblPr = tbl._tbl.tblPr
        tblLayout = tblPr.find(qn('w:tblLayout'))
        if tblLayout is None:
            tblLayout = OxmlElement('w:tblLayout'); tblPr.append(tblLayout)
        tblLayout.set(qn('w:type'), 'autofit')

        tblGrid = tbl._tbl.find(qn('w:tblGrid'))
        if tblGrid is not None:
            tbl._tbl.remove(tblGrid)
        tblGrid = OxmlElement('w:tblGrid')
        tbl._tbl.insert(1, tblGrid)
        for w in widths:
            gc = OxmlElement('w:gridCol')
            gc.set(qn('w:w'), str(int(Cm(w) / 635)))
            tblGrid.append(gc)

        for i, w in enumerate(widths):
            twips = int(Cm(w) / 635)
            tc    = tbl.rows[0].cells[i]._tc
            tcPr  = tc.get_or_add_tcPr()
            tcW   = tcPr.find(qn('w:tcW'))
            if tcW is None:
                tcW = OxmlElement('w:tcW'); tcPr.append(tcW)
            tcW.set(qn('w:w'), str(twips)); tcW.set(qn('w:type'), 'dxa')

        # Header row
        hrow = tbl.rows[0]
        for i, h in enumerate(['STT', 'Tên đơn vị','Đề xuất của đơn vị', 'Ghi chú']):
            run = add_cell(hrow.cells[i], h, bold=True, size=10,
                           align=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_shading(hrow.cells[i], '1B3A6B')
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        set_repeat_table_header(hrow)

        # Data rows
        for idx, item in enumerate(items, 1):
            ct          = item['ct']
            ket_qua_str = item['ket_qua_str']
            row_shading = item['row_shading']
            kq_color    = item['kq_color']
            if ct.bi_loai == True or ct.trang_thai == TrangThaiChiTiet.TU_CHOI:
                continue
            ten_dv = (ct.ten_don_vi_de_xuat or item['don_vi'] or '-')

            row = tbl.add_row()
            add_cell(row.cells[0], str(idx), align=WD_ALIGN_PARAGRAPH.CENTER)
            add_cell(row.cells[1], ten_dv)
           
            

            cell = row.cells[2]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after  = Pt(1)

            criteria_list = []
            td = ct.tap_the_dict or {}
            if td:
                from app.models.nomination import TieuChi as _TieuChi
                ma_truong_list = list(td.keys())
                tieu_chi_map = {}
                if ma_truong_list:
                    tc_rows = _TieuChi.query.filter(
                        _TieuChi.ma_truong.in_(ma_truong_list)
                    ).all()
                    tieu_chi_map = {tc.ma_truong: tc.ten for tc in tc_rows}
                for key, val in td.items():
                    if val and str(val).strip() not in ('', '0', 'None'):
                        label = tieu_chi_map.get(key, key)
                        criteria_list.append(f'{label}: {val}')

            if ct.muc_do_hoan_thanh:
                criteria_list.insert(0, f'Mức độ hoàn thành: {ct.muc_do_hoan_thanh}')
            if ct.ghi_chu and ct.ghi_chu.strip():
                criteria_list.append(f'Ghi chú: {ct.ghi_chu}')

            if criteria_list:
                for idx_c, item in enumerate(criteria_list):
                    if idx_c > 0:
                        p = cell.add_paragraph()
                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        p.paragraph_format.space_before = Pt(1)
                        p.paragraph_format.space_after  = Pt(1)
                    run = p.add_run(f'- {item}')
                    set_font(run, size=10)
            else:
                run = p.add_run('-')
                set_font(run, size=10)
                

            if row_shading:
                for c in row.cells:
                    set_cell_shading(c, row_shading)

        # Dòng tổng
        sum_row = tbl.add_row()
        merged  = sum_row.cells[0]
        for i in range(1, len(widths)):
            merged = merged.merge(sum_row.cells[i])
        add_cell(merged, f'Tổng cộng: {len(items)} đơn vị', bold=True, size=10)
        set_cell_shading(merged, 'EEF2FF')

    # ═══════════════════════════════════════════════════════════════════════════
    # DOCUMENT SETUP
    # ═══════════════════════════════════════════════════════════════════════════
    doc = Document()

    PAGE_W_CM     = 21.0
    MARGIN_L      = 3.5
    MARGIN_R      = 1.5
    PRINT_W_CM    = PAGE_W_CM - MARGIN_L - MARGIN_R   # 16.0
    HEADER_W_CM   = PAGE_W_CM                          # 21.0
    OVERFLOW_EACH = (HEADER_W_CM - PRINT_W_CM) / 2    # 2.5
    LEFT_CM       = 9.5
    RIGHT_CM      = 11.5

    for section in doc.sections:
        section.top_margin    = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin   = Cm(MARGIN_L)
        section.right_margin  = Cm(MARGIN_R)

    # ── Header Quốc hiệu ─────────────────────────────────────────────────────
    tbl_header = doc.add_table(rows=1, cols=2)
    tbl_header.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_header.autofit   = False

    tblPr = tbl_header._tbl.tblPr
    tblW  = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = OxmlElement('w:tblW'); tblPr.append(tblW)
    tblW.set(qn('w:w'), str(int(Cm(HEADER_W_CM) / 635)))
    tblW.set(qn('w:type'), 'dxa')

    tblLayout = tblPr.find(qn('w:tblLayout'))
    if tblLayout is None:
        tblLayout = OxmlElement('w:tblLayout'); tblPr.append(tblLayout)
    tblLayout.set(qn('w:type'), 'fixed')

    tblInd = tblPr.find(qn('w:tblInd'))
    if tblInd is not None:
        tblPr.remove(tblInd)

    tblGrid = tbl_header._tbl.find(qn('w:tblGrid'))
    if tblGrid is not None:
        tbl_header._tbl.remove(tblGrid)
    tblGrid = OxmlElement('w:tblGrid')
    tbl_header._tbl.insert(1, tblGrid)
    for w in [LEFT_CM, RIGHT_CM]:
        gridCol = OxmlElement('w:gridCol')
        gridCol.set(qn('w:w'), str(int(Cm(w) / 635)))
        tblGrid.append(gridCol)

    lc = tbl_header.rows[0].cells[0]
    rc = tbl_header.rows[0].cells[1]
    for cell, w in ((lc, LEFT_CM), (rc, RIGHT_CM)):
        _lock_cell_width(cell, w)
        _remove_cell_borders(cell)
        _clear_cell_margin(cell)
    _set_cell_pad_left(lc,  OVERFLOW_EACH)
    _set_cell_pad_right(rc, OVERFLOW_EACH)

    p_l1 = _para(lc, is_first=True)
    set_font(p_l1.add_run('TRƯỜNG SĨ QUAN CHÍNH TRỊ'), size=12)
    p_l2 = _para(lc)
    r_l2 = p_l2.add_run(phong_name.upper())
    set_font(r_l2, bold=True, size=12); r_l2.underline = True
    _para(lc).add_run('')

    p_r1 = _para(rc, is_first=True)
    set_font(p_r1.add_run('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM'), bold=True, size=12)
    p_r2 = _para(rc)
    r_r2 = p_r2.add_run('Độc lập - Tự do - Hạnh phúc')
    set_font(r_r2, bold=True, size=12); r_r2.underline = True
    p_r3 = _para(rc)
    p_r3.paragraph_format.space_before = Pt(3)
    today = date.today()
    set_font(
        p_r3.add_run(f'Hà Nội, ngày {today.day} tháng {today.month} năm {today.year}'),
        size=11, italic=True
    )
    for row in tbl_header.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                p.paragraph_format.space_before      = Pt(0)
                p.paragraph_format.space_after       = Pt(0)
                p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

    # ── Tiêu đề ───────────────────────────────────────────────────────────────
    doc.add_paragraph()
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(
        p_title.add_run(
            f'DANH SÁCH PHÊ DUYỆT KHEN THƯỞNG'
            f'{(" NĂM HỌC " + nam_hoc_filter) if nam_hoc_filter else ""}'
        ),
        bold=True, size=13
    )
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(
        p_sub.add_run(
            f'({phong_name} — Xuất lúc {datetime.now().strftime("%H:%M ngày %d/%m/%Y")})'
        ),
        size=10, italic=True
    )
    doc.add_paragraph()

    # ═══════════════════════════════════════════════════════════════════════════
    # 4 MỤC THEO DANH HIỆU
    # ═══════════════════════════════════════════════════════════════════════════
    add_unit_table(doc, ds_don_vi_qt,
                   'I. DANH HIỆU ĐƠN VỊ QUYẾT THẮNG')

    add_unit_table(doc, ds_don_vi_tt,
                   'II. DANH HIỆU ĐƠN VỊ TIÊN TIẾN')

    add_personnel_table(doc, ds_ca_nhan_td,
                        'III. DANH HIỆU CHIẾN SĨ THI ĐUA')

    add_personnel_table(doc, ds_ca_nhan_tt,
                        'IV. DANH HIỆU CHIẾN SĨ TIÊN TIẾN')

    # ── Footer ────────────────────────────────────────────────────────────────
    p_foot = doc.add_paragraph()
    p_foot.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_foot = p_foot.add_run(
        f'(Xuất lúc {datetime.now().strftime("%H:%M ngày %d/%m/%Y")})'
    )
    r_foot.font.size      = Pt(9)
    r_foot.font.italic    = True
    r_foot.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    # --- Ký tên ---
    doc.add_paragraph()
    tbl_sign = doc.add_table(rows=1, cols=2)
    tbl_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
    for cell in tbl_sign.rows[0].cells:
        for edge in ('top','left','bottom','right'):
            tc = cell._tc; tcPr = tc.get_or_add_tcPr()
            b = OxmlElement('w:tcBorders')
            tag = OxmlElement(f'w:{edge}')
            tag.set(qn('w:val'), 'none')
            b.append(tag)
            tcPr.append(b)

    left_sign = tbl_sign.rows[0].cells[0]
    right_sign = tbl_sign.rows[0].cells[1]

    # Ký xác nhận bên trái
    p_sl = left_sign.paragraphs[0]
    p_sl.alignment = WD_ALIGN_PARAGRAPH.CENTER
   # para_font(p_sl, 'XÁC NHẬN CỦA CẤP TRÊN', bold=True, size=11)
    left_sign.add_paragraph()
    left_sign.add_paragraph()
    left_sign.add_paragraph()

    # Ký đơn vị bên phải
    # Ký đơn vị bên phải
    p_sr = right_sign.paragraphs[0]
    # Sửa ở đây: Truyền thẳng tham số align vào para_font
    para_font(p_sr, 'THỦ TRƯỞNG ĐƠN VỊ', bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    
    p_sr2 = right_sign.add_paragraph()
    # Sửa ở đây: Truyền thẳng tham số align vào para_font
    para_font(p_sr2, '(Ký, ghi rõ họ tên)', size=10, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    try:
        protect_document_formatting_only(doc, 'bth123')
    except Exception:
        pass
    add_corner_logo(doc)
    # ── Xuất file ─────────────────────────────────────────────────────────────
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)

    ts_file  = datetime.now().strftime('%Y%m%d_%H%M')
    filename = f'phe_duyet_{phong_name.replace(" ", "_")}_{ts_file}.docx'
    return send_file(
        buf, as_attachment=True, download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )
def set_fixed_table_widths(tbl, widths_cm):
        """Can thiệp sâu vào XML để khóa chết chiều rộng bảng, Word không thể tự đổi"""
        # 1. Ép kiểu bảng thành Fixed Layout (Không tự co giãn)
        tbl.autofit = False
        tblPr = tbl._tbl.tblPr
        tblLayout = tblPr.find(qn('w:tblLayout'))
        if tblLayout is None:
            tblLayout = OxmlElement('w:tblLayout')
            tblPr.append(tblLayout)
        tblLayout.set(qn('w:type'), 'fixed')

        # 2. Xóa lưới cột cũ và xây lại khung lưới mới theo đúng kích thước cm
        tblGrid = tbl._tbl.find(qn('w:tblGrid'))
        if tblGrid is not None:
            tbl._tbl.remove(tblGrid)
        tblGrid = OxmlElement('w:tblGrid')
        tbl._tbl.insert(1, tblGrid)  # Chèn khung lưới vào đúng vị trí chuẩn XML

        for w in widths_cm:
            gridCol = OxmlElement('w:gridCol')
            # Chuyển đổi Cm sang đơn vị Twips của Word (1 twip = 635 EMUs)
            gridCol.set(qn('w:w'), str(int(Cm(w) / 635)))
            tblGrid.append(gridCol)

        # 3. Khóa cứng chiều rộng ở cấp độ từng Ô (Cell)
        for i, w in enumerate(widths_cm):
            twips_val = str(int(Cm(w) / 635))
            for row in tbl.rows:
                tcPr = row.cells[i]._tc.get_or_add_tcPr()
                tcW = tcPr.find(qn('w:tcW'))
                if tcW is None:
                    tcW = OxmlElement('w:tcW')
                    tcPr.append(tcW)
                tcW.set(qn('w:w'), twips_val)
                tcW.set(qn('w:type'), 'dxa')
def add_corner_logo(doc):
    """Thêm logo nhỏ ở góc phải trên cùng của trang (sau header table hiện tại)."""
    import os
    from flask import current_app
    
    logo_path = os.path.join(current_app.root_path, 'static', 'img', 'watermark.png')
    
    if not os.path.exists(logo_path):
        # Fallback to main logo if watermark doesn't exist
        logo_path = os.path.join(current_app.root_path, 'static', 'img', 'logo-Si-quan.png')
        if not os.path.exists(logo_path):
            return
    
    try:
        for section in doc.sections:
            header = section.header
            
            # Thêm paragraph mới vào cuối header (sau table header hiện tại)
            para = header.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            
            # Set paragraph spacing để logo sát lề trên
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(0)
            
            # Thêm logo nhỏ căn phải (1.5cm)
            run = para.add_run()
            run.add_picture(logo_path, width=Cm(1.5))
            
    except Exception as e:
        print(f"Warning: Could not add corner logo: {e}")


def protect_document_formatting_only(doc, password: str):
    """
    Khóa tài liệu: chỉ đọc nội dung (readOnly).
    Mật khẩu được hash theo chuẩn Office 2010+ (Agile Encryption).
    """
    # 1. Tạo salt ngẫu nhiên (16 bytes)
    salt = os.urandom(16)
    salt_b64 = binascii.b2a_base64(salt).strip().decode()

    # 2. Hash lần đầu: SHA-512(salt + password)
    # Lưu ý: password bắt buộc encode sang chuẩn UTF-16 Little Endian
    key = hashlib.sha512(salt + password.encode('utf-16le')).digest()
    
    # 3. Lặp 100.000 vòng để chống brute-force
    spin_count = 100000
    for i in range(spin_count):
        iterator = i.to_bytes(4, byteorder='little')
        # SỬA LỖI: Cần cộng iterator ở PHÍA SAU hash của vòng lặp liền trước
        key = hashlib.sha512(key + iterator).digest()
        
    hash_b64 = binascii.b2a_base64(key).strip().decode()

    # 4. Lấy cấu hình settings của docx
    settings = doc.settings.element

    # Xóa thẻ documentProtection cũ nếu có
    for old in settings.findall(qn('w:documentProtection')):
        settings.remove(old)

    # 5. Tạo thẻ <w:documentProtection> theo chuẩn Office đời mới
    doc_prot = OxmlElement('w:documentProtection')
    doc_prot.set(qn('w:edit'),          'readOnly')
    doc_prot.set(qn('w:enforcement'),   '1')
    
    # BỎ CÁC THẺ CŨ (cryptProviderType, v.v.). SỬ DỤNG CHUẨN AGILE MỚI:
    doc_prot.set(qn('w:algorithmName'), 'SHA-512')
    doc_prot.set(qn('w:spinCount'),     str(spin_count))
    doc_prot.set(qn('w:hashValue'),     hash_b64)     # Đã đổi từ w:hash thành w:hashValue
    doc_prot.set(qn('w:saltValue'),     salt_b64)     # Đã đổi từ w:salt thành w:saltValue

    # Chèn vào đầu <w:settings>
    settings.insert(0, doc_prot)
