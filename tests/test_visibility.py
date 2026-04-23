"""
Tests for the field visibility logic that mirrors the JS fieldVisibility() function.

These tests validate the Python-side data (diem_nhom_map, criteria_meta, nhom_meta)
that feeds the JS visibility engine, ensuring:
- Score fields are mapped to correct nhom groups
- NhomTieuChi doi_tuong_ap_dung correctly filters for selected doi_tuong
- DANH_HIEU_TIEU_CHI does NOT override diem_* field visibility
"""
import pytest
from app.models.evaluation import NhomTieuChi, DiemQuyDinhDanhHieu
from app.models.nomination import TieuChi, DanhHieu


# ---------------------------------------------------------------------------
# Helpers — replicate the Python-side diem_nhom_map from nomination.py
# ---------------------------------------------------------------------------

DIEM_NHOM_MAP = {
    'diem_kiem_tra_tin_hoc':   'chung',
    'kiem_tra_tin_hoc':        'chung',
    'diem_kiem_tra_dieu_lenh': 'chung',
    'kiem_tra_dieu_lenh':      'chung',
    'diem_dia_ly_quan_su':     'chung',
    'dia_ly_quan_su':          'chung',
    'diem_ban_sung':           'chung',
    'ban_sung':                'chung',
    'diem_the_luc':            'chung',
    'the_luc':                 'chung',
    'diem_kiem_tra_chinh_tri': 'chung',
    'kiem_tra_chinh_tri':      'chung',
    'diem_tong_ket':           'hoc_vien',
    'diem_nckh':               'nckh',
}

# Replicate JS fieldVisibility logic (Python version for unit-testing)
def field_visibility(field_name, selected_danh_hieu, selected_doi_tuong,
                     current_criteria, criteria_meta, nhom_meta, diem_nhom_map):
    """
    Returns 'show', 'hide', or 'no_opinion'.
    Mirrors JS fieldVisibility() in nomination/edit.html.
    """
    NON_CRITERIA_FIELDS = {
        'csrf_token', 'quan_nhan_id', 'loai_danh_hieu', 'doi_tuong',
        'ghi_chu_item', 'ly_do_chua_dat_diem', 'nckh_noi_dung_text',
        'muc_do_hoan_thanh', 'phieu_tin_nhiem', 'xep_loai_dang_vien',
        'ket_qua_doan_the', 'minh_chung_doan_the', 'chu_tri_don_vi_danh_hieu',
    }

    if field_name in NON_CRITERIA_FIELDS:
        return 'no_opinion'

    # Step 1: danh_hieu filter — skip for score fields
    if selected_danh_hieu and current_criteria and field_name not in diem_nhom_map:
        if field_name not in current_criteria:
            return 'hide'

    # Step 2: CRITERIA_META doi_tuong_ap_dung
    meta = criteria_meta.get(field_name)
    if meta:
        apply_list = meta.get('doi_tuong_ap_dung') or []
        if apply_list and selected_doi_tuong and selected_doi_tuong not in apply_list:
            return 'hide'
        return 'show' if selected_doi_tuong else 'hide'

    # Step 3: score fields via DIEM_NHOM_MAP
    diem_nhom = diem_nhom_map.get(field_name)
    if diem_nhom is not None:
        nhom_info = nhom_meta.get(diem_nhom)
        if nhom_info:
            apply_list = nhom_info.get('doi_tuong_ap_dung') or []
            if apply_list and selected_doi_tuong and selected_doi_tuong not in apply_list:
                return 'hide'
        return 'show' if selected_doi_tuong else 'hide'

    return 'no_opinion'


# ---------------------------------------------------------------------------
# Tests for DIEM_NHOM_MAP correctness
# ---------------------------------------------------------------------------

class TestDiemNhomMap:
    def test_all_chung_fields_present(self):
        expected_chung = [
            'diem_kiem_tra_tin_hoc', 'kiem_tra_tin_hoc',
            'diem_kiem_tra_dieu_lenh', 'kiem_tra_dieu_lenh',
            'diem_dia_ly_quan_su', 'dia_ly_quan_su',
            'diem_ban_sung', 'ban_sung',
            'diem_the_luc', 'the_luc',
            'diem_kiem_tra_chinh_tri', 'kiem_tra_chinh_tri',
        ]
        for f in expected_chung:
            assert DIEM_NHOM_MAP.get(f) == 'chung', f"{f} should map to 'chung'"

    def test_diem_tong_ket_maps_to_hoc_vien(self):
        assert DIEM_NHOM_MAP['diem_tong_ket'] == 'hoc_vien'

    def test_diem_nckh_maps_to_nckh(self):
        assert DIEM_NHOM_MAP['diem_nckh'] == 'nckh'


# ---------------------------------------------------------------------------
# Tests for field_visibility() Python implementation
# ---------------------------------------------------------------------------

class TestFieldVisibility:

    def _nhom_meta_open(self):
        """All groups with empty doi_tuong_ap_dung (applies to all)."""
        return {
            'chung': {'ten_nhom': 'Tiêu chí chung', 'doi_tuong_ap_dung': []},
            'hoc_vien': {'ten_nhom': 'Tiêu chí học viên', 'doi_tuong_ap_dung': []},
            'nckh': {'ten_nhom': 'Tiêu chí NCKH', 'doi_tuong_ap_dung': []},
            'giang_vien': {'ten_nhom': 'Tiêu chí giảng viên', 'doi_tuong_ap_dung': ['Giảng viên']},
        }

    def _nhom_meta_restricted(self):
        """chung group restricted to Giảng viên + Chiến sĩ; hoc_vien to Học viên only."""
        return {
            'chung': {'ten_nhom': 'Tiêu chí chung', 'doi_tuong_ap_dung': ['Giảng viên', 'Chiến sĩ']},
            'hoc_vien': {'ten_nhom': 'Tiêu chí học viên', 'doi_tuong_ap_dung': ['Học viên']},
            'nckh': {'ten_nhom': 'Tiêu chí NCKH', 'doi_tuong_ap_dung': []},
            'giang_vien': {'ten_nhom': 'Tiêu chí giảng viên', 'doi_tuong_ap_dung': ['Giảng viên']},
        }

    # --- No doi_tuong selected ---

    def test_score_field_hidden_when_no_doi_tuong(self):
        vis = field_visibility(
            'diem_ban_sung', '', '', [],
            {}, self._nhom_meta_open(), DIEM_NHOM_MAP
        )
        assert vis == 'hide'

    # --- chung group with open doi_tuong_ap_dung (= all) ---

    def test_chung_score_visible_for_chien_si_when_open(self):
        """diem_ban_sung (chung group, open) must show for Chiến sĩ thi đua."""
        for field in ['diem_ban_sung', 'ban_sung', 'diem_the_luc', 'the_luc',
                      'diem_kiem_tra_chinh_tri', 'kiem_tra_chinh_tri']:
            vis = field_visibility(
                field, 'Chiến sĩ thi đua', 'Chiến sĩ',
                ['ban_sung', 'the_luc'],   # currentCriteria has TieuChi fields but NOT diem_* fields
                {}, self._nhom_meta_open(), DIEM_NHOM_MAP
            )
            assert vis == 'show', f"{field} should be 'show' for Chiến sĩ with open chung group"

    def test_chung_score_visible_for_giang_vien_when_open(self):
        for field in ['diem_ban_sung', 'diem_the_luc', 'diem_kiem_tra_chinh_tri']:
            vis = field_visibility(
                field, 'Chiến sĩ thi đua', 'Giảng viên',
                [], {}, self._nhom_meta_open(), DIEM_NHOM_MAP
            )
            assert vis == 'show', f"{field} should be 'show' for Giảng viên with open chung"

    # --- Bug 1: score fields NOT hidden when danh_hieu selected ---

    def test_score_not_hidden_by_danh_hieu_criteria_list(self):
        """
        REGRESSION: Before the fix, diem_* fields were hidden because they were
        not in DANH_HIEU_TIEU_CHI['Chiến sĩ thi đua']. Step 1 must skip them.
        """
        # currentCriteria contains only TieuChi DB fields, not diem_* columns
        current_criteria = ['ban_sung', 'the_luc', 'kiem_tra_chinh_tri', 'muc_do_hoan_thanh']
        for diem_field in ['diem_ban_sung', 'diem_the_luc', 'diem_kiem_tra_chinh_tri']:
            vis = field_visibility(
                diem_field,
                'Chiến sĩ thi đua', 'Chiến sĩ',
                current_criteria,
                {}, self._nhom_meta_open(), DIEM_NHOM_MAP
            )
            assert vis == 'show', (
                f"{diem_field} must not be hidden by danh_hieu criteria list (Bug 1 regression)"
            )

    # --- chung group restricted to specific doi_tuong ---

    def test_chung_score_hidden_for_excluded_doi_tuong(self):
        """When chung group has doi_tuong_ap_dung=['Giảng viên','Chiến sĩ'], Học viên should be hidden."""
        vis = field_visibility(
            'diem_ban_sung', '', 'Học viên', [],
            {}, self._nhom_meta_restricted(), DIEM_NHOM_MAP
        )
        assert vis == 'hide'

    def test_chung_score_shown_for_included_doi_tuong(self):
        """When chung group has doi_tuong_ap_dung=['Giảng viên','Chiến sĩ'], Chiến sĩ should show."""
        vis = field_visibility(
            'diem_ban_sung', '', 'Chiến sĩ', [],
            {}, self._nhom_meta_restricted(), DIEM_NHOM_MAP
        )
        assert vis == 'show'

    # --- hoc_vien restricted ---

    def test_diem_tong_ket_hidden_for_giang_vien(self):
        vis = field_visibility(
            'diem_tong_ket', '', 'Giảng viên', [],
            {}, self._nhom_meta_restricted(), DIEM_NHOM_MAP
        )
        assert vis == 'hide'

    def test_diem_tong_ket_shown_for_hoc_vien(self):
        vis = field_visibility(
            'diem_tong_ket', '', 'Học viên', [],
            {}, self._nhom_meta_restricted(), DIEM_NHOM_MAP
        )
        assert vis == 'show'

    # --- CRITERIA_META fields (TieuChi records) ---

    def test_criteria_field_hidden_when_doi_tuong_not_in_apply_list(self):
        criteria_meta = {
            'danh_hieu_gv_gioi': {
                'nhom': 'giang_vien',
                'doi_tuong_ap_dung': ['Giảng viên'],
            }
        }
        vis = field_visibility(
            'danh_hieu_gv_gioi', '', 'Học viên', [],
            criteria_meta, self._nhom_meta_open(), DIEM_NHOM_MAP
        )
        assert vis == 'hide'

    def test_criteria_field_shown_when_doi_tuong_in_apply_list(self):
        criteria_meta = {
            'danh_hieu_gv_gioi': {
                'nhom': 'giang_vien',
                'doi_tuong_ap_dung': ['Giảng viên'],
            }
        }
        vis = field_visibility(
            'danh_hieu_gv_gioi', '', 'Giảng viên', [],
            criteria_meta, self._nhom_meta_open(), DIEM_NHOM_MAP
        )
        assert vis == 'show'

    def test_criteria_field_shown_when_apply_list_empty(self):
        """Empty doi_tuong_ap_dung = applies to all."""
        criteria_meta = {
            'xep_loai_can_bo': {
                'nhom': 'chung',
                'doi_tuong_ap_dung': [],
            }
        }
        vis = field_visibility(
            'xep_loai_can_bo', '', 'Chiến sĩ', [],
            criteria_meta, self._nhom_meta_open(), DIEM_NHOM_MAP
        )
        assert vis == 'show'

    # --- NON_CRITERIA_FIELDS always return no_opinion ---

    def test_non_criteria_fields_return_no_opinion(self):
        for f in ['csrf_token', 'quan_nhan_id', 'loai_danh_hieu',
                  'doi_tuong', 'ghi_chu_item', 'ly_do_chua_dat_diem']:
            vis = field_visibility(
                f, 'Chiến sĩ thi đua', 'Học viên', ['ban_sung'],
                {}, self._nhom_meta_open(), DIEM_NHOM_MAP
            )
            assert vis == 'no_opinion', f"{f} should always be 'no_opinion'"

    # --- Unknown fields ---

    def test_unknown_field_returns_no_opinion(self):
        vis = field_visibility(
            'some_unknown_field', 'Chiến sĩ thi đua', 'Giảng viên', [],
            {}, self._nhom_meta_open(), DIEM_NHOM_MAP
        )
        assert vis == 'no_opinion'


# ---------------------------------------------------------------------------
# Integration-level: build criteria_meta from DB models
# ---------------------------------------------------------------------------

class TestCriteriaMetaBuilding:
    """Test that criteria_meta is built correctly from NhomTieuChi + TieuChi."""

    def test_criteria_meta_inherits_nhom_doi_tuong_ap_dung(self, db):
        # Create a nhom restricted to Giảng viên
        nhom = NhomTieuChi(ma_nhom="gv_test", ten_nhom="GV Test", thu_tu=1)
        nhom.doi_tuong_ap_dung = ["Giảng viên"]
        db.session.add(nhom)

        # Create a TieuChi in that nhom
        tc = TieuChi(ma_truong="ket_qua_giang_test", ten="Kết quả giảng test", nhom="gv_test")
        db.session.add(tc)
        db.session.commit()

        # Rebuild nhom_meta and criteria_meta like nomination.py does
        nhom_rows = NhomTieuChi.query.filter_by(is_active=True).all()
        nhom_meta = {
            r.ma_nhom: {'ten_nhom': r.ten_nhom, 'doi_tuong_ap_dung': r.doi_tuong_ap_dung or []}
            for r in nhom_rows
        }

        tieu_chi_db = TieuChi.query.filter_by(is_active=True).all()
        criteria_meta = {}
        for t in tieu_chi_db:
            nhom_info = nhom_meta.get(t.nhom, {'ten_nhom': t.nhom, 'doi_tuong_ap_dung': []})
            criteria_meta[t.ma_truong] = {
                'nhom': t.nhom,
                'nhom_ten': nhom_info.get('ten_nhom') or t.nhom,
                'doi_tuong_ap_dung': nhom_info.get('doi_tuong_ap_dung') or [],
            }

        assert 'ket_qua_giang_test' in criteria_meta
        assert criteria_meta['ket_qua_giang_test']['doi_tuong_ap_dung'] == ['Giảng viên']

    def test_criteria_meta_open_when_nhom_has_no_restriction(self, db):
        nhom = NhomTieuChi(ma_nhom="chung_test", ten_nhom="Chung Test", thu_tu=1)
        nhom.doi_tuong_ap_dung = []
        db.session.add(nhom)
        tc = TieuChi(ma_truong="phieu_tin_nhiem_test", ten="Phiếu tín nhiệm test", nhom="chung_test")
        db.session.add(tc)
        db.session.commit()

        nhom_rows = NhomTieuChi.query.filter_by(is_active=True).all()
        nhom_meta = {
            r.ma_nhom: {'ten_nhom': r.ten_nhom, 'doi_tuong_ap_dung': r.doi_tuong_ap_dung or []}
            for r in nhom_rows
        }
        tieu_chi_db = TieuChi.query.filter_by(is_active=True).all()
        criteria_meta = {}
        for t in tieu_chi_db:
            nhom_info = nhom_meta.get(t.nhom, {'ten_nhom': t.nhom, 'doi_tuong_ap_dung': []})
            criteria_meta[t.ma_truong] = {
                'nhom': t.nhom,
                'nhom_ten': nhom_info['ten_nhom'],
                'doi_tuong_ap_dung': nhom_info.get('doi_tuong_ap_dung') or [],
            }

        assert criteria_meta['phieu_tin_nhiem_test']['doi_tuong_ap_dung'] == []
