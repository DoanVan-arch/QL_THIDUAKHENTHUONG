"""
Tests for NhomTieuChi model — JSON serialization/deserialization of doi_tuong_ap_dung.
"""
import pytest
from app.models.evaluation import NhomTieuChi, DiemQuyDinhDanhHieu


class TestNhomTieuChiModel:
    """Unit tests for NhomTieuChi model property behaviour."""

    def test_doi_tuong_ap_dung_empty_list(self, db):
        nhom = NhomTieuChi(ma_nhom="test_chung", ten_nhom="Test Chung", thu_tu=1)
        nhom.doi_tuong_ap_dung = []
        db.session.add(nhom)
        db.session.commit()

        loaded = NhomTieuChi.query.filter_by(ma_nhom="test_chung").first()
        assert loaded is not None
        assert loaded.doi_tuong_ap_dung == []

    def test_doi_tuong_ap_dung_single_value(self, db):
        nhom = NhomTieuChi(ma_nhom="test_gv", ten_nhom="Test GV", thu_tu=2)
        nhom.doi_tuong_ap_dung = ["Giảng viên"]
        db.session.add(nhom)
        db.session.commit()

        loaded = NhomTieuChi.query.filter_by(ma_nhom="test_gv").first()
        assert loaded.doi_tuong_ap_dung == ["Giảng viên"]

    def test_doi_tuong_ap_dung_multiple_values(self, db):
        nhom = NhomTieuChi(ma_nhom="test_multi", ten_nhom="Test Multi", thu_tu=3)
        nhom.doi_tuong_ap_dung = ["Giảng viên", "Học viên", "Chiến sĩ"]
        db.session.add(nhom)
        db.session.commit()

        loaded = NhomTieuChi.query.filter_by(ma_nhom="test_multi").first()
        assert set(loaded.doi_tuong_ap_dung) == {"Giảng viên", "Học viên", "Chiến sĩ"}

    def test_doi_tuong_ap_dung_none_returns_empty_list(self, db):
        nhom = NhomTieuChi(ma_nhom="test_none", ten_nhom="Test None", thu_tu=4)
        nhom._doi_tuong_ap_dung = None
        db.session.add(nhom)
        db.session.commit()

        loaded = NhomTieuChi.query.filter_by(ma_nhom="test_none").first()
        assert loaded.doi_tuong_ap_dung == []

    def test_doi_tuong_ap_dung_invalid_json_returns_empty_list(self, db):
        nhom = NhomTieuChi(ma_nhom="test_bad", ten_nhom="Test Bad JSON", thu_tu=5)
        nhom._doi_tuong_ap_dung = "not-valid-json{"
        db.session.add(nhom)
        db.session.commit()

        loaded = NhomTieuChi.query.filter_by(ma_nhom="test_bad").first()
        assert loaded.doi_tuong_ap_dung == []

    def test_is_active_default_true(self, db):
        nhom = NhomTieuChi(ma_nhom="test_active", ten_nhom="Test Active", thu_tu=6)
        nhom.doi_tuong_ap_dung = []
        db.session.add(nhom)
        db.session.commit()

        loaded = NhomTieuChi.query.filter_by(ma_nhom="test_active").first()
        assert loaded.is_active is True


class TestDiemQuyDinhDanhHieu:
    """Unit tests for DiemQuyDinhDanhHieu model."""

    def test_create_score_rule(self, db):
        rule = DiemQuyDinhDanhHieu(
            loai_danh_hieu="Chiến sĩ thi đua",
            tieu_chi_field="diem_ban_sung",
            min_diem="7.0",
            is_active=True,
        )
        db.session.add(rule)
        db.session.commit()

        loaded = DiemQuyDinhDanhHieu.query.filter_by(
            loai_danh_hieu="Chiến sĩ thi đua", tieu_chi_field="diem_ban_sung"
        ).first()
        assert loaded is not None
        assert loaded.min_diem == "7.0"

    def test_score_rule_query_by_danh_hieu(self, db):
        unique_danh_hieu = "Chiến sĩ thi đua UNIQUE_QUERY_TEST"
        for field in ["diem_the_luc", "diem_kiem_tra_chinh_tri"]:
            db.session.add(DiemQuyDinhDanhHieu(
                loai_danh_hieu=unique_danh_hieu,
                tieu_chi_field=field,
                min_diem="7.5",
                is_active=True,
            ))
        db.session.commit()

        rules = DiemQuyDinhDanhHieu.query.filter_by(
            loai_danh_hieu=unique_danh_hieu, is_active=True
        ).all()
        assert len(rules) == 2
        fields = {r.tieu_chi_field for r in rules}
        assert "diem_the_luc" in fields
        assert "diem_kiem_tra_chinh_tri" in fields
