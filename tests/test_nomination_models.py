"""
Tests for TieuChi and DanhHieu models.
"""
import pytest
from app.models.nomination import TieuChi, DanhHieu


class TestTieuChiModel:
    """Unit tests for TieuChi (criteria) model."""

    def test_create_tieu_chi(self, db):
        tc = TieuChi(ma_truong="ban_sung", ten="Bắn súng", nhom="chung")
        db.session.add(tc)
        db.session.commit()

        loaded = TieuChi.query.filter_by(ma_truong="ban_sung").first()
        assert loaded is not None
        assert loaded.ten == "Bắn súng"
        assert loaded.nhom == "chung"

    def test_default_is_active(self, db):
        tc = TieuChi(ma_truong="the_luc", ten="Thể lực", nhom="chung")
        db.session.add(tc)
        db.session.commit()

        loaded = TieuChi.query.filter_by(ma_truong="the_luc").first()
        assert loaded.is_active is True

    def test_phong_duyet_json_roundtrip(self, db):
        tc = TieuChi(ma_truong="chinh_tri", ten="Chính trị", nhom="chung")
        tc.phong_duyet = ["phong_chinhTri", "ban_canBo"]
        db.session.add(tc)
        db.session.commit()

        loaded = TieuChi.query.filter_by(ma_truong="chinh_tri").first()
        assert "phong_chinhTri" in loaded.phong_duyet
        assert "ban_canBo" in loaded.phong_duyet

    def test_phong_duyet_none_returns_empty_list(self, db):
        tc = TieuChi(ma_truong="tin_hoc", ten="Tin học", nhom="chung")
        tc._phong_duyet = None
        db.session.add(tc)
        db.session.commit()

        loaded = TieuChi.query.filter_by(ma_truong="tin_hoc").first()
        assert loaded.phong_duyet == []

    def test_nhom_choices_contains_expected_keys(self):
        assert "chung" in TieuChi.NHOM_CHOICES
        assert "giang_vien" in TieuChi.NHOM_CHOICES
        assert "hoc_vien" in TieuChi.NHOM_CHOICES
        assert "nckh" in TieuChi.NHOM_CHOICES
        assert "khac" in TieuChi.NHOM_CHOICES

    def test_filter_active_only(self, db):
        db.session.add(TieuChi(ma_truong="tc_active", ten="Active", nhom="chung", is_active=True))
        db.session.add(TieuChi(ma_truong="tc_inactive", ten="Inactive", nhom="chung", is_active=False))
        db.session.commit()

        active = TieuChi.query.filter_by(is_active=True).all()
        inactive = TieuChi.query.filter_by(is_active=False).all()
        active_fields = {tc.ma_truong for tc in active}
        inactive_fields = {tc.ma_truong for tc in inactive}
        assert "tc_active" in active_fields
        assert "tc_inactive" not in active_fields
        assert "tc_inactive" in inactive_fields


class TestDanhHieuModel:
    """Unit tests for DanhHieu (award) model."""

    def test_create_danh_hieu(self, db):
        dh = DanhHieu(
            ten_danh_hieu="Chiến sĩ thi đua",
            ma_danh_hieu="CSTD",
            pham_vi="Cá nhân",
            thu_tu=1,
        )
        dh.tieu_chi = ["ban_sung", "the_luc", "kiem_tra_chinh_tri"]
        db.session.add(dh)
        db.session.commit()

        loaded = DanhHieu.query.filter_by(ten_danh_hieu="Chiến sĩ thi đua").first()
        assert loaded is not None
        assert "ban_sung" in loaded.tieu_chi
        assert "the_luc" in loaded.tieu_chi

    def test_tieu_chi_empty_returns_empty_list(self, db):
        dh = DanhHieu(
            ten_danh_hieu="Chiến sĩ tiên tiến",
            ma_danh_hieu="CSTT",
            pham_vi="Cá nhân",
        )
        dh._tieu_chi = None
        db.session.add(dh)
        db.session.commit()

        loaded = DanhHieu.query.filter_by(ten_danh_hieu="Chiến sĩ tiên tiến").first()
        assert loaded.tieu_chi == []

    def test_tieu_chi_invalid_json_returns_empty_list(self, db):
        dh = DanhHieu(
            ten_danh_hieu="Đơn vị quyết thắng",
            ma_danh_hieu="DVQT",
            pham_vi="Đơn vị",
        )
        dh._tieu_chi = "not-json{"
        db.session.add(dh)
        db.session.commit()

        loaded = DanhHieu.query.filter_by(ten_danh_hieu="Đơn vị quyết thắng").first()
        assert loaded.tieu_chi == []

    def test_query_active_danh_hieu(self, db):
        db.session.add(DanhHieu(ten_danh_hieu="Active DH", ma_danh_hieu="ADH", pham_vi="Cá nhân", is_active=True))
        db.session.add(DanhHieu(ten_danh_hieu="Inactive DH", ma_danh_hieu="IDH", pham_vi="Cá nhân", is_active=False))
        db.session.commit()

        active = DanhHieu.query.filter_by(is_active=True).all()
        names = {dh.ten_danh_hieu for dh in active}
        assert "Active DH" in names
        assert "Inactive DH" not in names
