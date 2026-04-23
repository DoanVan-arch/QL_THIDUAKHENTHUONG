"""
Tests for admin sidebar menu structure — verifies that the base template
contains the required navigation items and groupings.
"""
import pytest
import re


class TestAdminSidebarMenu:
    """Verify admin menu structure in base.html template."""

    def _get_sidebar_html(self, app):
        """Render base.html sidebar portion by parsing it directly."""
        with open('app/templates/base.html', encoding='utf-8') as f:
            return f.read()

    def test_admin_section_exists(self, app):
        html = self._get_sidebar_html(app)
        assert 'QUẢN TRỊ' in html

    def test_admin_has_manage_tieu_chi_link(self, app):
        html = self._get_sidebar_html(app)
        assert "manage_tieu_chi" in html

    def test_admin_has_manage_nhom_tieu_chi_link(self, app):
        html = self._get_sidebar_html(app)
        assert "manage_nhom_tieu_chi" in html

    def test_admin_has_manage_danh_hieu_link(self, app):
        html = self._get_sidebar_html(app)
        assert "manage_danh_hieu" in html

    def test_admin_has_manage_chuc_vu_link(self, app):
        html = self._get_sidebar_html(app)
        assert "manage_chuc_vu" in html

    def test_admin_has_manage_doi_tuong_link(self, app):
        html = self._get_sidebar_html(app)
        assert "manage_doi_tuong" in html

    def test_admin_has_manage_cap_bac_link(self, app):
        html = self._get_sidebar_html(app)
        assert "manage_cap_bac" in html

    def test_admin_has_diem_quy_dinh_link(self, app):
        html = self._get_sidebar_html(app)
        assert "manage_diem_quy_dinh" in html

    def test_admin_has_report_link(self, app):
        html = self._get_sidebar_html(app)
        assert "report_summary" in html

    def test_admin_has_approval_tracking_link(self, app):
        html = self._get_sidebar_html(app)
        assert "approval_tracking" in html

    def test_danh_muc_group_exists(self, app):
        """After the sidebar refactor, Danh mục group should exist."""
        html = self._get_sidebar_html(app)
        assert 'DANH MỤC' in html or 'danh-muc-group' in html or 'danh_muc' in html.lower() \
               or 'Danh mục' in html

    def test_tieu_chi_group_exists(self, app):
        """After the sidebar refactor, Tiêu chí group should exist."""
        html = self._get_sidebar_html(app)
        assert 'TIÊU CHÍ' in html or 'tieu-chi-group' in html or 'Tiêu chí' in html
