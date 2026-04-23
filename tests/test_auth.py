"""
Tests for authentication routes (login/logout).
"""
import pytest


class TestAuthRoutes:

    def test_login_page_loads(self, client):
        resp = client.get('/login')
        assert resp.status_code == 200
        assert 'đăng nhập'.encode() in resp.data.lower() or b'login' in resp.data.lower() or b'username' in resp.data.lower() or 'tên' in resp.data.decode('utf-8', errors='ignore').lower()

    def test_login_wrong_password(self, client, unit_user):
        resp = client.post('/login', data={
            'username': unit_user.username,
            'password': 'wrongpass',
        }, follow_redirects=True)
        assert resp.status_code == 200
        # Should stay on login page or show error, not redirect to dashboard
        body = resp.data.decode('utf-8', errors='ignore').lower()
        assert 'dashboard' not in body or 'sai' in body or 'không' in body or 'incorrect' in body or 'invalid' in body

    def test_login_correct_credentials(self, client, unit_user):
        resp = client.post('/login', data={
            'username': unit_user.username,
            'password': 'password123',
        }, follow_redirects=True)
        assert resp.status_code == 200

    def test_unauthenticated_redirect_to_login(self, client):
        resp = client.get('/nomination/', follow_redirects=False)
        assert resp.status_code in (302, 401)

    def test_admin_route_blocked_for_unit_user(self, logged_in_unit):
        resp = logged_in_unit.get('/admin/users', follow_redirects=False)
        assert resp.status_code in (302, 403)
