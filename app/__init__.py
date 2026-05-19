from flask import Flask, send_from_directory
from flask_login import login_required
from config import Config
from app.extensions import db, login_manager, csrf, migrate


def create_app(config_class=None):
    app = Flask(__name__)
    app.config.from_object(config_class or Config)

    db.init_app(app)
    login_manager.init_app(app)
    csrf.init_app(app)
    migrate.init_app(app, db)

    # Import models so they are registered
    from app import models  # noqa: F401

    # Register blueprints
    from app.routes.auth import auth_bp
    from app.routes.dashboard import dashboard_bp
    from app.routes.personnel import personnel_bp
    from app.routes.nomination import nomination_bp
    from app.routes.approval import approval_bp
    from app.routes.admin import admin_bp
    from app.routes.hoi_dong import hoi_dong_bp

    app.register_blueprint(auth_bp)
    app.register_blueprint(dashboard_bp)
    app.register_blueprint(personnel_bp, url_prefix='/personnel')
    app.register_blueprint(nomination_bp, url_prefix='/nomination')
    app.register_blueprint(approval_bp, url_prefix='/approval')
    app.register_blueprint(admin_bp, url_prefix='/admin')
    app.register_blueprint(hoi_dong_bp, url_prefix='/hoi-dong')

    @app.route('/uploads/<path:filename>')
    @login_required
    def uploaded_file(filename):
        return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

    @app.errorhandler(403)
    def forbidden(e):
        return '<h1>403 - Không có quyền truy cập</h1><a href="/">Về trang chủ</a>', 403

    @app.errorhandler(404)
    def not_found(e):
        return '<h1>404 - Không tìm thấy trang</h1><a href="/">Về trang chủ</a>', 404

    @app.context_processor
    def inject_notification_count():
        from flask_login import current_user
        count = 0
        pending_transfers = 0
        dept_pending_count = 0
        admin_pending_count = 0
        if current_user.is_authenticated:
            if current_user.is_unit_user:
                from app.models.notification import ThongBao
                from app.models.transfer import ChuyenDonVi, TrangThaiChuyen
                count = ThongBao.query.filter_by(
                    user_id=current_user.id, da_doc=False
                ).count()
                if current_user.don_vi_id:
                    pending_transfers = ChuyenDonVi.query.filter_by(
                        don_vi_dich_id=current_user.don_vi_id,
                        trang_thai=TrangThaiChuyen.PENDING,
                    ).count()
            if current_user.is_department:
                # Count PheDuyet records assigned to this dept user's role, not yet reviewed
                try:
                    from app.models.approval import PheDuyet, KetQuaDuyet
                    from app.routes.approval import ROLE_TO_PHONG
                    user_role = current_user.role.value if hasattr(current_user.role, 'value') else str(current_user.role)
                    phong = ROLE_TO_PHONG.get(current_user.role)
                    if phong:
                        dept_pending_count = PheDuyet.query.filter_by(
                            phong_duyet=phong, ket_qua=KetQuaDuyet.CHO_DUYET.value
                        ).count()
                except Exception:
                    dept_pending_count = 0
            if current_user.is_admin:
                # Count nominations not yet final-approved or rejected
                try:
                    from app.models.nomination import DeXuat, TrangThaiDeXuat
                    admin_pending_count = DeXuat.query.filter(
                        DeXuat.trang_thai.notin_([
                            TrangThaiDeXuat.PHE_DUYET_CUOI.value,
                            TrangThaiDeXuat.TU_CHOI.value,
                            TrangThaiDeXuat.NHAP.value,
                        ])
                    ).count()
                except Exception:
                    admin_pending_count = 0
        return dict(
            unread_notification_count=count,
            pending_transfer_count=pending_transfers,
            dept_pending_count=dept_pending_count,
            admin_pending_count=admin_pending_count,
        )

    # Jinja2 filter: clean ngay_nhap_ngu display (strip datetime noise)
    import re as _re
    from datetime import datetime as _dt

    @app.template_filter('clean_nhap_ngu')
    def clean_nhap_ngu_filter(value):
        if not value:
            return ''
        text = str(value).strip()
        if _re.match(r'^\d{2}/\d{4}$', text):
            return text
        for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d'):
            try:
                return _dt.strptime(text, fmt).strftime('%m/%Y')
            except ValueError:
                continue
        if ' ' in text:
            text = text.split(' ')[0]
        return text

    return app
