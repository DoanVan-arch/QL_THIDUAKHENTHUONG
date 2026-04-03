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

    app.register_blueprint(auth_bp)
    app.register_blueprint(dashboard_bp)
    app.register_blueprint(personnel_bp, url_prefix='/personnel')
    app.register_blueprint(nomination_bp, url_prefix='/nomination')
    app.register_blueprint(approval_bp, url_prefix='/approval')
    app.register_blueprint(admin_bp, url_prefix='/admin')

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
        if current_user.is_authenticated and current_user.is_unit_user:
            from app.models.notification import ThongBao
            count = ThongBao.query.filter_by(
                user_id=current_user.id, da_doc=False
            ).count()
        return dict(unread_notification_count=count)

    return app
