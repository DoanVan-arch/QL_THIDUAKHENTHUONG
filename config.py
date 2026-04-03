import os


class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY', 'sqct-thidua-khenthuong-secret-2025')
    SQLALCHEMY_DATABASE_URI = 'mysql://root:jSeoBzPcgCwPFEVbsozvKjwtavpAfkdj@junction.proxy.rlwy.net:33521/quanly_thidua_khenthuong'
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    SQLALCHEMY_ENGINE_OPTIONS = {
        'pool_recycle': 3600,
        'pool_pre_ping': True,
    }
    UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB
    ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'pdf'}
    SESSION_COOKIE_HTTPONLY = True
    SESSION_COOKIE_SAMESITE = 'Lax'
