import os


class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY', 'sqct-thidua-khenthuong-secret-2025')
    #SECRET_KEY = os.getenv('SECRET_KEY', 'dev-secret-key-change-in-production')
    
    # Database Configuration - MariaDB/MySQL
    SQLALCHEMY_DATABASE_URI = os.getenv('DATABASE_URL') or \
        'mysql+pymysql://root:1111@localhost:3306/quanly_thidua_khenthuong?charset=utf8mb4'
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    SQLALCHEMY_ENGINE_OPTIONS = {
        'pool_pre_ping': True,
        'pool_recycle': 300,
        'pool_size': 10,
        'max_overflow': 20
    }
    
    # JWT Configuration
    JWT_SECRET_KEY = os.getenv('JWT_SECRET_KEY', 'jwt-secret-key')
    JWT_ALGORITHM = os.getenv('JWT_ALGORITHM', 'HS256')
    ACCESS_TOKEN_EXPIRE_MINUTES = int(os.getenv('ACCESS_TOKEN_EXPIRE_MINUTES', 1440))
    
    # Admin Configuration
    ADMIN_USERNAME = os.getenv('ADMIN_USERNAME', 'admin')
    ADMIN_PASSWORD = os.getenv('ADMIN_PASSWORD', 'admin123')
    # SQLALCHEMY_DATABASE_URI = os.environ.get(
    #     'DATABASE_URL',
    #     'mysql+pymysql://root:jSeoBzPcgCwPFEVbsozvKjwtavpAfkdj@junction.proxy.rlwy.net:33521/quanly_thidua_khenthuong'
    # )
    # SQLALCHEMY_TRACK_MODIFICATIONS = False
    # SQLALCHEMY_ENGINE_OPTIONS = {
    #     'pool_recycle': 3600,
    #     'pool_pre_ping': True,
    # }
    # UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
    # MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB
    # ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'pdf'}
    # SESSION_COOKIE_HTTPONLY = True
    # SESSION_COOKIE_SAMESITE = 'Lax'
