"""
Shared pytest fixtures for all tests.
Uses an in-memory SQLite DB so tests never touch the real MySQL instance.
"""
import pytest
from app import create_app
from app.extensions import db as _db
from config import Config


class TestConfig(Config):
    TESTING = True
    WTF_CSRF_ENABLED = False
    SQLALCHEMY_DATABASE_URI = "sqlite:///:memory:"
    SQLALCHEMY_ENGINE_OPTIONS = {}          # disable MySQL-specific options
    SERVER_NAME = "localhost"
    SECRET_KEY = "test-secret"


@pytest.fixture(scope="session")
def app():
    """Create application with test config, create all tables once per session."""
    application = create_app(TestConfig)
    with application.app_context():
        _db.create_all()
        yield application
        _db.drop_all()


@pytest.fixture(scope="function")
def db(app):
    """Provide a clean DB state per test by dropping and recreating all tables."""
    with app.app_context():
        _db.drop_all()
        _db.create_all()
        yield _db
        _db.session.remove()


@pytest.fixture(scope="function")
def client(app, db):
    """Flask test client."""
    return app.test_client()


@pytest.fixture(scope="function")
def admin_user(db):
    """Create and return an admin user."""
    from app.models.user import User, Role
    user = User(
        username="admin_test",
        role=Role.ADMIN,
        ho_ten="Admin Test",
    )
    user.set_password("password123")
    db.session.add(user)
    db.session.commit()
    return user


@pytest.fixture(scope="function")
def unit_user(db):
    """Create and return a unit user."""
    from app.models.user import User, Role
    user = User(
        username="unit_test",
        role=Role.UNIT_USER,
        ho_ten="Unit Test",
    )
    user.set_password("password123")
    db.session.add(user)
    db.session.commit()
    return user


@pytest.fixture(scope="function")
def logged_in_admin(client, admin_user, app):
    """Client with admin session active."""
    with client.session_transaction() as sess:
        sess["_user_id"] = str(admin_user.id)
        sess["_fresh"] = True
    return client


@pytest.fixture(scope="function")
def logged_in_unit(client, unit_user, app):
    """Client with unit user session active."""
    with client.session_transaction() as sess:
        sess["_user_id"] = str(unit_user.id)
        sess["_fresh"] = True
    return client
