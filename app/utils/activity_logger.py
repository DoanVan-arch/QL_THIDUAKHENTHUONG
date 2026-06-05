"""Utility for writing ActivityLog records.

Usage:
    from app.utils.activity_logger import log_action

    log_action('submit_nomination', resource_type='de_xuat', resource_id=dx.id,
                detail=f'Năm học {dx.nam_hoc}')
"""
from flask import request as _flask_request
from flask_login import current_user as _current_user


def log_action(action, resource_type=None, resource_id=None, detail=None, user=None):
    """Write one ActivityLog record.

    Silently swallows any exception so logging never breaks the main request.
    Must be called within an active Flask application context.
    """
    try:
        # Lazy import to avoid circular dependencies at module load time
        from app.extensions import db
        from app.models.activity_log import ActivityLog

        actor = user or (_current_user if _current_user and _current_user.is_authenticated else None)

        ip = None
        try:
            ip = _flask_request.remote_addr
        except RuntimeError:
            pass  # outside request context

        entry = ActivityLog(
            user_id=actor.id if actor else None,
            username=actor.username if actor else None,
            ho_ten=actor.ho_ten if actor else None,
            role=actor.role.value if actor and actor.role else None,
            action=action,
            resource_type=resource_type,
            resource_id=resource_id,
            detail=detail,
            ip_address=ip,
        )
        db.session.add(entry)
        # Use a nested savepoint so a log failure never rolls back the real transaction
        try:
            db.session.flush()
        except Exception:
            db.session.rollback()
    except Exception:
        pass  # logging must never break the caller
