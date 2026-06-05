from datetime import datetime
from flask import Blueprint, render_template, redirect, url_for, flash, request, session
from flask_login import login_user, logout_user, login_required, current_user
from app.models.user import User
from app.extensions import db
from app.utils.activity_logger import log_action

auth_bp = Blueprint('auth', __name__)


@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard.index'))

    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')

        user = User.query.filter_by(username=username).first()
        if user and user.check_password(password) and user.is_active:
            # Generate new session token — invalidates any existing session on other devices
            token = user.generate_session_token()
            user.last_login_ip = request.remote_addr
            user.last_login_at = datetime.utcnow()
            user.last_login_device = request.user_agent.string[:256] if request.user_agent else None
            db.session.commit()

            login_user(user, remember=request.form.get('remember'))
            session['_session_token'] = token
            log_action('login', user=user, detail=f'IP: {request.remote_addr}')
            db.session.commit()
            next_page = request.args.get('next')
            flash(f'Chào mừng {user.ho_ten}!', 'success')
            return redirect(next_page or url_for('dashboard.index'))

        log_action('login_failed', detail=f'username={username}, IP={request.remote_addr}')
        db.session.commit()
        flash('Tên đăng nhập hoặc mật khẩu không đúng.', 'danger')

    return render_template('auth/login.html')


@auth_bp.route('/logout')
@login_required
def logout():
    if current_user.is_authenticated:
        log_action('logout')
        current_user.session_token = None
        db.session.commit()
    logout_user()
    session.pop('_session_token', None)
    flash('Đã đăng xuất thành công.', 'info')
    return redirect(url_for('auth.login'))
