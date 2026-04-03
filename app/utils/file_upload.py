import os
import uuid
from werkzeug.utils import secure_filename
from flask import current_app


ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'pdf'}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def save_upload(file, subfolder='evidence'):
    if not file or not file.filename or not allowed_file(file.filename):
        return None

    ext = file.filename.rsplit('.', 1)[1].lower()
    unique_name = f"{uuid.uuid4().hex}.{ext}"

    target_dir = os.path.join(current_app.config['UPLOAD_FOLDER'], subfolder)
    os.makedirs(target_dir, exist_ok=True)

    filepath = os.path.join(target_dir, unique_name)
    file.save(filepath)

    return f"{subfolder}/{unique_name}"


def delete_upload(relative_path):
    if not relative_path:
        return
    full_path = os.path.join(current_app.config['UPLOAD_FOLDER'], relative_path)
    if os.path.exists(full_path):
        os.remove(full_path)
