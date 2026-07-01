import os
from app import create_app

app = create_app()

# Ensure upload subdirectories exist on every startup
# (critical after Railway persistent volume mount or first deploy)
with app.app_context():
    for subfolder in ('evidence', 'certificates'):
        path = os.path.join(app.config['UPLOAD_FOLDER'], subfolder)
        os.makedirs(path, exist_ok=True)

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=5000, threaded=True)  # Enable threaded mode for handling multiple requests
