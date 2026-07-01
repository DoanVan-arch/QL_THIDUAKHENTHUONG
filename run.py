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
    # ★ QUAN TRỌNG: threaded=True — nếu không, Werkzeug dev server chỉ xử lý được
    # 1 request tại 1 thời điểm (đơn luồng). Một request xuất Word chậm (10-30s+)
    # sẽ khiến MỌI user khác bị "treo"/chờ cho tới khi request đó xong — đây là
    # nguyên nhân khả dĩ nhất gây ra hiện tượng "request không hoàn thành".
    # debug=True chỉ nên bật khi phát triển local, KHÔNG bật ở production (Werkzeug
    # debugger cho phép thực thi mã tùy ý nếu bị lộ ra internet).
    is_production = bool(os.environ.get('RAILWAY_ENVIRONMENT'))
    port = int(os.environ.get('PORT', 5000))
    app.run(
        debug=not is_production,
        host='0.0.0.0',
        port=port,
        threaded=True,
    )
