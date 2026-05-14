"""
Script xóa quân nhân trùng CCCD - dùng bulk CASE UPDATE.
"""
from app import create_app
from app.extensions import db
from sqlalchemy import text

app = create_app()

FK_TABLES = [
    ('de_xuat_chi_tiet', 'quan_nhan_id'),
    ('chuyen_don_vi', 'quan_nhan_id'),
    ('chung_chi', 'quan_nhan_id'),
    ('danh_gia_hang_nam', 'quan_nhan_id'),
    ('khen_thuong', 'quan_nhan_id'),
]

with app.app_context():
    rows = db.session.execute(text('''
        SELECT MIN(id) AS keeper_id,
               GROUP_CONCAT(id ORDER BY id) AS all_ids
        FROM quan_nhan
        WHERE can_cuoc_cong_dan IS NOT NULL AND can_cuoc_cong_dan != ''
        GROUP BY can_cuoc_cong_dan
        HAVING COUNT(*) > 1
    ''')).fetchall()

    print(f'Found {len(rows)} duplicate groups')

    remap = {}
    for r in rows:
        keeper_id = int(r[0])
        all_ids = [int(x) for x in r[1].split(',')]
        for i in all_ids:
            if i != keeper_id:
                remap[i] = keeper_id

    print(f'Will delete {len(remap)} duplicate records')
    if not remap:
        print('Nothing to do.')
    else:
        old_ids = list(remap.keys())
        # Build CASE expression: CASE WHEN id=x THEN keeper ... END
        case_parts = ' '.join(f'WHEN {old} THEN {keeper}' for old, keeper in remap.items())
        in_list = ','.join(str(i) for i in old_ids)
        case_expr = f'CASE quan_nhan_id {case_parts} END'

        for table, col in FK_TABLES:
            sql = text(f'''
                UPDATE {table}
                SET {col} = {case_expr}
                WHERE {col} IN ({in_list})
            ''')
            cnt = db.session.execute(sql).rowcount
            print(f'  {table}: updated {cnt} rows')

        # Delete in batches
        batch_size = 500
        total_deleted = 0
        for i in range(0, len(old_ids), batch_size):
            batch = old_ids[i:i+batch_size]
            ids_str = ','.join(str(x) for x in batch)
            cnt = db.session.execute(text(f'DELETE FROM quan_nhan WHERE id IN ({ids_str})')).rowcount
            total_deleted += cnt

        db.session.commit()
        print(f'Deleted {total_deleted} duplicate records')

    remaining = db.session.execute(text('''
        SELECT COUNT(*) FROM (
            SELECT can_cuoc_cong_dan FROM quan_nhan
            WHERE can_cuoc_cong_dan IS NOT NULL AND can_cuoc_cong_dan != ''
            GROUP BY can_cuoc_cong_dan HAVING COUNT(*) > 1
        ) t
    ''')).scalar()
    print(f'Remaining duplicate groups: {remaining}')
