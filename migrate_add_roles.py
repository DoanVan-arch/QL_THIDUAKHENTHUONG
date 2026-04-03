"""
Migration: Add BAN_CANBO and BAN_QUANLUC to users.role ENUM column.

The MySQL ENUM column type has a fixed set of allowed values defined at the
database level.  When new Python enum members were added, the DB column was
not updated, causing "Data truncated for column 'role'" errors.

Run:  python migrate_add_roles.py
"""
import pymysql

conn = pymysql.connect(
    host='localhost', port=3306,
    user='root', password='1111',
    database='quanly_thidua_khenthuong',
)
cursor = conn.cursor()

# Alter the ENUM to include all current roles
alter_sql = """
ALTER TABLE users
MODIFY COLUMN role ENUM(
    'UNIT_USER',
    'PHONG_CHINHTRI',
    'PHONG_THAMMUU',
    'PHONG_KHOAHOC',
    'PHONG_DAOTAO',
    'BAN_CANBO',
    'BAN_QUANLUC',
    'ADMIN'
) NOT NULL
"""

try:
    cursor.execute(alter_sql)
    conn.commit()
    print("OK: users.role ENUM updated — added BAN_CANBO, BAN_QUANLUC.")
except Exception as e:
    print(f"ERROR: {e}")
    conn.rollback()
finally:
    cursor.close()
    conn.close()
