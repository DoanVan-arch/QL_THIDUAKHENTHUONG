"""add chuc_vu_option table

Revision ID: d1f35b5ad8ce
Revises: b8b0f3f9a8f1
Create Date: 2026-04-14 00:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


revision = 'd1f35b5ad8ce'
down_revision = 'b8b0f3f9a8f1'
branch_labels = None
depends_on = None


def upgrade():
    op.create_table(
        'chuc_vu_option',
        sa.Column('id', sa.Integer(), autoincrement=True, nullable=False),
        sa.Column('ten', sa.String(length=120), nullable=False),
        sa.Column('thu_tu', sa.Integer(), nullable=True),
        sa.Column('is_active', sa.Boolean(), nullable=True),
        sa.Column('created_at', sa.DateTime(), nullable=True),
        sa.Column('updated_at', sa.DateTime(), nullable=True),
        sa.PrimaryKeyConstraint('id'),
        sa.UniqueConstraint('ten')
    )

    default_values = [
        'Hiệu trưởng, Chính ủy',
        'Phó Giám đốc, Phó Hiệu trưởng, Phó Chính ủy',
        'Chủ nhiệm',
        'Phó Chủ nhiệm',
        'Chủ nhiệm khoa',
        'Phó Chủ nhiệm khoa',
        'Hệ trưởng, Chính trị viên hệ',
        'Phó hệ trưởng',
        'Chủ nhiệm bộ môn',
        'Phó Chủ nhiệm bộ môn',
        'Tiểu đoàn trưởng, Chính trị viên',
        'Phó tiểu đoàn trưởng, Chính trị viên phó',
        'Đại đội trưởng, Chính trị viên',
        'Phó Đại đội trưởng, Chính trị viên phó',
        'Trung đội trưởng',
        'Chủ nhiệm lớp',
        'Phó chủ nhiệm lớp',
        'Giảng viên',
        'Trợ giảng',
        'Trưởng phòng',
        'Phó trưởng phòng',
        'Trưởng ban',
        'Phó trưởng ban',
        'Trợ lý',
        'Nhân viên',
        'Y sĩ',
        'Bệnh xá trưởng',
        'Phó bệnh xá trưởng',
        'Bác sĩ',
        'Trợ lý chính trị',
        'Trợ lý hậu cần',
        'Trợ lý tham mưu',
        'Chủ nhiệm nhà văn hóa',
        'Lái xe',
        'Trạm trưởng',
    ]
    for i, ten in enumerate(default_values, start=1):
        op.execute(
            sa.text(
                "INSERT INTO chuc_vu_option (ten, thu_tu, is_active, created_at, updated_at) "
                "VALUES (:ten, :thu_tu, 1, NOW(), NOW())"
            ).bindparams(ten=ten, thu_tu=i)
        )


def downgrade():
    op.drop_table('chuc_vu_option')
