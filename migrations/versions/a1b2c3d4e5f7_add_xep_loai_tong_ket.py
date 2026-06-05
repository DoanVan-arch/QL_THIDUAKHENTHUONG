"""add xep_loai_tong_ket to de_xuat_chi_tiet

Revision ID: a1b2c3d4e5f7
Revises: 34e988c85c2d
Create Date: 2026-06-05 00:00:00.000000

"""
from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = 'a1b2c3d4e5f7'
down_revision = '34e988c85c2d'
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table('de_xuat_chi_tiet', schema=None) as batch_op:
        batch_op.add_column(sa.Column('xep_loai_tong_ket', sa.String(length=50), nullable=True))


def downgrade():
    with op.batch_alter_table('de_xuat_chi_tiet', schema=None) as batch_op:
        batch_op.drop_column('xep_loai_tong_ket')
