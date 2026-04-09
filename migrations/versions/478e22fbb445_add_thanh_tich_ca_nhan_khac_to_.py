"""Add thanh_tich_ca_nhan_khac to DeXuatChiTiet

Revision ID: 478e22fbb445
Revises: 
Create Date: 2026-04-08 13:45:53.727809

"""
from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = '478e22fbb445'
down_revision = None
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table('de_xuat_chi_tiet', schema=None) as batch_op:
        batch_op.add_column(sa.Column('thanh_tich_ca_nhan_khac', sa.Text(), nullable=True))


def downgrade():
    with op.batch_alter_table('de_xuat_chi_tiet', schema=None) as batch_op:
        batch_op.drop_column('thanh_tich_ca_nhan_khac')
