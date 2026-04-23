"""add ten_don_vi_de_xuat to de_xuat_chi_tiet

Revision ID: 562238791222
Revises: 8a0dd0255cb4
Create Date: 2026-04-22 23:11:11.840678

"""
from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = '562238791222'
down_revision = '8a0dd0255cb4'
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table('de_xuat_chi_tiet', schema=None) as batch_op:
        batch_op.add_column(sa.Column('ten_don_vi_de_xuat', sa.String(length=255), nullable=True))


def downgrade():
    with op.batch_alter_table('de_xuat_chi_tiet', schema=None) as batch_op:
        batch_op.drop_column('ten_don_vi_de_xuat')
