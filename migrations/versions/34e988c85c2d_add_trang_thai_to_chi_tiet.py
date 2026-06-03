"""add_trang_thai_to_chi_tiet

Revision ID: 34e988c85c2d
Revises: l9m0n1o2p3q4
Create Date: 2026-06-03 09:41:36.181603

"""
from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = '34e988c85c2d'
down_revision = 'l9m0n1o2p3q4'
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table('de_xuat_chi_tiet', schema=None) as batch_op:
        batch_op.add_column(sa.Column('trang_thai', sa.String(length=30),
                                      server_default='nhap', nullable=False))


def downgrade():
    with op.batch_alter_table('de_xuat_chi_tiet', schema=None) as batch_op:
        batch_op.drop_column('trang_thai')
