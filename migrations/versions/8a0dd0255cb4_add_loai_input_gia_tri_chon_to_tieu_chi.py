"""add loai_input gia_tri_chon to tieu_chi

Revision ID: 8a0dd0255cb4
Revises: d2a7f1c9e4b6
Create Date: 2026-04-22 21:51:10.336341

"""
from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = '8a0dd0255cb4'
down_revision = 'd2a7f1c9e4b6'
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table('tieu_chi', schema=None) as batch_op:
        batch_op.add_column(sa.Column('loai_input', sa.String(length=20), nullable=False, server_default='textbox'))
        batch_op.add_column(sa.Column('gia_tri_chon', sa.Text(), nullable=True))


def downgrade():
    with op.batch_alter_table('tieu_chi', schema=None) as batch_op:
        batch_op.drop_column('gia_tri_chon')
        batch_op.drop_column('loai_input')
