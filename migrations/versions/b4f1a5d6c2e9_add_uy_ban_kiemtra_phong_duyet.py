"""add uy ban kiem tra phong duyet

Revision ID: b4f1a5d6c2e9
Revises: 9b2fd1c4e6aa
Create Date: 2026-04-15 00:00:00.000000

"""
from alembic import op


revision = 'b4f1a5d6c2e9'
down_revision = '9b2fd1c4e6aa'
branch_labels = None
depends_on = None


def upgrade():
    # No schema change required; this migration marks code-level enum/mapping update.
    pass


def downgrade():
    pass
