"""add session tracking fields to users

Revision ID: e2f3a4b5c6d7
Revises: d1e2f3a4b5c6
Create Date: 2026-05-21
"""
from alembic import op
import sqlalchemy as sa

revision = 'e2f3a4b5c6d7'
down_revision = 'd1e2f3a4b5c6'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column('users', sa.Column('session_token', sa.String(64), nullable=True))
    op.add_column('users', sa.Column('last_login_ip', sa.String(45), nullable=True))
    op.add_column('users', sa.Column('last_login_at', sa.DateTime(), nullable=True))
    op.add_column('users', sa.Column('last_login_device', sa.String(256), nullable=True))
    op.create_index('ix_users_session_token', 'users', ['session_token'], unique=False)


def downgrade():
    op.drop_index('ix_users_session_token', table_name='users')
    op.drop_column('users', 'last_login_device')
    op.drop_column('users', 'last_login_at')
    op.drop_column('users', 'last_login_ip')
    op.drop_column('users', 'session_token')
