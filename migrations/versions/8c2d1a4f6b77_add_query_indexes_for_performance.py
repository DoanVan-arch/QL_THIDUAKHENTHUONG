"""add query indexes for performance

Revision ID: 8c2d1a4f6b77
Revises: 5f4a8c2b9d11
Create Date: 2026-04-19 16:30:00.000000

"""
from alembic import op


revision = '8c2d1a4f6b77'
down_revision = '5f4a8c2b9d11'
branch_labels = None
depends_on = None


def upgrade():
    # de_xuat: common filters/order
    op.create_index('ix_de_xuat_don_vi_trang_thai', 'de_xuat', ['don_vi_id', 'trang_thai'], unique=False)
    op.create_index('ix_de_xuat_trang_thai_ngay_gui', 'de_xuat', ['trang_thai', 'ngay_gui'], unique=False)
    op.create_index('ix_de_xuat_nam_hoc_trang_thai', 'de_xuat', ['nam_hoc', 'trang_thai'], unique=False)

    # de_xuat_chi_tiet: frequent joins/filters
    op.create_index('ix_dxct_nam_hoc_quan_nhan', 'de_xuat_chi_tiet', ['nam_hoc', 'quan_nhan_id'], unique=False)
    op.create_index('ix_dxct_de_xuat_doi_tuong', 'de_xuat_chi_tiet', ['de_xuat_id', 'doi_tuong'], unique=False)
    op.create_index('ix_dxct_de_xuat_loai_danh_hieu', 'de_xuat_chi_tiet', ['de_xuat_id', 'loai_danh_hieu'], unique=False)

    # phe_duyet / item result query hot paths
    op.create_index('ix_phe_duyet_de_xuat_ket_qua', 'phe_duyet', ['de_xuat_id', 'ket_qua'], unique=False)
    op.create_index('ix_kqdct_chi_tiet_ket_qua', 'ket_qua_duyet_chi_tiet', ['chi_tiet_id', 'ket_qua'], unique=False)

    # rewards list/reporting
    op.create_index('ix_khen_thuong_don_vi_nam_hoc', 'khen_thuong', ['don_vi_id', 'nam_hoc'], unique=False)
    op.create_index('ix_khen_thuong_loai_nam_hoc', 'khen_thuong', ['loai_danh_hieu', 'nam_hoc'], unique=False)

    # personnel list/query
    op.create_index('ix_quan_nhan_don_vi_ho_ten', 'quan_nhan', ['don_vi_id', 'ho_ten'], unique=False)


def downgrade():
    op.drop_index('ix_quan_nhan_don_vi_ho_ten', table_name='quan_nhan')
    op.drop_index('ix_khen_thuong_loai_nam_hoc', table_name='khen_thuong')
    op.drop_index('ix_khen_thuong_don_vi_nam_hoc', table_name='khen_thuong')
    op.drop_index('ix_kqdct_chi_tiet_ket_qua', table_name='ket_qua_duyet_chi_tiet')
    op.drop_index('ix_phe_duyet_de_xuat_ket_qua', table_name='phe_duyet')
    op.drop_index('ix_dxct_de_xuat_loai_danh_hieu', table_name='de_xuat_chi_tiet')
    op.drop_index('ix_dxct_de_xuat_doi_tuong', table_name='de_xuat_chi_tiet')
    op.drop_index('ix_dxct_nam_hoc_quan_nhan', table_name='de_xuat_chi_tiet')
    op.drop_index('ix_de_xuat_nam_hoc_trang_thai', table_name='de_xuat')
    op.drop_index('ix_de_xuat_trang_thai_ngay_gui', table_name='de_xuat')
    op.drop_index('ix_de_xuat_don_vi_trang_thai', table_name='de_xuat')
