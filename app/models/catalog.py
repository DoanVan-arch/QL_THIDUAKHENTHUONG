from sqlalchemy import Column, Integer, String, Boolean, DateTime, func
from app.extensions import db


class ChucVuOption(db.Model):
    __tablename__ = 'chuc_vu_option'

    id = Column(Integer, primary_key=True, autoincrement=True)
    ten = Column(String(120), nullable=False, unique=True)
    thu_tu = Column(Integer, default=0)
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=func.now())
    updated_at = Column(DateTime, default=func.now(), onupdate=func.now())


class CapBacOption(db.Model):
    __tablename__ = 'cap_bac_option'

    id = Column(Integer, primary_key=True, autoincrement=True)
    ten = Column(String(120), nullable=False, unique=True)
    thu_tu = Column(Integer, default=0)
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=func.now())
    updated_at = Column(DateTime, default=func.now(), onupdate=func.now())
