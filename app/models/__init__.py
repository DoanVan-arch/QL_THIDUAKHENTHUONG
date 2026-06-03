from app.models.user import User, Role
from app.models.unit import DonVi, LoaiDonVi
from app.models.personnel import QuanNhan, CapBac, HocHam, HocVi, DoiTuong, MucDoHoanThanh
from app.models.certificate import ChungChi, LoaiChungChi
from app.models.nomination import DeXuat, DeXuatChiTiet, MinhChung, LoaiDanhHieu, TrangThaiDeXuat, DanhHieu, TieuChi
from app.models.approval import PheDuyet, PhongDuyet, KetQuaDuyet, KetQuaDuyetChiTiet
from app.models.reward import KhenThuong
from app.models.notification import ThongBao
from app.models.catalog import ChucVuOption, CapBacOption, DoiTuongOption
from app.models.evaluation import NhomTieuChi, DanhGiaHangNam, DiemQuyDinhDanhHieu
from app.models.hoi_dong import HoiDongBieuQuyet
from app.models.transfer import ChuyenDonVi, TrangThaiChuyen
from app.models.edit_request import YeuCauChinhSua, TrangThaiYeuCauSua