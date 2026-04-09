// ==========================================
// QUẢN LÝ THI ĐUA KHEN THƯỞNG - App JS
// ==========================================

// Sidebar toggle
function toggleSidebar() {
    const sidebar = document.getElementById('sidebar');
    sidebar.classList.toggle('show');
}

// Close sidebar on mobile when clicking outside
document.addEventListener('click', function (e) {
    const sidebar = document.getElementById('sidebar');
    const toggle = document.querySelector('.sidebar-toggle');
    if (sidebar && toggle && window.innerWidth <= 768) {
        if (!sidebar.contains(e.target) && !toggle.contains(e.target)) {
            sidebar.classList.remove('show');
        }
    }
});

// Auto-dismiss flash messages after 5 seconds
document.addEventListener('DOMContentLoaded', function () {
    const alerts = document.querySelectorAll('.flash-container .alert');
    alerts.forEach(function (alert) {
        setTimeout(function () {
            const bsAlert = bootstrap.Alert.getOrCreateInstance(alert);
            bsAlert.close();
        }, 5000);
    });
});

// Conditional field visibility based on doi_tuong
const NCKH_CRITERIA = {
    'Giảng viên': 'Tham gia nghiên cứu đề tài, sáng kiến khoa học các cấp đúng tiến độ, nghiệm thu đạt khá trở lên; hoặc hướng dẫn học viên NCKH đạt giải Ba cấp hệ trở lên; hoặc biên soạn giáo trình, tài liệu; hoặc có sáng kiến ứng dụng hiệu quả',
    'Cán bộ': 'Tham gia nghiên cứu đề tài, sáng kiến khoa học các cấp đúng tiến độ, nghiệm thu đạt khá trở lên; hoặc có sáng kiến, đề xuất giải pháp mang tính đột phá',
    'Công chức quốc phòng': 'Tham gia nghiên cứu đề tài, sáng kiến khoa học các cấp đúng tiến độ, nghiệm thu đạt khá trở lên; hoặc có sáng kiến, đề xuất giải pháp mang tính đột phá',
    'Quân nhân chuyên nghiệp': 'Tham gia nghiên cứu đề tài, sáng kiến khoa học các cấp đúng tiến độ, nghiệm thu đạt khá trở lên; hoặc có sáng kiến, đề xuất giải pháp mang tính đột phá',
    'Công nhân viên': 'Tham gia nghiên cứu đề tài, sáng kiến khoa học các cấp đúng tiến độ, nghiệm thu đạt khá trở lên; hoặc có sáng kiến, đề xuất giải pháp mang tính đột phá',
    'Học viên năm thứ I': 'Tối thiểu 01 bài đăng trên Website hoặc 01 bài đăng Bản tin thi đua của Nhà trường',
    'Học viên năm thứ II': 'Tối thiểu nghiên cứu chuyên đề khoa học đạt giải Nhì cấp hệ',
    'Học viên năm thứ III': 'Tối thiểu nghiên cứu đề tài khoa học, sáng kiến đạt giải Nhì cấp hệ, tiểu đoàn trở lên hoặc chuyên đề đạt giải Ba cấp Trường trở lên',
    'Học viên năm thứ IV': 'Tối thiểu nghiên cứu đề tài khoa học, sáng kiến đạt giải Nhì cấp hệ, tiểu đoàn trở lên hoặc chuyên đề đạt giải Ba cấp Trường trở lên',
    'Học viên sau đại học': 'Tối thiểu nghiên cứu đề tài khoa học, sáng kiến đạt giải Ba cấp Trường trở lên',
    'Học viên VB2': 'Tối thiểu nghiên cứu đề tài khoa học, sáng kiến đạt giải Nhì cấp hệ, tiểu đoàn trở lên hoặc chuyên đề đạt giải Ba cấp Trường trở lên',
    'Học viên tiến sĩ': 'Tối thiểu nghiên cứu đề tài khoa học, sáng kiến đạt giải Ba cấp Trường trở lên',
    'Học viên quốc tế': 'Hoàn thành tốt nhiệm vụ NCKH theo đúng mức quy định'
};

const STUDENT_TYPES = [
    'Học viên năm thứ I', 'Học viên năm thứ II', 'Học viên năm thứ III',
    'Học viên năm thứ IV', 'Học viên sau đại học', 'Học viên VB2',
    'Học viên tiến sĩ', 'Học viên quốc tế'
];

function updateConditionalFields(doiTuong, hocVi) {
    const isLecturer = doiTuong === 'Giảng viên';
    const isStudent = STUDENT_TYPES.includes(doiTuong);
    const hasTienSi = hocVi === 'Tiến sĩ';

    // Lecturer-specific fields
    toggleFieldGroup('group-gv', isLecturer);
    toggleFieldGroup('group-danh-hieu-gv', isLecturer);
    toggleFieldGroup('group-dinh-muc', isLecturer);
    toggleFieldGroup('group-lao-dong-kh', isLecturer);
    toggleFieldGroup('group-kiem-tra-giang', isLecturer);

    // PGS progress - only for Tien si
    toggleFieldGroup('group-tien-do-pgs', hasTienSi);

    // Student-specific fields
    toggleFieldGroup('group-hv', isStudent);
    toggleFieldGroup('group-danh-hieu-hv', isStudent);
    toggleFieldGroup('group-thuc-hanh', isStudent);

    // NCKH criteria label
    const nckhLabel = document.getElementById('nckh-criteria-label');
    if (nckhLabel && doiTuong) {
        nckhLabel.textContent = NCKH_CRITERIA[doiTuong] || '';
    }

    // Toggle NCKH sections based on danh_hieu
    updateNckhSection();
}

function updateNckhSection() {
    const dhSelect = document.getElementById('loai_danh_hieu_select');
    const danhHieu = dhSelect ? dhSelect.value : '';

    // Use DB-driven mapping if available, otherwise fall back to hardcoded logic
    if (typeof DANH_HIEU_TIEU_CHI !== 'undefined' && danhHieu && DANH_HIEU_TIEU_CHI[danhHieu]) {
        const tieuChi = DANH_HIEU_TIEU_CHI[danhHieu];
        const hasFullNckh = tieuChi.includes('nckh_noi_dung');
        const hasSimpleNckh = tieuChi.includes('diem_nckh') && !hasFullNckh;
        toggleFieldGroup('nckh-section-cstd', hasFullNckh);
        toggleFieldGroup('nckh-section-cstt', hasSimpleNckh);
    } else {
        // Fallback: CSTT gets simple NCKH, others get full
        const isCSTT = danhHieu === 'Chiến sĩ tiên tiến';
        toggleFieldGroup('nckh-section-cstd', !isCSTT && danhHieu !== '');
        toggleFieldGroup('nckh-section-cstt', isCSTT);
    }
}

function onDanhHieuChange() {
    updateNckhSection();
}

function toggleFieldGroup(groupId, show) {
    const el = document.getElementById(groupId);
    if (el) {
        if (show) {
            el.classList.remove('hidden');
        } else {
            el.classList.add('hidden');
        }
    }
}

// Personnel selection in nomination form - fetch details via data attributes
function onPersonnelSelect(selectEl) {
    const option = selectEl.options[selectEl.selectedIndex];
    if (!option || !option.value) return;

    const doiTuong = option.dataset.doituong || '';
    const hocVi = option.dataset.hocvi || '';

    // Auto-select doi_tuong in the dropdown
    const dtSelect = document.getElementById('doi_tuong_select');
    if (dtSelect && doiTuong) {
        dtSelect.value = doiTuong;
    }

    updateConditionalFields(doiTuong, hocVi);
}

// When doi_tuong dropdown changes manually
function onDoiTuongSelectChange() {
    const dtSelect = document.getElementById('doi_tuong_select');
    const doiTuong = dtSelect ? dtSelect.value : '';

    // Get hoc_vi from selected personnel
    const qnSelect = document.getElementById('quan_nhan_select');
    let hocVi = '';
    if (qnSelect) {
        const option = qnSelect.options[qnSelect.selectedIndex];
        if (option && option.value) {
            hocVi = option.dataset.hocvi || '';
        }
    }

    updateConditionalFields(doiTuong, hocVi);
}

// File upload preview
function previewImage(input, previewId) {
    const preview = document.getElementById(previewId);
    if (!preview) return;

    if (input.files && input.files[0]) {
        const reader = new FileReader();
        reader.onload = function (e) {
            preview.src = e.target.result;
            preview.style.display = 'block';
        };
        reader.readAsDataURL(input.files[0]);
    }
}

// Image modal viewer
function showImageModal(src) {
    let modal = document.getElementById('imagePreviewModal');
    if (!modal) {
        modal = document.createElement('div');
        modal.id = 'imagePreviewModal';
        modal.className = 'modal fade image-preview-modal';
        modal.innerHTML = `
            <div class="modal-dialog modal-lg modal-dialog-centered">
                <div class="modal-content">
                    <div class="modal-header modal-header-navy">
                        <h6 class="modal-title">Xem ảnh minh chứng</h6>
                        <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                    </div>
                    <div class="modal-body">
                        <img id="modalPreviewImg" src="" alt="Preview">
                    </div>
                </div>
            </div>`;
        document.body.appendChild(modal);
    }
    document.getElementById('modalPreviewImg').src = src;
    const bsModal = new bootstrap.Modal(modal);
    bsModal.show();
}

// Confirm delete
function confirmDelete(formId, itemName) {
    if (confirm('Bạn có chắc chắn muốn xóa ' + (itemName || 'mục này') + '?')) {
        document.getElementById(formId).submit();
    }
}
