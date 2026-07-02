"""
Replace the python-docx building section of export_tracking_word
(lines 2624-3179, 0-indexed: 2623-3178) with the fast XML approach.
"""
NEW_CODE = '''
    # ── Tạo docx nhanh bằng XML template (thay python-docx ~700x nhanh hơn) ──────
    from app.utils.docx_fast import (
        cm_to_twips, _para, _build_table, _data_row,
        _build_document_xml, build_docx,
    )
    import datetime as _dt

    SP = \'xml:space="preserve"\'

    def _section_heading(text):
        return _para(text, bold=True, size_pt=12, space_before=120, space_after=40)

    def _empty_notice():
        return _para(\'(Không có)\', italic=True, size_pt=10, space_before=40, space_after=40)

    # ── Chiều rộng cột cá nhân: STT|Họ tên|Cấp bậc|Chức vụ|Đơn vị|Năm học|Tóm tắt ──
    CN_WIDTHS = [
        cm_to_twips(0.8),   # STT
        cm_to_twips(3.5),   # Họ tên
        cm_to_twips(2.0),   # Cấp bậc
        cm_to_twips(2.5),   # Chức vụ
        cm_to_twips(2.5),   # Đơn vị
        cm_to_twips(1.5),   # Năm học
        cm_to_twips(4.0),   # Tóm tắt
    ]
    CN_HEADERS = [\'STT\', \'Họ và tên\', \'Cấp bậc\', \'Chức vụ\', \'Đơn vị\', \'Năm học\', \'Tóm tắt thành tích\']

    # ── Chiều rộng cột tập thể: STT|Tên đơn vị|Đề xuất đơn vị|Năm học ──
    TT_WIDTHS = [
        cm_to_twips(0.8),
        cm_to_twips(4.5),
        cm_to_twips(9.5),
        cm_to_twips(2.0),
    ]
    TT_HEADERS = [\'STT\', \'Tên đơn vị\', \'Đề xuất của đơn vị\', \'Năm học\']

    # ── TieuChi map (batch load 1 lần) ───────────────────────────────────────────
    from app.models.nomination import TieuChi as _TieuChi
    _tieu_chi_map_all = {tc.ma_truong: tc.ten for tc in _TieuChi.query.all()}

    def _build_tomtat(ct):
        """Tóm tắt thành tích ngắn gọn cho cột Word."""
        parts = []
        if ct.muc_do_hoan_thanh:
            parts.append(ct.muc_do_hoan_thanh)
        if ct.diem_tong_ket:
            parts.append(f\'ĐTK: {ct.diem_tong_ket}\')
        if ct.ket_qua_ren_luyen:
            parts.append(f\'RL: {ct.ket_qua_ren_luyen}\')
        if ct.xep_loai_dang_vien:
            parts.append(f\'ĐV: {ct.xep_loai_dang_vien}\')
        if ct.thanh_tich_ca_nhan_khac:
            parts.append(ct.thanh_tich_ca_nhan_khac[:80])
        return \'; \'.join(parts)

    def _build_tt_criteria(ct):
        """Tiêu chí tập thể dạng text cho cột Word."""
        td = ct.tap_the_dict or {}
        lines = []
        for k, v in td.items():
            if v and str(v).strip() not in (\'\', \'0\', \'None\'):
                label = _tieu_chi_map_all.get(k, k)
                lines.append(f\'- {label}: {v}\')
        if ct.muc_do_hoan_thanh:
            lines.insert(0, f\'- Mức độ HT: {ct.muc_do_hoan_thanh}\')
        return \'\\n\'.join(lines)

    def _cn_rows(items):
        rows_xml = []
        for i, (ct, dx) in enumerate(items, 1):
            qn = ct.quan_nhan
            row_cells = [
                (str(i), False, \'center\'),
                (qn.ho_ten if qn else \'\', True, \'left\'),
                (qn.cap_bac if qn else \'\', False, \'left\'),
                (qn.chuc_vu if qn else \'\', False, \'left\'),
                (dx.don_vi.ten_don_vi if dx.don_vi else \'\', False, \'left\'),
                (dx.nam_hoc or \'\', False, \'center\'),
                (_build_tomtat(ct), False, \'left\'),
            ]
            shade = \'F8F9FA\' if i % 2 == 0 else None
            rows_xml.append(_data_row(row_cells, CN_WIDTHS, size_pt=9, shade=shade))
        return rows_xml

    def _tt_rows(items):
        rows_xml = []
        for i, (ct, dx) in enumerate(items, 1):
            criteria_text = _build_tt_criteria(ct)
            row_cells = [
                (str(i), False, \'center\'),
                (ct.ten_don_vi_de_xuat or (dx.don_vi.ten_don_vi if dx.don_vi else \'\'), True, \'left\'),
                (criteria_text, False, \'left\'),
                (dx.nam_hoc or \'\', False, \'center\'),
            ]
            shade = \'F8F9FA\' if i % 2 == 0 else None
            rows_xml.append(_data_row(row_cells, TT_WIDTHS, size_pt=9, shade=shade))
        return rows_xml

    # ── Xây dựng nội dung tài liệu ───────────────────────────────────────────────
    body = []

    # Tiêu đề
    today_str = _dt.date.today().strftime(\'%d/%m/%Y\')
    body.append(_para(\'TRƯỜNG SĨ QUAN CHÍNH TRỊ\', bold=True, size_pt=12, align=\'center\', space_before=0, space_after=20))
    body.append(_para(f\'THEO DÕI PHÊ DUYỆT KHEN THƯỞNG — {title_nam_hoc}\', bold=True, size_pt=14, align=\'center\', space_before=60, space_after=20))
    body.append(_para(f\'(Xuất lúc {_dt.datetime.now().strftime("%H:%M")} ngày {today_str})\', italic=True, size_pt=10, align=\'center\', space_before=0, space_after=120))

    def _add_section(label, items, is_tap_the=False):
        if not items:
            return
        body.append(_section_heading(label))
        if is_tap_the:
            rows_xml = _tt_rows(items)
            total = f\'Tổng cộng: {len(items)} đơn vị\'
            body.append(_build_table(TT_HEADERS, rows_xml, TT_WIDTHS, total_label=total, size_pt=9))
        else:
            rows_xml = _cn_rows(items)
            total = f\'Tổng cộng: {len(items)} người\'
            body.append(_build_table(CN_HEADERS, rows_xml, CN_WIDTHS, total_label=total, size_pt=9))
        body.append(_para(\'\', space_before=60, space_after=0))

    _add_section(\'I. DANH HIỆU ĐƠN VỊ QUYẾT THẮNG\', ds_quyet_thang, is_tap_the=True)
    _add_section(\'II. DANH HIỆU ĐƠN VỊ TIÊN TIẾN\', ds_tien_tien_dv, is_tap_the=True)
    _add_section(\'III. CHIẾN SĨ THI ĐUA\', ds_chien_si_tdcs)
    _add_section(\'IV. CHIẾN SĨ TIÊN TIẾN\', ds_chien_si_tt)
    for extra_dh, extra_items in ds_khac.items():
        _add_section(extra_dh, extra_items)

    # Footer
    body.append(_para(f\'(Xuất lúc {_dt.datetime.now().strftime("%H:%M ngày %d/%m/%Y")})\',
                      italic=True, size_pt=9, align=\'right\', space_before=120, space_after=0))

    doc_xml = _build_document_xml(body, margin_left=2016, margin_right=720, margin_top=1440, margin_bottom=1440)
    buf = build_docx(doc_xml)

    fname_parts = [\'TheoDoiPheduyet\']
    if nam_hoc_filter:
        fname_parts.append(nam_hoc_filter.replace(\'-\', \'_\'))
    fname_parts.append(_dt.datetime.now().strftime(\'%d%m%Y\'))
    filename = \'_\'.join(fname_parts) + \'.docx\'

    response = send_file(
        buf, as_attachment=True, download_name=filename,
        mimetype=\'application/vnd.openxmlformats-officedocument.wordprocessingml.document\'
    )
    response.set_cookie(
        \'export_done\', \'1\',
        max_age=600, httponly=False, samesite=\'Lax\', path=\'/\'
    )
    return response
'''

with open('app/routes/admin.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Replace lines 2624-3179 (0-indexed: 2623 to 3178 inclusive)
start_idx = 2623  # 0-indexed: line 2624
end_idx   = 3179  # 0-indexed: line 3179 (return response) inclusive

new_lines = lines[:start_idx] + [NEW_CODE + '\n'] + lines[end_idx:]

with open('app/routes/admin.py', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print('Done. New total lines:', len(new_lines))
print('Replaced lines', start_idx+1, 'to', end_idx)
