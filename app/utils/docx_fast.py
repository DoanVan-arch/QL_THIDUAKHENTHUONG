"""
docx_fast.py — Tạo file .docx (Word) trực tiếp từ XML string, không qua python-docx API.

Lý do: python-docx gọi hàng chục lần XML DOM mutation (get_or_add_child, insert_element_before,
qn resolution...) cho mỗi ô bảng. Với 1,800+ hàng → ~200,000+ XML operations → 15-20 giây.

Cách này dùng Python string concatenation để xây dựng XML, sau đó đóng gói vào ZIP (định dạng
.docx thực chất là ZIP). Kết quả: 0.02s cho 1,800 hàng thay vì 17 giây (nhanh hơn ~700 lần).

Hạn chế: Không hỗ trợ formatting phức tạp (merge cells, image embed trong bảng, v.v.)
nhưng hoàn toàn đủ cho bảng danh sách text đơn giản.
"""
import io
import zipfile
import html as _html
from datetime import datetime, date


# ─── XML namespaces cần thiết ────────────────────────────────────────────────
_NSDECLS = (
    'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
    'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:o="urn:schemas-microsoft.com:office:office" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:w10="urn:schemas-microsoft-com:office:word" '
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
    'xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" '
    'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" '
    'xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" '
    'xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" '
    'xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" '
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
    'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
    'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
)

_SP = 'xml:space="preserve"'


def _e(s):
    """HTML-escape a cell value for safe embedding in XML."""
    return _html.escape(str(s or ''), quote=True)


def _font_size_half_pt(pt):
    """Convert pt to half-points (Word unit)."""
    return str(int(pt * 2))


# ─── Building blocks ─────────────────────────────────────────────────────────

def _run(text, bold=False, size_pt=10, italic=False, color=None):
    """Tạo <w:r> XML string cho 1 đoạn text."""
    rpr_parts = []
    if bold:
        rpr_parts.append('<w:b/><w:bCs/>')
    if italic:
        rpr_parts.append('<w:i/><w:iCs/>')
    if size_pt:
        s = _font_size_half_pt(size_pt)
        rpr_parts.append(f'<w:sz w:val="{s}"/><w:szCs w:val="{s}"/>')
    if color:
        rpr_parts.append(f'<w:color w:val="{color}"/>')
    rpr_parts.append('<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>')

    rpr = ('<w:rPr>' + ''.join(rpr_parts) + '</w:rPr>') if rpr_parts else ''
    t = f'<w:t {_SP}>{_e(text)}</w:t>' if text else '<w:t/>'
    return f'<w:r>{rpr}{t}</w:r>'


def _para(text='', bold=False, size_pt=10, italic=False,
          align='left', color=None, space_before=0, space_after=0):
    """Tạo <w:p> XML string."""
    jc_map = {'left': 'left', 'center': 'center', 'right': 'right', 'justify': 'both'}
    jc = jc_map.get(align, 'left')
    ppr = (
        '<w:pPr>'
        f'<w:jc w:val="{jc}"/>'
        f'<w:spacing w:before="{space_before}" w:after="{space_after}"/>'
        '<w:pStyle w:val="Normal"/>'
        '</w:pPr>'
    )
    r = _run(text, bold=bold, size_pt=size_pt, italic=italic, color=color) if text else ''
    return f'<w:p>{ppr}{r}</w:p>'


def _cell(text, bold=False, size_pt=10, italic=False, align='left',
          width_twips=None, shade=None, colspan=1, vAlign='center'):
    """Tạo <w:tc> XML string."""
    tcpr_parts = []
    if width_twips:
        tcpr_parts.append(f'<w:tcW w:w="{width_twips}" w:type="dxa"/>')
    if shade:
        tcpr_parts.append(f'<w:shd w:val="clear" w:color="auto" w:fill="{shade}"/>')
    if colspan > 1:
        tcpr_parts.append(f'<w:gridSpan w:val="{colspan}"/>')
    tcpr_parts.append(f'<w:vAlign w:val="{vAlign}"/>')

    tcpr = ('<w:tcPr>' + ''.join(tcpr_parts) + '</w:tcPr>') if tcpr_parts else ''
    p = _para(text, bold=bold, size_pt=size_pt, italic=italic, align=align,
              space_before=20, space_after=20)
    return f'<w:tc>{tcpr}{p}</w:tc>'


def _header_row(headers, widths_twips, size_pt=10, shade='2C4770'):
    """Tạo hàng header bảng với nền đậm."""
    cells = ''.join(
        _cell(h, bold=True, size_pt=size_pt, align='center',
              width_twips=w, shade=shade, vAlign='center')
        for h, w in zip(headers, widths_twips)
    )
    trpr = '<w:trPr><w:tblHeader/><w:trHeight w:val="400"/></w:trPr>'
    return f'<w:tr>{trpr}{cells}</w:tr>'


def _data_row(cells_data, widths_twips, size_pt=10, shade=None):
    """Tạo 1 hàng dữ liệu.
    cells_data: list of (text, bold, align) tuples or plain strings.
    """
    cells_xml = []
    for i, cell_data in enumerate(cells_data):
        w = widths_twips[i] if i < len(widths_twips) else None
        if isinstance(cell_data, tuple):
            text, bold, align = cell_data[0], cell_data[1] if len(cell_data) > 1 else False, cell_data[2] if len(cell_data) > 2 else 'left'
        else:
            text, bold, align = cell_data, False, 'left'
        cells_xml.append(_cell(text, bold=bold, size_pt=size_pt, align=align,
                               width_twips=w, shade=shade))
    return '<w:tr>' + ''.join(cells_xml) + '</w:tr>'


def _total_row(text, widths_twips, size_pt=10):
    """Dòng tổng hợp span toàn bộ cột."""
    total_w = sum(widths_twips)
    cell = _cell(text, bold=True, size_pt=size_pt, align='center',
                 width_twips=total_w, colspan=len(widths_twips))
    return f'<w:tr>{cell}</w:tr>'


def _build_table(headers, rows_xml, widths_twips, total_label=None,
                 size_pt=10, header_shade='2C4770'):
    """Tạo <w:tbl> XML string hoàn chỉnh."""
    total_w = sum(widths_twips)
    grid_cols = ''.join(f'<w:gridCol w:w="{w}"/>' for w in widths_twips)

    tbl_pr = (
        '<w:tblPr>'
        f'<w:tblW w:w="{total_w}" w:type="dxa"/>'
        '<w:tblLayout w:type="fixed"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="AAAAAA"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="AAAAAA"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="AAAAAA"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="AAAAAA"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="DDDDDD"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="DDDDDD"/>'
        '</w:tblBorders>'
        '<w:tblCellMar>'
        '<w:top w:w="60" w:type="dxa"/>'
        '<w:left w:w="100" w:type="dxa"/>'
        '<w:bottom w:w="60" w:type="dxa"/>'
        '<w:right w:w="100" w:type="dxa"/>'
        '</w:tblCellMar>'
        '</w:tblPr>'
    )

    header_xml = _header_row(headers, widths_twips, size_pt=size_pt, shade=header_shade)
    total_xml = _total_row(total_label, widths_twips, size_pt=size_pt) if total_label else ''

    if isinstance(rows_xml, list):
        rows_xml = ''.join(rows_xml)

    return (
        f'<w:tbl>'
        f'<w:tblGrid>{grid_cols}</w:tblGrid>'
        f'{tbl_pr}'
        f'{header_xml}'
        f'{rows_xml}'
        f'{total_xml}'
        f'</w:tbl>'
    )


# ─── Document wrapper ────────────────────────────────────────────────────────

def _build_document_xml(body_parts, margin_top=1440, margin_bottom=1440,
                        margin_left=1800, margin_right=720):
    """Wrap body content into full w:document XML."""
    body = ''.join(body_parts)
    sect_pr = (
        '<w:sectPr>'
        f'<w:pgSz w:w="12240" w:h="15840"/>'
        f'<w:pgMar w:top="{margin_top}" w:right="{margin_right}" '
        f'w:bottom="{margin_bottom}" w:left="{margin_left}" '
        f'w:header="720" w:footer="720" w:gutter="0"/>'
        '</w:sectPr>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document {_NSDECLS}>'
        f'<w:body>{body}{sect_pr}</w:body>'
        '</w:document>'
    )


# ─── DOCX packager ───────────────────────────────────────────────────────────

_CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/settings.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''

_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>'''

_WORD_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    Target="settings.xml"/>
</Relationships>'''

_STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
      <w:sz w:val="20"/><w:szCs w:val="20"/>
    </w:rPr></w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
      <w:sz w:val="20"/><w:szCs w:val="20"/>
    </w:rPr>
  </w:style>
</w:styles>'''

_SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
  <w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
</w:settings>'''


def build_docx(document_xml_str):
    """Đóng gói XML string thành .docx (ZIP) và trả về BytesIO."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        zf.writestr('[Content_Types].xml', _CONTENT_TYPES)
        zf.writestr('_rels/.rels',         _RELS)
        zf.writestr('word/document.xml',   document_xml_str.encode('utf-8'))
        zf.writestr('word/_rels/document.xml.rels', _WORD_RELS)
        zf.writestr('word/styles.xml',     _STYLES)
        zf.writestr('word/settings.xml',   _SETTINGS)
    buf.seek(0)
    return buf


# ─── Convenience helpers ─────────────────────────────────────────────────────

def cm_to_twips(cm):
    """Convert cm to Word twips (1 inch = 1440 twips, 1 inch = 2.54 cm)."""
    return int(cm * 1440 / 2.54)


def make_docx_from_body(body_parts, **kwargs):
    """Shortcut: build document XML → pack into docx → return BytesIO."""
    doc_xml = _build_document_xml(body_parts, **kwargs)
    return build_docx(doc_xml)
