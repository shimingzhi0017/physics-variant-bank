"""
高中物理变式题库一站式确定性预检、双卷契约审计、编译后处理与七重自动化终检库。

设计目标：无论输入哪套试卷/讲义/错题集、跨哪个物理模块，均通过确定性门禁生成完全一致的 A4 印刷级 DOCX 与 PDF。

一键命令行用法：
  # 阶段一独立预检 + 学生/教师双卷 1:1 契约联检（仅需 Python 标准库，零第三方依赖）
  python3 scripts/build_docx.py precheck student_bank.md teacher_bank.md

  # 阶段二一键编译 DOCX/PDF 并执行七重终检（使用 ensure_env.py 配置的专用虚拟环境）
  ~/.local/share/physics-variant-bank-venv/bin/python scripts/build_docx.py compile \
      student_bank.md 输出_学生版.docx --expected-imgs 10 --pdf --preview-dir preview_s
"""
import argparse
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from io import BytesIO

EMU_PER_INCH = 914400.0
EMU_PER_CM = 360000.0
M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

STUDENT_FORBIDDEN_MARKERS = (
    '【参考答案】', '【分步解析】', '【防错点拨】', '【答案】', '【解析】', '【详解】', '（原第',
)
UNSUPPORTED_PANDOC_TEX = (
    (r'\\cfrac\b', r'Pandoc 不支持 \cfrac，请改用 \frac 或 \dfrac'),
    (r'\\xlongequal\b', r'Pandoc 不支持 \xlongequal，请改用 \overset{...}{=}'),
    (r'\\degree\b', r'Pandoc 不支持 \degree，请改用 ^\circ'),
    (r'(?<![a-zA-Z\\])\\rm\b', r'Pandoc 不支持旧式 \rm，请改用 \mathrm{...}'),
)


def _require_docx_deps():
    """延迟加载 DOCX/图像依赖，确保 precheck 子命令在纯标准库环境下可直接运行。"""
    global Image, etree, Document, Cm, Pt, RGBColor
    global WD_ALIGN_PARAGRAPH, WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
    global OxmlElement, parse_xml, qn, nsdecls
    if 'Document' in globals():
        return
    try:
        from PIL import Image as _Image
        from lxml import etree as _etree
        from docx import Document as _Document
        from docx.shared import Cm as _Cm, Pt as _Pt, RGBColor as _RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH as _WD_ALIGN_PARAGRAPH
        from docx.enum.table import WD_TABLE_ALIGNMENT as _WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL as _WD_ALIGN_VERTICAL
        from docx.oxml import OxmlElement as _OxmlElement, parse_xml as _parse_xml
        from docx.oxml.ns import qn as _qn, nsdecls as _nsdecls
    except ImportError as exc:
        raise RuntimeError(
            f"缺少编译依赖 ({exc})。请先运行 `python3 scripts/ensure_env.py` 自动配置环境，"
            "再使用 ~/.local/share/physics-variant-bank-venv/bin/python 执行编译。"
        ) from exc
    Image, etree, Document = _Image, _etree, _Document
    Cm, Pt, RGBColor = _Cm, _Pt, _RGBColor
    WD_ALIGN_PARAGRAPH = _WD_ALIGN_PARAGRAPH
    WD_TABLE_ALIGNMENT = _WD_TABLE_ALIGNMENT
    WD_ALIGN_VERTICAL = _WD_ALIGN_VERTICAL
    OxmlElement, parse_xml, qn, nsdecls = _OxmlElement, _parse_xml, _qn, _nsdecls


def _set_run_fonts(style, east_asia, ascii_font='Times New Roman'):
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), east_asia)
    rFonts.set(qn('w:ascii'), ascii_font)
    rFonts.set(qn('w:hAnsi'), ascii_font)


def build_reference_docx(path='reference.docx', base_docx=None):
    """
    生成标准 A4 打印样式 Word 模板（宋体 + Times New Roman 10.5 pt，黑体标题）。
    自动以 `pandoc --print-default-data-file reference.docx` 为基底，完整保留 Pandoc
    内置的 Table / Compact / FirstParagraph 样式，彻底防止 LibreOffice 表格坍塌错位。
    """
    _require_docx_deps()
    tmp_base = None
    if base_docx is None and shutil.which('pandoc'):
        fd, tmp_base = tempfile.mkstemp(suffix='.docx')
        os.close(fd)
        raw = subprocess.check_output(['pandoc', '--print-default-data-file', 'reference.docx'])
        with open(tmp_base, 'wb') as f:
            f.write(raw)
        base_docx = tmp_base

    try:
        doc = Document(base_docx) if (base_docx and os.path.exists(base_docx)) else Document()
        for sec in doc.sections:
            sec.page_width = Cm(21.0)
            sec.page_height = Cm(29.7)
            sec.top_margin = Cm(2.0)
            sec.bottom_margin = Cm(2.0)
            sec.left_margin = Cm(2.2)
            sec.right_margin = Cm(2.2)

        by_lower = {s.name.lower(): s for s in doc.styles if getattr(s, 'name', None)}

        for sname, sz, ls, sa in [
            ('normal', 10.5, 1.25, 4),
            ('body text', 10.5, 1.25, 4),
            ('first paragraph', 10.5, 1.25, 4),
            ('block text', 10.0, 1.20, 4),
            ('compact', 9.5, 1.15, 2),
            ('captioned figure', 10.5, 1.20, 4),
            ('image caption', 9.5, 1.15, 4),
        ]:
            st = by_lower.get(sname)
            if st is None:
                continue
            st.font.name = 'Times New Roman'
            st.font.size = Pt(sz)
            st.font.italic = False
            st.font.color.rgb = RGBColor(0, 0, 0)
            _set_run_fonts(st, '宋体')
            st.paragraph_format.line_spacing = ls
            st.paragraph_format.space_after = Pt(sa)

        for level, size_pt in [(1, 15), (2, 12.5), (3, 11)]:
            st = by_lower.get(f'heading {level}')
            if st is None:
                continue
            st.font.name = 'Times New Roman'
            st.font.size = Pt(size_pt)
            st.font.bold = True
            st.font.color.rgb = RGBColor(0, 0, 0)
            _set_run_fonts(st, '黑体')
            st.paragraph_format.space_before = Pt(8)
            st.paragraph_format.space_after = Pt(4)
            st.paragraph_format.keep_with_next = True

        doc.save(path)
    finally:
        if tmp_base and os.path.exists(tmp_base):
            os.remove(tmp_base)
    return path


def set_table_full_grid(table):
    """为 Word 表格显式写入黑色细实线全框线、单元格内边距、跨页重复表头及防劈裂属性。"""
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = table._tbl.tblPr
    for b in tblPr.findall(qn('w:tblBorders')):
        tblPr.remove(b)
    borders_xml = (
        f'<w:tblBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        '<w:insideH w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        '<w:insideV w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        '</w:tblBorders>'
    )
    tblPr.append(parse_xml(borders_xml))

    for cm in tblPr.findall(qn('w:tblCellMar')):
        tblPr.remove(cm)
    cell_mar_xml = (
        f'<w:tblCellMar {nsdecls("w")}>'
        '<w:top w:w="70" w:type="dxa"/>'
        '<w:bottom w:w="70" w:type="dxa"/>'
        '<w:left w:w="100" w:type="dxa"/>'
        '<w:right w:w="100" w:type="dxa"/>'
        '</w:tblCellMar>'
    )
    tblPr.append(parse_xml(cell_mar_xml))

    for r_idx, row in enumerate(table.rows):
        trPr = row._tr.get_or_add_trPr()
        if trPr.find(qn('w:cantSplit')) is None:
            trPr.append(OxmlElement('w:cantSplit'))
        if r_idx == 0 and trPr.find(qn('w:tblHeader')) is None:
            trPr.append(OxmlElement('w:tblHeader'))
        for cell in row.cells:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            tcPr = cell._tc.get_or_add_tcPr()
            for shd in tcPr.findall(qn('w:shd')):
                tcPr.remove(shd)
            if r_idx == 0:
                tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="F2F2F2"/>'))
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for r in p.runs:
                        r.bold = True


def _is_display_math_or_drawing(p):
    has_draw = bool(p._p.xpath('.//*[local-name()="drawing"]'))
    has_omath_para = bool(p._p.xpath('.//*[local-name()="oMathPara"]'))
    return has_draw or has_omath_para


def _mrun_text(elem):
    return ''.join(t.text or '' for t in elem.findall(f'{{{M_NS}}}t'))


def _is_plain_alpha_mrun(elem):
    """判断是否为带 m:sty='p' 且含字母/汉字的 OMML 字符 run。"""
    if elem.tag != f'{{{M_NS}}}r':
        return False
    rpr = elem.find(f'{{{M_NS}}}rPr')
    if rpr is None:
        return False
    sty = rpr.find(f'{{{M_NS}}}sty')
    if sty is None or sty.get(f'{{{M_NS}}}val') != 'p':
        return False
    txt = _mrun_text(elem)
    return bool(txt) and any(ch.isalpha() for ch in txt)


def fix_omml_upright(doc):
    """
    修复 Pandoc 导出 OMML 正体符号与核反应前置上下标在 LibreOffice / WPS 中的跨平台缺陷：
      1. 对独立大写希腊算子/单位 Δ (\\Delta) 与 Ω (\\Omega) 自动补齐正体 <m:sty m:val="p"/> + <m:nor/>；
      2. 对所有带 <m:sty m:val="p"/> 且含字母/汉字的 <m:r>，按 ECMA-376 XSD 顺序在 <m:rPr> 首位补齐 <m:nor/>；
      3. 仅合并由 \\mathrm{...} 拆出的单字符正体 <m:r>（如 c+m -> cm, B+a -> Ba, μ+s -> μs），
         绝不将多字符数学算子（如 sin, cos, ln, lim）与后续正体点名（如 A, B）误合并为 'sinA'；
      4. 将 Pandoc 用 U+200B 零宽空格模拟的前置上下标（如 {}_{92}^{235}\\mathrm{U}）转为原生 ECMA-376
         <m:sPre>，彻底消除 LibreOffice 渲染出的 '□' 缺字方框。
    """
    _require_docx_deps()
    m_r_tag = f'{{{M_NS}}}r'
    m_rpr_tag = f'{{{M_NS}}}rPr'
    m_sty_tag = f'{{{M_NS}}}sty'
    m_nor_tag = f'{{{M_NS}}}nor'
    m_val_attr = f'{{{M_NS}}}val'
    m_t_tag = f'{{{M_NS}}}t'

    fixed = 0
    single_char_runs = set()

    for r in doc.element.iter(m_r_tag):
        txt = _mrun_text(r)
        if txt in ('Δ', 'Ω'):
            rpr = r.find(m_rpr_tag)
            if rpr is None:
                rpr = etree.Element(m_rpr_tag)
                r.insert(0, rpr)
            sty = rpr.find(m_sty_tag)
            if sty is None:
                sty = etree.SubElement(rpr, m_sty_tag)
            sty.set(m_val_attr, 'p')

        if _is_plain_alpha_mrun(r):
            rpr = r.find(m_rpr_tag)
            if rpr.find(m_nor_tag) is None:
                rpr.insert(0, etree.Element(m_nor_tag))
                fixed += 1
            if len(txt) == 1 and txt.isalpha() and r.find(qn('w:rPr')) is None:
                single_char_runs.add(r)

    parents = set()
    for r in single_char_runs:
        p = r.getparent()
        if p is not None:
            parents.add(p)

    for parent in parents:
        children = list(parent)
        i = 0
        while i < len(children) - 1:
            cur = children[i]
            nxt = children[i + 1]
            if cur in single_char_runs and nxt in single_char_runs:
                cur_ts = cur.findall(m_t_tag)
                nxt_txt = _mrun_text(nxt)
                if cur_ts and nxt_txt:
                    cur_ts[-1].text = (cur_ts[-1].text or '') + nxt_txt
                    parent.remove(nxt)
                    single_char_runs.discard(nxt)
                    children.pop(i + 1)
                    continue
            i += 1

    # 转换 U+200B 空基底 sSubSup 为原生 ECMA-376 <m:sPre>（核物理同位素符号）
    s_subsup_tag = f'{{{M_NS}}}sSubSup'
    s_pre_tag = f'{{{M_NS}}}sPre'
    m_e_tag = f'{{{M_NS}}}e'
    m_sub_tag = f'{{{M_NS}}}sub'
    m_sup_tag = f'{{{M_NS}}}sup'
    allowed_next = {m_r_tag, f'{{{M_NS}}}sSup', f'{{{M_NS}}}sSub'}

    for sss in list(doc.element.iter(s_subsup_tag)):
        e_el = sss.find(m_e_tag)
        if e_el is None:
            continue
        e_txt = ''.join(t.text or '' for t in e_el.iter(m_t_tag))
        if e_txt == '​':
            parent = sss.getparent()
            if parent is None:
                continue
            siblings = list(parent)
            idx = siblings.index(sss)
            if idx + 1 < len(siblings) and siblings[idx + 1].tag in allowed_next:
                nxt_el = siblings[idx + 1]
                parent.remove(nxt_el)
                spre = etree.Element(s_pre_tag)
                sub_el = sss.find(m_sub_tag)
                sup_el = sss.find(m_sup_tag)
                if sub_el is not None:
                    spre.append(sub_el)
                if sup_el is not None:
                    spre.append(sup_el)
                new_e = etree.SubElement(spre, m_e_tag)
                new_e.append(nxt_el)
                parent.replace(sss, spre)

    # 清除任何残留的 U+200B 零宽空格，防止 LibreOffice 数学字体渲染出 '□'
    for t_el in doc.element.iter(m_t_tag):
        if t_el.text and '​' in t_el.text:
            t_el.text = t_el.text.replace('​', '')

    return fixed


def postprocess_docx(path):
    """后处理：卷头居中、含图段落居中、引导句与下方公式/插图同页绑定、全部表格全框线、OMML正体锁死。"""
    _require_docx_deps()
    doc = Document(path)
    fix_omml_upright(doc)
    if doc.paragraphs:
        doc.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    header_keys = ('考试时间', '建议用时', '满分', '物理试卷', '学生版', '教师版')
    for idx in range(1, min(4, len(doc.paragraphs))):
        p = doc.paragraphs[idx]
        sname = (p.style.name or '').lower() if p.style else ''
        if sname.startswith('heading 2'):
            break
        t = p.text.strip()
        if t and any(k in t for k in header_keys):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    paras = doc.paragraphs
    for i, p in enumerate(paras):
        if p._p.xpath('.//*[local-name()="drawing"]'):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(6)
        sname = (p.style.name or '').lower() if p.style else ''
        if sname.startswith('heading'):
            p.paragraph_format.keep_with_next = True
        elif i + 1 < len(paras) and _is_display_math_or_drawing(paras[i + 1]):
            txt = p.text.strip()
            if txt.endswith(('：', ':')) or '【参考答案】' in txt:
                p.paragraph_format.keep_with_next = True

    for tbl in doc.tables:
        set_table_full_grid(tbl)
    doc.save(path)


def verify_docx(docx_path, expected_imgs=None, max_ar_err_pct=0.5,
                min_print_scale=0.65, max_print_scale=1.15, expected_dpi=300.0):
    """
    七重自动化交付终检：
      1. 全部节的页面尺寸锁定为标准 A4（21.0 x 29.7 cm）；
      2. 正文及表格中 0 残留 LaTeX '$' 源码；
      3. 无误判代码块样式 (SourceCode / VerbatimChar)；
      4. OMML 公式末尾无孤立 '='，且正体符号均含 <m:nor/>；
      5. 含表格文档必含 Pandoc 'Table' 样式、跨页重复表头与防劈裂属性；
      6. 嵌图总数准确，且全部图片宽高比误差 < max_ar_err_pct；
      7. 全部图片相对 300 DPI 原始尺寸的印刷缩放比落入 [min_print_scale, max_print_scale]。
    """
    _require_docx_deps()
    doc = Document(docx_path)

    for s_idx, sec in enumerate(doc.sections, 1):
        w_cm = sec.page_width / EMU_PER_CM
        h_cm = sec.page_height / EMU_PER_CM
        if abs(w_cm - 21.0) > 0.2 or abs(h_cm - 29.7) > 0.2:
            raise AssertionError(
                f"[FAIL] {docx_path} 第 {s_idx} 节尺寸为 {w_cm:.2f}x{h_cm:.2f} cm，非标准 A4 (21.0x29.7 cm)！"
            )

    all_text = "\n".join(p.text for p in doc.paragraphs)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                all_text += "\n" + cell.text
    raw_dollars = all_text.count('$')
    if raw_dollars > 0:
        raise AssertionError(f"[FAIL] {docx_path}: 发现 {raw_dollars} 处残留 '$' 源码未转为公式！")

    xml_str = doc.element.xml
    if 'SourceCode' in xml_str or 'VerbatimChar' in xml_str:
        raise AssertionError(f"[FAIL] {docx_path}: 检测到误判代码块样式，请检查 Markdown 是否存在 4 空格缩进！")

    omaths = doc.element.xpath('.//*[local-name()="oMath"]')
    trailing_eq = 0
    for om in omaths:
        texts = [t.text for t in om.xpath('.//*[local-name()="t"]') if t.text]
        if ''.join(texts).strip().endswith('='):
            trailing_eq += 1
    if trailing_eq > 0:
        raise AssertionError(
            f"[FAIL] {docx_path}: 发现 {trailing_eq} 处 OMML 公式以 '=' 结尾，LibreOffice 会渲染为 '¿'！请将填空等号移到公式外。"
        )
    unfixed_rm = 0
    for r in doc.element.iter(f'{{{M_NS}}}r'):
        if _is_plain_alpha_mrun(r):
            rpr = r.find(f'{{{M_NS}}}rPr')
            if rpr is not None and rpr.find(f'{{{M_NS}}}nor') is None:
                unfixed_rm += 1
    if unfixed_rm > 0:
        raise AssertionError(
            f"[FAIL] {docx_path}: 发现 {unfixed_rm} 处 OMML 正体字母/单位缺失 <m:nor/>，LibreOffice 会退化为斜体！"
        )

    tbl_nosplit_str = 'N/A'
    if len(doc.tables) > 0:
        with zipfile.ZipFile(docx_path) as zf:
            styles_xml = zf.read('word/styles.xml').decode('utf-8', errors='ignore')
        if 'styleId="Table"' not in styles_xml:
            raise AssertionError(
                f"[FAIL] {docx_path}: 含表格但 word/styles.xml 缺失 Pandoc 'Table' 样式，LibreOffice 导出 PDF 会表格坍塌！"
            )
        for t_idx, tbl in enumerate(doc.tables, 1):
            for r_idx, row in enumerate(tbl.rows):
                trPr = row._tr.find(qn('w:trPr'))
                if trPr is None or trPr.find(qn('w:cantSplit')) is None:
                    raise AssertionError(
                        f"[FAIL] {docx_path} 表 {t_idx} 第 {r_idx + 1} 行缺失 <w:cantSplit/>，跨页会被劈裂！"
                    )
                if r_idx == 0 and trPr.find(qn('w:tblHeader')) is None:
                    raise AssertionError(
                        f"[FAIL] {docx_path} 表 {t_idx} 首行缺失 <w:tblHeader/>，跨页无法重复表头！"
                    )
        tbl_nosplit_str = 'OK'

    inlines = doc.element.xpath('.//*[local-name()="inline"]')
    if expected_imgs is not None and len(inlines) != expected_imgs:
        raise AssertionError(f"[FAIL] {docx_path}: 期望嵌入 {expected_imgs} 张图片，实际仅 {len(inlines)} 张！")

    r_embed_key = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed'
    max_err = 0.0
    min_sc, max_sc = 999.0, 0.0
    for idx, inline in enumerate(inlines, 1):
        ext_nodes = inline.xpath('./*[local-name()="extent"]')
        blip_nodes = inline.xpath('.//*[local-name()="blip"]')
        if not ext_nodes or not blip_nodes:
            continue
        cx = int(ext_nodes[0].attrib['cx'])
        cy = int(ext_nodes[0].attrib['cy'])
        rid = blip_nodes[0].attrib[r_embed_key]
        img_part = doc.part.related_parts[rid]
        with Image.open(BytesIO(img_part.blob)) as im:
            w, h = im.size
        err = abs((cx / cy) - (w / h)) / (w / h) * 100
        max_err = max(max_err, err)
        if err >= max_ar_err_pct:
            raise AssertionError(
                f"[FAIL] {docx_path} 图 {idx}: 宽高比误差 {err:.3f}% >= {max_ar_err_pct}%"
            )
        orig_cx = (w / expected_dpi) * EMU_PER_INCH
        scale = cx / orig_cx
        min_sc = min(min_sc, scale)
        max_sc = max(max_sc, scale)
        if not (min_print_scale <= scale <= max_print_scale):
            raise AssertionError(
                f"[FAIL] {docx_path} 图 {idx}: 印刷缩放比 {scale:.3f} 超出安全区间 "
                f"[{min_print_scale}, {max_print_scale}]，请对照 drawing-and-qa.md 固定预设调整嵌宽！"
            )

    sc_info = f"{min_sc:.2f}~{max_sc:.2f}" if inlines else "N/A"
    print(
        f"[PASS] {docx_path} | A4 | imgs={len(inlines)} (max_ar_err={max_err:.4f}%, scale={sc_info}) "
        f"| omml={len(omaths)} | tables={len(doc.tables)} | raw_$=0 | trailing_eq=0 | upright_rm=OK | tbl_nosplit={tbl_nosplit_str}"
    )
    return {
        'imgs': len(inlines),
        'max_ar_err': max_err,
        'scale_range': (min_sc, max_sc) if inlines else None,
        'omml': len(omaths),
        'tables': len(doc.tables),
        'tbl_nosplit': tbl_nosplit_str,
    }


def _split_pipe_row_math_safe(line):
    """按 Markdown 表格分隔符 '|' 切分单元格，忽略 $...$ 公式内部的绝对值竖线与转义 \\|。"""
    s = line.strip()
    if s.startswith('|'):
        s = s[1:]
    if s.endswith('|') and not s.endswith('\\|'):
        s = s[:-1]
    cells, cur = [], []
    in_math = False
    i = 0
    while i < len(s):
        ch = s[i]
        if ch == '\\' and i + 1 < len(s):
            cur.extend([ch, s[i + 1]])
            i += 2
            continue
        if ch == '$':
            if i + 1 < len(s) and s[i + 1] == '$':
                cur.append('$$')
                in_math = not in_math
                i += 2
                continue
            in_math = not in_math
            cur.append(ch)
            i += 1
            continue
        if ch == '|' and not in_math:
            cells.append(''.join(cur).strip())
            cur = []
        else:
            cur.append(ch)
        i += 1
    cells.append(''.join(cur).strip())
    return cells


def precheck_markdown(md_path, require_resolved_figs=False):
    """
    编译前 Markdown 确定性预检（纯标准库，跳过围栏代码块）：
      1. 无序列表项前缺空行粘连；
      2. 选项格式：4 空格缩进、同行并排多个选项、相邻选项间缺空行粘连；
      3. 公式安全：行内公式以 '=' 结尾、公式内塞入 \\underline 填空线、Pandoc 不支持的 LaTeX 宏；
      4. 嵌图语法：必须指定 {width=XXcm} 且禁写 height=；阶段二编译时拦截未替换的 <!-- 待配图 占位注释及缺失图片。
    """
    md_dir = os.path.dirname(os.path.abspath(md_path))
    with open(md_path, 'r', encoding='utf-8') as f:
        lines = f.read().splitlines()

    errors = []
    in_fence = False
    in_ul = False
    prev_was_option = False

    for i, line in enumerate(lines):
        lineno = i + 1
        s = line.lstrip()
        if s.startswith('```'):
            in_fence = not in_fence
            in_ul = False
            prev_was_option = False
            continue
        if in_fence:
            continue

        if require_resolved_figs and '<!--' in line and '待配图' in line:
            errors.append(f"第 {lineno} 行：仍残留阶段一占位注释 `{line.strip()}`，请先生成配图并替换为嵌图语法")

        if not s:
            in_ul = False
            prev_was_option = False
            continue

        # 剥离 HTML 注释与行内反引号代码后再做语法扫描
        clean_line = re.sub(r'<!--.*?-->', '', line)
        clean_no_code = re.sub(r'`[^`]*`', '', clean_line)

        # 1. 无序列表项前缺空行
        is_hr = bool(re.match(r'^([-*_])\s*(\1\s*){2,}$', s.strip()))
        is_ul = (s.startswith('- ') or s.startswith('* ')) and not is_hr
        if is_ul:
            if i > 0 and not in_ul:
                prev_s = lines[i - 1].strip()
                if (
                    prev_s
                    and not prev_s.startswith('```')
                    and not re.match(r'^([-*+]|\d+\.)\s', prev_s)
                    and not prev_s.startswith(('#', '|', '>'))
                ):
                    errors.append(
                        f"第 {lineno} 行：无序列表项前缺少空行，Pandoc 会将其并入上一段变成行内短横线"
                    )
            in_ul = True
        elif s.startswith(('#', '|', '>')) or is_hr or re.match(r'^\d+\.\s', s):
            in_ul = False

        # 2. 选择题选项格式检查
        if re.match(r'^ {4,}[A-E][.．、]', line):
            errors.append(f"第 {lineno} 行：选项存在 4 空格以上缩进，Pandoc 会误判为代码块，请改为顶格")
        is_opt = bool(re.match(r'^[A-E]\.\s+', line))
        if re.match(r'^[A-E][．、]', line):
            errors.append(f"第 {lineno} 行：选项标点请统一使用半角英文点号加空格（如 `A. `），勿用 `A．` 或 `A、`")
        if is_opt:
            if prev_was_option:
                errors.append(f"第 {lineno} 行：相邻选项之间缺少空行，Pandoc 会将多个选项合并到同一段落")
            if re.search(r'\s{2,}[B-E][.．、]\s+', clean_no_code):
                errors.append(f"第 {lineno} 行：检测到同行并排多个选项，请改为每个选项独立成段并空行分隔")
        prev_was_option = is_opt

        # 3. 公式语法与 Pandoc 禁宏检查
        for pat, msg in UNSUPPORTED_PANDOC_TEX:
            if re.search(pat, clean_no_code):
                errors.append(f"第 {lineno} 行：{msg}")

        if not s.startswith('$$'):
            line_no_disp = re.sub(r'\$\$.*?\$\$', '', clean_no_code)
            for m_inline in re.finditer(r'(?<!\$)\$(?!\$)([^$\n]+?)\$(?!\$)', line_no_disp):
                expr = m_inline.group(1)
                if expr.strip().endswith('='):
                    errors.append(
                        f"第 {lineno} 行：行内公式 `${expr}$` 以 '=' 结尾，LibreOffice 会渲染为 '¿'，请将等号移至公式外"
                    )
                if re.search(r'\\underline\s*\{', expr):
                    errors.append(
                        f"第 {lineno} 行：行内公式 `${expr}$` 内部包含 \\underline 填空线，请移至公式外改用 __________"
                    )

        # 4. 嵌图语法与文件存在性检查
        for m_img in re.finditer(r'!\[([^\]]*)\]\(([^)]+)\)(\{[^}]*\})?', clean_no_code):
            img_rel = m_img.group(2).strip()
            attrs = m_img.group(3) or ''
            if not re.search(r'\bwidth\s*=\s*[\d.]+cm\b', attrs):
                errors.append(
                    f"第 {lineno} 行：图片 `{img_rel}` 缺少标准固定嵌宽属性 `{{width=XXcm}}`"
                )
            if re.search(r'\bheight\s*=', attrs):
                errors.append(
                    f"第 {lineno} 行：图片 `{img_rel}` 严禁指定 `height=`（会破坏原图宽高比），仅保留 `{{width=XXcm}}`"
                )
            if require_resolved_figs:
                cand1 = os.path.join(md_dir, img_rel)
                cand2 = os.path.abspath(img_rel)
                if not (os.path.isfile(cand1) or os.path.isfile(cand2)):
                    errors.append(f"第 {lineno} 行：未找到图片文件 `{img_rel}`（查找路径：{cand1}）")

    if errors:
        msg = f"[PRECHECK FAIL] {md_path} 发现 {len(errors)} 处 Markdown 隐患：\n  - " + "\n  - ".join(errors)
        raise AssertionError(msg)
    print(f"[PRECHECK PASS] {md_path}")


def _is_teacher_markdown(text, path=''):
    base = os.path.basename(path).lower()
    if 'teacher' in base or '教师' in base:
        return True
    return '## 一、参考答案总览表' in text or '【分步解析】' in text


def validate_pair_contract(student_path, teacher_path):
    """
    学生版与教师版双卷 1:1 交付契约确定性联检（纯标准库）：
      1. 学生版题号连续（**1.** ~ **N.**），且严禁混入答案、解析或原题号提示；
      2. 教师版必含三大固定一级分节（一、参考答案总览表；二、原题考点诊断...；三、逐题参考答案与深度详析）；
      3. 教师版第一节选择题表每行不超过 10 题（防 A4 列宽挤压），且选择题+非选择题简答完整覆盖 1..N；
      4. 教师版第二节诊断映射表严格为 4 列，完整覆盖 1..N；
      5. 教师版第三节逐题详析完整覆盖 1..N，且每题均按顺序包含【参考答案】->【分步解析】->【防错点拨】。
    """
    with open(student_path, 'r', encoding='utf-8') as f:
        s_text = f.read()
    with open(teacher_path, 'r', encoding='utf-8') as f:
        t_text = f.read()

    errs = []

    # 1. 学生版题号与防泄题检查
    s_qnums = []
    in_fence = False
    for idx, line in enumerate(s_text.splitlines(), 1):
        st = line.strip()
        if st.startswith('```'):
            in_fence = not in_fence
            continue
        if in_fence:
            continue
        for marker in STUDENT_FORBIDDEN_MARKERS:
            if marker in line:
                errs.append(f"学生版第 {idx} 行：检测到教师版/原题泄题标记 `{marker}`")
        m_q = re.match(r'^\*\*(\d+)\.\*\*', st)
        if m_q:
            s_qnums.append(int(m_q.group(1)))

    if not s_qnums:
        errs.append("学生版未检测到任何标准题号（格式须为顶格 `**1.**` ～ `**N.**`）")
    else:
        expected_qnums = list(range(1, len(s_qnums) + 1))
        if s_qnums != expected_qnums:
            errs.append(f"学生版题号不连续或存在重号：实际 {s_qnums}，期望 {expected_qnums}")

    target_set = set(s_qnums)

    # 2. 教师版三大固定分节检查
    for req_sec in ('## 一、参考答案总览表', '## 二、原题考点诊断', '## 三、逐题参考答案与深度详析'):
        if req_sec not in t_text:
            errs.append(f"教师版缺失固定骨架标题：`{req_sec}...`")

    # 3~5. 解析教师版三节内容与学生版 1:1 比对
    sec = 0
    table_headers = None
    sec1_choices = set()
    sec1_nonchoices = set()
    sec2_qnums = set()
    sec3_blocks = {}
    cur_q = None
    buf = []

    def flush_q():
        nonlocal cur_q, buf
        if cur_q is not None:
            sec3_blocks[cur_q] = '\n'.join(buf)
        cur_q = None
        buf = []

    for idx, line in enumerate(t_text.splitlines(), 1):
        st = line.strip()
        if st.startswith('## 一、'):
            sec = 1
            continue
        if st.startswith('## 二、'):
            sec = 2
            continue
        if st.startswith('## 三、'):
            sec = 3
            continue

        if sec == 1:
            if st.startswith('|'):
                cells = _split_pipe_row_math_safe(st)
                c0 = cells[0].replace('*', '').strip() if cells else ''
                if c0 == '题号':
                    if len(cells) - 1 > 10:
                        errs.append(
                            f"教师版第 {idx} 行：参考答案表同行排入了 {len(cells) - 1} 道选择题（上限 10 道），"
                            "请按每表最多 10 题折成多张表以防 A4 列宽挤压"
                        )
                    table_headers = [c.replace('*', '').strip() for c in cells[1:]]
                elif '答案' in c0 and table_headers:
                    for h in table_headers:
                        if h.isdigit():
                            sec1_choices.add(int(h))
            else:
                m_nc = re.match(r'^(?:[-*+]\s*)?\*\*(\d+)\.?\*\*', st)
                if m_nc:
                    sec1_nonchoices.add(int(m_nc.group(1)))

        elif sec == 2:
            if st.startswith('|'):
                cells = _split_pipe_row_math_safe(st)
                c0 = cells[0].replace('*', '').strip() if cells else ''
                if c0 in ('题号', '') or set(c0) <= {':', '-', ' '}:
                    continue
                if len(cells) != 4:
                    errs.append(
                        f"教师版第 {idx} 行：诊断映射表须严格为 4 列，当前解析为 {len(cells)} 列"
                    )
                m_id = re.match(r'^(\d+)\b', c0)
                if m_id:
                    sec2_qnums.add(int(m_id.group(1)))

        elif sec == 3:
            m_q = re.match(r'^###\s*\*\*(\d+)\.\*\*$', st)
            if m_q:
                flush_q()
                cur_q = int(m_q.group(1))
                buf = []
                continue
            if re.match(r'^###\s+', st):
                errs.append(f"教师版第 {idx} 行：第三节题号标题格式须严格为 `### **题号.**`（当前为 `{st}`）")
            if cur_q is not None:
                if st == '---':
                    flush_q()
                else:
                    buf.append(line)
    flush_q()

    if target_set:
        sec1_all = sec1_choices | sec1_nonchoices
        if sec1_all != target_set:
            missing_1 = sorted(target_set - sec1_all)
            extra_1 = sorted(sec1_all - target_set)
            errs.append(f"教师版第一节答案总览与学生版题号不一致（缺失：{missing_1}，多余：{extra_1}）")
        if sec2_qnums != target_set:
            missing_2 = sorted(target_set - sec2_qnums)
            extra_2 = sorted(sec2_qnums - target_set)
            errs.append(f"教师版第二节诊断映射表与学生版题号不一致（缺失：{missing_2}，多余：{extra_2}）")
        if set(sec3_blocks.keys()) != target_set:
            missing_3 = sorted(target_set - set(sec3_blocks.keys()))
            extra_3 = sorted(set(sec3_blocks.keys()) - target_set)
            errs.append(f"教师版第三节逐题详析与学生版题号不一致（缺失：{missing_3}，多余：{extra_3}）")

        for q in sorted(target_set & set(sec3_blocks.keys())):
            blk = sec3_blocks[q]
            p_ans = blk.find('【参考答案】')
            p_step = blk.find('【分步解析】')
            p_tip = blk.find('【防错点拨】')
            if p_ans < 0 or p_step < 0 or p_tip < 0:
                errs.append(
                    f"教师版第三节第 {q} 题：缺少必填模块（须同时含【参考答案】、【分步解析】、【防错点拨】）"
                )
            elif not (p_ans < p_step < p_tip):
                errs.append(
                    f"教师版第三节第 {q} 题：模块顺序错误（须严格按【参考答案】->【分步解析】->【防错点拨】排列）"
                )

    if errs:
        msg = f"[CONTRACT FAIL] 双卷 1:1 契约联检发现 {len(errs)} 处违规：\n  - " + "\n  - ".join(errs)
        raise AssertionError(msg)
    print(f"[CONTRACT PASS] 学生版与教师版（共 {len(s_qnums)} 题）1:1 结构契约联检全部通过。")


def postprocess_and_verify(docx_path, expected_imgs=None, max_ar_err_pct=0.5,
                           min_print_scale=0.65, max_print_scale=1.15):
    """一步完成全框线与防孤行后处理及七重自动化终检。"""
    postprocess_docx(docx_path)
    return verify_docx(
        docx_path,
        expected_imgs=expected_imgs,
        max_ar_err_pct=max_ar_err_pct,
        min_print_scale=min_print_scale,
        max_print_scale=max_print_scale,
    )


def compile_bank(md_path, out_docx, reference_docx='reference.docx',
                 expected_imgs=None, export_pdf=True, preview_dir=None,
                 allow_fig_placeholders=False):
    """
    一键确定性编译管线：
      1. 执行阶段二严格 Markdown 预检（含未替换配图占位符与图片文件存在性核验）；
      2. 自动生成 Pandoc 兼容 A4 reference.docx；
      3. 绑定 --resource-path 与 -f markdown-implicit_figures 调用 Pandoc 编译 DOCX，拦截全部公式/资源告警；
      4. 执行后处理与七重自动化终检；
      5. 可选调用 LibreOffice 导出标准 A4 PDF 与 pdftoppm 渲染预览图供真实目视抽检。
    """
    if not shutil.which('pandoc'):
        raise RuntimeError("未检测到 pandoc 命令，请先安装 Pandoc。")
    precheck_markdown(md_path, require_resolved_figs=not allow_fig_placeholders)
    out_parent = os.path.dirname(os.path.abspath(out_docx))
    if out_parent:
        os.makedirs(out_parent, exist_ok=True)
    build_reference_docx(reference_docx)

    md_dir = os.path.dirname(os.path.abspath(md_path)) or '.'
    resource_path = os.pathsep.join([md_dir, '.'])
    proc = subprocess.run(
        [
            'pandoc', '-f', 'markdown-implicit_figures',
            f'--resource-path={resource_path}',
            md_path, '-o', out_docx, f'--reference-doc={reference_docx}',
        ],
        capture_output=True, text=True,
    )
    if proc.returncode != 0:
        raise RuntimeError(f"Pandoc 编译失败：\n{proc.stderr}")
    if proc.stderr and ('Could not convert TeX math' in proc.stderr or 'Could not fetch resource' in proc.stderr):
        raise AssertionError(f"[PANDOC FAIL] 检测到公式转换失败或图片资源丢失：\n{proc.stderr.strip()}")

    res = postprocess_and_verify(out_docx, expected_imgs=expected_imgs)

    if export_pdf:
        lo = shutil.which('libreoffice') or shutil.which('soffice')
        if not lo:
            print("[WARN] 环境未安装 LibreOffice，跳过 PDF 导出。")
        else:
            out_dir = os.path.dirname(os.path.abspath(out_docx)) or '.'
            subprocess.check_call([
                lo, '--headless', '--convert-to', 'pdf', out_docx, '--outdir', out_dir
            ])
            pdf_path = os.path.splitext(out_docx)[0] + '.pdf'
            res['pdf'] = pdf_path
            if preview_dir and shutil.which('pdftoppm'):
                os.makedirs(preview_dir, exist_ok=True)
                prefix = os.path.join(preview_dir, os.path.splitext(os.path.basename(out_docx))[0])
                subprocess.check_call(['pdftoppm', '-png', '-r', '110', pdf_path, prefix])
                res['preview_dir'] = preview_dir
    return res


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='高中物理变式题库一站式预检、契约联检、编译与七重终检工具')
    sub = parser.add_subparsers(dest='cmd', required=True)

    p_pre = sub.add_parser('precheck', help='执行 Markdown 预检；若同时传入学生版与教师版则自动执行 1:1 契约联检')
    p_pre.add_argument('md_paths', nargs='+', help='待预检的 Markdown 文件路径')
    p_pre.add_argument('--require-figs', action='store_true', help='要求所有待配图占位已替换且图片文件已存在')

    p_aud = sub.add_parser('audit', help='显式执行学生版与教师版 1:1 契约联检')
    p_aud.add_argument('student_md')
    p_aud.add_argument('teacher_md')

    p_comp = sub.add_parser('compile', help='从 Markdown 一键编译 DOCX/PDF 并执行七重终检')
    p_comp.add_argument('md_path')
    p_comp.add_argument('out_docx')
    p_comp.add_argument('--reference', default='reference.docx')
    p_comp.add_argument('--expected-imgs', type=int, default=None)
    p_comp.add_argument('--pdf', action='store_true')
    p_comp.add_argument('--preview-dir', default=None)
    p_comp.add_argument('--allow-fig-placeholders', action='store_true', help='允许保留阶段一待配图注释（仅用于无图草稿调试）')

    args = parser.parse_args()
    if args.cmd == 'precheck':
        for md_p in args.md_paths:
            precheck_markdown(md_p, require_resolved_figs=args.require_figs)
        if len(args.md_paths) == 2:
            with open(args.md_paths[0], 'r', encoding='utf-8') as f0, open(args.md_paths[1], 'r', encoding='utf-8') as f1:
                t0, t1 = f0.read(), f1.read()
            is_t0 = _is_teacher_markdown(t0, args.md_paths[0])
            is_t1 = _is_teacher_markdown(t1, args.md_paths[1])
            if not is_t0 and is_t1:
                validate_pair_contract(args.md_paths[0], args.md_paths[1])
            elif is_t0 and not is_t1:
                validate_pair_contract(args.md_paths[1], args.md_paths[0])
    elif args.cmd == 'audit':
        precheck_markdown(args.student_md)
        precheck_markdown(args.teacher_md)
        validate_pair_contract(args.student_md, args.teacher_md)
    elif args.cmd == 'compile':
        compile_bank(
            args.md_path,
            args.out_docx,
            reference_docx=args.reference,
            expected_imgs=args.expected_imgs,
            export_pdf=args.pdf,
            preview_dir=args.preview_dir,
            allow_fig_placeholders=args.allow_fig_placeholders,
        )
