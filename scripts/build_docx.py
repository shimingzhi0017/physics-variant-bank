"""
高中物理变式题库一站式确定性编译、后处理与七重自动化终检库。

设计目标：无论在哪个项目、哪次会话运行，均通过唯一标准管线生成完全一致的 A4 印刷级 DOCX 与 PDF。

一键命令行用法：
  python scripts/build_docx.py compile student_bank.md 输出_学生版.docx --expected-imgs 10 --pdf --preview-dir preview_s
  python scripts/build_docx.py compile teacher_bank.md 输出_教师版.docx --expected-imgs 9 --pdf --preview-dir preview_t
"""
import argparse
import os
import shutil
import subprocess
import tempfile
import zipfile
from io import BytesIO

from PIL import Image
from lxml import etree
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

EMU_PER_INCH = 914400.0
EMU_PER_CM = 360000.0
M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'


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
            ('compact', 9.5, 1.15, 2),
        ]:
            st = by_lower.get(sname)
            if st is None:
                continue
            st.font.name = 'Times New Roman'
            st.font.size = Pt(sz)
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
    """为 Word 表格显式写入黑色细实线全框线、单元格内边距及浅灰加粗表头。"""
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


def _is_plain_alpha_mrun(elem):
    """判断是否为带 m:sty='p' 且含字母/汉字的 OMML 字符 run（如单位 m/s、cm、Hz、点名 A/B、下标 min）。"""
    if elem.tag != f'{{{M_NS}}}r':
        return False
    rpr = elem.find(f'{{{M_NS}}}rPr')
    if rpr is None:
        return False
    sty = rpr.find(f'{{{M_NS}}}sty')
    if sty is None or sty.get(f'{{{M_NS}}}val') != 'p':
        return False
    txt = ''.join(t.text or '' for t in elem.findall(f'{{{M_NS}}}t'))
    return bool(txt) and any(ch.isalpha() for ch in txt)


def fix_omml_upright(doc):
    """
    修复 Pandoc 导出 OMML \\mathrm{...} 在 LibreOffice / WPS 中退化为斜体的跨平台缺陷：
      1. 对所有带 <m:sty m:val="p"/> 且含字母/汉字的 <m:r>，按 ECMA-376 XSD 顺序在 <m:rPr>
         首位补齐 <m:nor/>（运算符 =, +, /, × 等保持无 <m:nor/> 以保留原生数学间距）；
      2. 合并同一父节点下相邻的纯字母正体 <m:r>（如 c+m -> cm, H+z -> Hz, m+i+n -> min, k+m -> km），
         避免多字符单位与下标在 LibreOffice StarMath 中出现额外字距缝隙。
    """
    m_r_tag = f'{{{M_NS}}}r'
    m_rpr_tag = f'{{{M_NS}}}rPr'
    m_nor_tag = f'{{{M_NS}}}nor'
    m_t_tag = f'{{{M_NS}}}t'

    fixed = 0
    for r in doc.element.iter(m_r_tag):
        if _is_plain_alpha_mrun(r):
            rpr = r.find(m_rpr_tag)
            if rpr.find(m_nor_tag) is None:
                rpr.insert(0, etree.Element(m_nor_tag))
                fixed += 1

    # 合并相邻的纯字母正体 <m:r>（如 cm, Hz, km, min, AB）
    parents = set()
    for r in doc.element.iter(m_r_tag):
        p = r.getparent()
        if p is not None:
            parents.add(p)

    for parent in parents:
        children = list(parent)
        i = 0
        while i < len(children) - 1:
            cur = children[i]
            nxt = children[i + 1]
            if _is_plain_alpha_mrun(cur) and _is_plain_alpha_mrun(nxt):
                cur_ts = cur.findall(m_t_tag)
                nxt_ts = nxt.findall(m_t_tag)
                cur_txt = ''.join(t.text or '' for t in cur_ts)
                nxt_txt = ''.join(t.text or '' for t in nxt_ts)
                # 仅合并无额外 Word 字符格式差异的纯字母正体 run
                if (cur_txt.isalpha() and nxt_txt.isalpha()
                        and cur.find(qn('w:rPr')) is None and nxt.find(qn('w:rPr')) is None
                        and cur_ts and nxt_ts):
                    cur_ts[-1].text = (cur_ts[-1].text or '') + nxt_txt
                    parent.remove(nxt)
                    children.pop(i + 1)
                    continue
            i += 1
    return fixed


def postprocess_docx(path):
    """后处理：卷头居中、含图段落居中、引导句与下方公式/插图同页绑定、全部表格写入全框线、OMML正体单位锁死。"""
    doc = Document(path)
    fix_omml_upright(doc)
    if doc.paragraphs:
        doc.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    if len(doc.paragraphs) > 1:
        t1 = doc.paragraphs[1].text
        if any(k in t1 for k in ('考试时间', '满分', '说明')):
            doc.paragraphs[1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    paras = doc.paragraphs
    for i, p in enumerate(paras):
        if p._p.xpath('.//*[local-name()="drawing"]'):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(6)
        # 标题或紧接公式/插图的冒号引导句自动设置 keep_with_next，防止页末孤行
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
    七重自动化交付终检（确保每次输出在页面、公式、表格、图片与印刷字号上确定一致）：
      1. 正文及表格中 0 残留 LaTeX '$' 源码；
      2. 无误判代码块样式 (SourceCode / VerbatimChar)；
      3. OMML 公式末尾无孤立 '='（防止 LibreOffice 渲染为 '¿'）；
      4. 含表格文档的 styles.xml 必含 Pandoc 'Table' 样式（防止 LibreOffice 表格坍塌）；
      5. 全部节的页面尺寸锁定为标准 A4（21.0 x 29.7 cm，误差 < 0.2 cm，防 Letter 漂移）；
      6. 嵌图总数准确，且全部图片宽高比误差 < max_ar_err_pct；
      7. 全部图片相对 300 DPI 原始尺寸的印刷缩放比落入 [min_print_scale, max_print_scale]，
         硬性锁定纸面等效印刷字号 >= 11.5 pt，杜绝打印字小看不清。
    """
    doc = Document(docx_path)

    # 1. A4 页面尺寸锁死检查
    for s_idx, sec in enumerate(doc.sections, 1):
        w_cm = sec.page_width / EMU_PER_CM
        h_cm = sec.page_height / EMU_PER_CM
        if abs(w_cm - 21.0) > 0.2 or abs(h_cm - 29.7) > 0.2:
            raise AssertionError(
                f"[FAIL] {docx_path} 第 {s_idx} 节尺寸为 {w_cm:.2f}x{h_cm:.2f} cm，非标准 A4 (21.0x29.7 cm)！"
            )

    # 2. 残留 '$' 检查
    all_text = "\n".join(p.text for p in doc.paragraphs)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                all_text += "\n" + cell.text
    raw_dollars = all_text.count('$')
    if raw_dollars > 0:
        raise AssertionError(f"[FAIL] {docx_path}: 发现 {raw_dollars} 处残留 '$' 源码未转为公式！")

    # 3. 误判代码块检查
    xml_str = doc.element.xml
    if 'SourceCode' in xml_str or 'VerbatimChar' in xml_str:
        raise AssertionError(f"[FAIL] {docx_path}: 检测到误判代码块样式，请检查 Markdown 是否存在 4 空格缩进！")

    # 4. OMML 末尾孤立 '=' 与正体单位/点名 <m:nor/> 锁死检查
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
            f"[FAIL] {docx_path}: 发现 {unfixed_rm} 处 OMML 正体字母/单位缺失 <m:nor/>，LibreOffice 会退化为斜体！请运行 postprocess_docx()。"
        )

    # 5. 表格内建 Table 样式检查
    if len(doc.tables) > 0:
        with zipfile.ZipFile(docx_path) as zf:
            styles_xml = zf.read('word/styles.xml').decode('utf-8', errors='ignore')
        if 'styleId="Table"' not in styles_xml:
            raise AssertionError(
                f"[FAIL] {docx_path}: 含表格但 word/styles.xml 缺失 Pandoc 'Table' 样式，LibreOffice 导出 PDF 会表格坍塌！"
            )

    # 6 & 7. 嵌图数量、宽高比与印刷缩放比检查
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
        f"| omml={len(omaths)} | tables={len(doc.tables)} | raw_$=0 | trailing_eq=0 | upright_rm=OK"
    )
    return {
        'imgs': len(inlines),
        'max_ar_err': max_err,
        'scale_range': (min_sc, max_sc) if inlines else None,
        'omml': len(omaths),
        'tables': len(doc.tables),
    }


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
                 expected_imgs=None, export_pdf=True, preview_dir=None):
    """
    一键确定性编译管线：
      1. 自动生成 Pandoc 兼容 A4 reference.docx；
      2. 调用 Pandoc 编译 DOCX；
      3. 执行后处理与七重自动化终检；
      4. 可选调用 LibreOffice 导出标准 A4 PDF 与 pdftoppm 渲染预览图供真实目视抽检。
    """
    if not shutil.which('pandoc'):
        raise RuntimeError("未检测到 pandoc 命令，请先安装 Pandoc。")
    build_reference_docx(reference_docx)
    subprocess.check_call([
        'pandoc', md_path, '-o', out_docx, f'--reference-doc={reference_docx}'
    ])
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
    parser = argparse.ArgumentParser(description='高中物理变式题库一站式编译与七重终检工具')
    sub = parser.add_subparsers(dest='cmd', required=True)
    p_comp = sub.add_parser('compile', help='从 Markdown 一键编译 DOCX/PDF 并终检')
    p_comp.add_argument('md_path')
    p_comp.add_argument('out_docx')
    p_comp.add_argument('--reference', default='reference.docx')
    p_comp.add_argument('--expected-imgs', type=int, default=None)
    p_comp.add_argument('--pdf', action='store_true')
    p_comp.add_argument('--preview-dir', default=None)

    args = parser.parse_args()
    if args.cmd == 'compile':
        compile_bank(
            args.md_path,
            args.out_docx,
            reference_docx=args.reference,
            expected_imgs=args.expected_imgs,
            export_pdf=args.pdf,
            preview_dir=args.preview_dir,
        )
