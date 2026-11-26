"""
原卷 / 答案（DOCX / DOC / PDF）无损辅助提取与整页视觉预览脚本（仅依赖 Python 标准库 + 系统工具）。

核心能力：
  1. DOCX 输入：
     - 将原生 OMML 公式（含分式、根号、上下标、核反应前置上下标 sPre、矢量重音 acc、
       上划线 bar、求和积分 nary、方程组 eqArr 等）还原为可读 LaTeX；
     - 同时解析 DrawingML (<w:drawing>/<a:blip>) 与旧版 VML/OLE (<w:pict>/<w:object>/<v:imagedata>)，
       准确标注 [IMG: media/xxx] 锚点，并对 <mc:AlternateContent> 只取主分支防漏防重、提取文本框内容；
     - 若检测到文档内含 .wmf/.emf 对象（常见于学科网/菁优网 MathType 公式或矢量图）或显式指定
       --render-pages，自动调用 LibreOffice + pdftoppm 导出整页高清 PNG（pages/page-*.png），
       彻底杜绝 MathType 公式在纯文本中变成空白。
  2. PDF / DOC 输入：
     - 自动导出整页高清 PNG（pages/page-*.png）供多模态视觉逐页精读，并辅以 pdftotext 粗索引。

用法：
  python3 scripts/extract_source_docx.py 原卷.docx --out-dir /tmp/src_paper
  python3 scripts/extract_source_docx.py 原卷.pdf  --out-dir /tmp/src_paper
"""
import argparse
import os
import shutil
import subprocess
import xml.etree.ElementTree as ET
import zipfile

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'v': 'urn:schemas-microsoft-com:vml',
    'o': 'urn:schemas-microsoft-com:office:office',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
}
REL_EMBED = f"{{{NS['r']}}}embed"
REL_ID = f"{{{NS['r']}}}id"
O_RELID = f"{{{NS['o']}}}relid"
M_VAL = f"{{{NS['m']}}}val"

SKIP_OMML_PR = {
    'rPr', 'fPr', 'radPr', 'dPr', 'ctrlPr', 'sPrePr', 'sSubPr', 'sSupPr',
    'sSubSupPr', 'accPr', 'barPr', 'naryPr', 'eqArrPr', 'limLowPr', 'limUppPr',
    'boxPr', 'borderBoxPr', 'mPr', 'funcPr',
}


def _local(tag):
    return tag.split('}')[-1] if '}' in tag else tag


def _join_children(elem):
    return ''.join(render_omml(c) for c in elem) if elem is not None else ''


def render_omml(elem):
    """递归将 OMML 节点转写为可读 LaTeX 表达式。"""
    tag = _local(elem.tag)
    if tag in SKIP_OMML_PR:
        return ''
    if tag == 't':
        return elem.text or ''
    if tag == 'f':
        return f"\\frac{{{_join_children(elem.find('m:num', NS))}}}{{{_join_children(elem.find('m:den', NS))}}}"
    if tag == 'rad':
        deg_s = _join_children(elem.find('m:deg', NS)).strip()
        e_s = _join_children(elem.find('m:e', NS))
        return f'\\sqrt[{deg_s}]{{{e_s}}}' if deg_s else f'\\sqrt{{{e_s}}}'
    if tag == 'sSup':
        return f"{_join_children(elem.find('m:e', NS))}^{{{_join_children(elem.find('m:sup', NS))}}}"
    if tag == 'sSub':
        return f"{_join_children(elem.find('m:e', NS))}_{{{_join_children(elem.find('m:sub', NS))}}}"
    if tag == 'sSubSup':
        return (
            f"{_join_children(elem.find('m:e', NS))}"
            f"_{{{_join_children(elem.find('m:sub', NS))}}}"
            f"^{{{_join_children(elem.find('m:sup', NS))}}}"
        )
    if tag == 'sPre':
        # 核反应方程前置质量数/电荷数，如 {}_{92}^{235}U
        sub_s = _join_children(elem.find('m:sub', NS))
        sup_s = _join_children(elem.find('m:sup', NS))
        e_s = _join_children(elem.find('m:e', NS))
        return f"{{}}_{{{sub_s}}}^{{{sup_s}}}{e_s}"
    if tag == 'acc':
        acc_pr = elem.find('m:accPr', NS)
        chr_el = acc_pr.find('m:chr', NS) if acc_pr is not None else None
        ch = chr_el.attrib.get(M_VAL, '⃗') if chr_el is not None else '⃗'
        e_s = _join_children(elem.find('m:e', NS))
        if ch in ('⃗', '⃗', '→'):
            return f'\\vec{{{e_s}}}'
        if ch in ('¯', 'ˉ', '̄', '̅'):
            return f'\\bar{{{e_s}}}'
        if ch in ('^', 'ˆ', '̂'):
            return f'\\hat{{{e_s}}}'
        if ch in ('˙', '̇'):
            return f'\\dot{{{e_s}}}'
        return f'\\vec{{{e_s}}}'
    if tag == 'bar':
        return f"\\overline{{{_join_children(elem.find('m:e', NS))}}}"
    if tag == 'func':
        fn_s = _join_children(elem.find('m:fName', NS)).strip()
        e_s = _join_children(elem.find('m:e', NS))
        return f"{fn_s} {e_s}" if fn_s else e_s
    if tag == 'nary':
        nary_pr = elem.find('m:naryPr', NS)
        chr_el = nary_pr.find('m:chr', NS) if nary_pr is not None else None
        op = chr_el.attrib.get(M_VAL, '∑') if chr_el is not None else '∑'
        sub_s = _join_children(elem.find('m:sub', NS)).strip()
        sup_s = _join_children(elem.find('m:sup', NS)).strip()
        e_s = _join_children(elem.find('m:e', NS))
        res = op
        if sub_s:
            res += f'_{{{sub_s}}}'
        if sup_s:
            res += f'^{{{sup_s}}}'
        return f'{res} {e_s}'
    if tag == 'limLow':
        return f"{_join_children(elem.find('m:e', NS))}_{{{_join_children(elem.find('m:lim', NS))}}}"
    if tag == 'limUpp':
        return f"{_join_children(elem.find('m:e', NS))}^{{{_join_children(elem.find('m:lim', NS))}}}"
    if tag == 'eqArr':
        eqs = [_join_children(e) for e in elem.findall('m:e', NS)]
        return ' \\\\ '.join(eqs)
    if tag == 'm':
        rows = []
        for mr in elem.findall('m:mr', NS):
            rows.append(' & '.join(_join_children(e) for e in mr.findall('m:e', NS)))
        return '\\begin{matrix} ' + ' \\\\ '.join(rows) + ' \\end{matrix}'
    if tag == 'd':
        dpr = elem.find('m:dPr', NS)
        beg, end = '(', ')'
        if dpr is not None:
            bc = dpr.find('m:begChr', NS)
            ec = dpr.find('m:endChr', NS)
            if bc is not None:
                beg = bc.attrib.get(M_VAL, '(')
            if ec is not None:
                end = ec.attrib.get(M_VAL, ')')
        inner = ', '.join(_join_children(e) for e in elem.findall('m:e', NS))
        return f'{beg}{inner}{end}'
    return _join_children(elem)


def _extract_run_content(r_elem, rels):
    parts = []
    for sub in r_elem:
        stag = _local(sub.tag)
        if stag == 'AlternateContent':
            choice = sub.find('mc:Choice', NS)
            target = choice if choice is not None else sub.find('mc:Fallback', NS)
            if target is not None:
                parts.append(_extract_run_content(target, rels))
            continue
        if stag == 't':
            parts.append(sub.text or '')
        elif stag == 'tab':
            parts.append('\t')
        elif stag in ('br', 'cr'):
            parts.append(' ')
        elif stag == 'drawing':
            for blip in sub.findall('.//a:blip', NS):
                rid = blip.attrib.get(REL_EMBED)
                if rid and rid in rels:
                    parts.append(f' [IMG: {rels[rid]}] ')
            for txbx in sub.findall('.//w:txbxContent', NS):
                tb_lines = [render_paragraph(tp, rels) for tp in txbx.findall('w:p', NS)]
                tb_txt = ' / '.join(x for x in tb_lines if x)
                if tb_txt:
                    parts.append(f' [文本框: {tb_txt}] ')
        elif stag in ('pict', 'object'):
            for imdata in sub.findall('.//v:imagedata', NS):
                rid = imdata.attrib.get(REL_ID) or imdata.attrib.get(O_RELID)
                if rid and rid in rels:
                    target = rels[rid]
                    ext = os.path.splitext(target)[1].lower()
                    tag_lbl = 'OLE/WMF' if ext in ('.wmf', '.emf') else 'IMG'
                    parts.append(f' [{tag_lbl}: {target}] ')
            for txbx in sub.findall('.//w:txbxContent', NS):
                tb_lines = [render_paragraph(tp, rels) for tp in txbx.findall('w:p', NS)]
                tb_txt = ' / '.join(x for x in tb_lines if x)
                if tb_txt:
                    parts.append(f' [文本框: {tb_txt}] ')
    return ''.join(parts)


def render_paragraph(p, rels):
    parts = []
    for child in p:
        tag = _local(child.tag)
        if tag == 'r':
            parts.append(_extract_run_content(child, rels))
        elif tag in ('oMath', 'oMathPara'):
            parts.append(f'${render_omml(child)}$')
        elif tag == 'AlternateContent':
            choice = child.find('mc:Choice', NS)
            target = choice if choice is not None else child.find('mc:Fallback', NS)
            if target is not None:
                parts.append(render_paragraph(target, rels))
        elif tag in ('hyperlink', 'smartTag', 'ins', 'fldSimple', 'sdt', 'sdtContent'):
            parts.append(render_paragraph(child, rels))
    return ''.join(parts).strip()


def _render_pdf_pages(pdf_path, out_dir, dpi=160):
    if not shutil.which('pdftoppm'):
        print('[WARN] 未安装 pdftoppm，无法渲染整页 PNG 预览图。')
        return 0
    pages_dir = os.path.join(out_dir, 'pages')
    os.makedirs(pages_dir, exist_ok=True)
    prefix = os.path.join(pages_dir, 'page')
    subprocess.check_call(['pdftoppm', '-png', '-r', str(dpi), pdf_path, prefix])
    pngs = sorted(f for f in os.listdir(pages_dir) if f.endswith('.png'))
    print(f'[OK] 已渲染整页高清图 {len(pngs)} 页至: {pages_dir}')
    return len(pngs)


def _convert_office_to_pdf_pages(src_path, out_dir, dpi=160):
    lo = shutil.which('libreoffice') or shutil.which('soffice')
    if not lo:
        print('[WARN] 未安装 LibreOffice，跳过整页视觉预览渲染。')
        return 0
    subprocess.check_call(
        [lo, '--headless', '--convert-to', 'pdf', src_path, '--outdir', out_dir],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    base = os.path.splitext(os.path.basename(src_path))[0]
    pdf_path = os.path.join(out_dir, f'{base}.pdf')
    if os.path.isfile(pdf_path):
        return _render_pdf_pages(pdf_path, out_dir, dpi=dpi)
    return 0


def extract_source(src_path, out_dir, render_pages=False, dpi=160):
    os.makedirs(out_dir, exist_ok=True)
    ext = os.path.splitext(src_path)[1].lower()

    if ext == '.pdf':
        txt_out = os.path.join(out_dir, 'extracted.md')
        if shutil.which('pdftotext'):
            raw = subprocess.check_output(['pdftotext', src_path, '-'], text=True)
            with open(txt_out, 'w', encoding='utf-8') as f:
                f.write(
                    '> [注意] PDF 纯文本提取极易丢失公式（如 MathType 嵌入对象会变空白）和根号结构，'
                    '请务必逐页目视读取 pages/page-*.png 核对题干、公式与配图！\n\n' + raw
                )
            print(f'[OK] PDF 粗文本索引已保存: {txt_out}')
        _render_pdf_pages(src_path, out_dir, dpi=dpi)
        return txt_out

    if ext == '.doc':
        print('[INFO] 检测到旧版二进制 .doc，自动转为整页高清 PNG 供视觉读取...')
        _convert_office_to_pdf_pages(src_path, out_dir, dpi=dpi)
        return os.path.join(out_dir, 'pages')

    media_out = os.path.join(out_dir, 'media')
    os.makedirs(media_out, exist_ok=True)
    wmf_count = 0

    with zipfile.ZipFile(src_path, 'r') as zf:
        rels = {}
        if 'word/_rels/document.xml.rels' in zf.namelist():
            rtree = ET.fromstring(zf.read('word/_rels/document.xml.rels'))
            for rel in rtree:
                rels[rel.attrib['Id']] = rel.attrib['Target']

        img_count = 0
        for name in zf.namelist():
            if name.startswith('word/media/') and not name.endswith('/'):
                dest = os.path.join(media_out, os.path.basename(name))
                with open(dest, 'wb') as f:
                    f.write(zf.read(name))
                img_count += 1
                if name.lower().endswith(('.wmf', '.emf')):
                    wmf_count += 1

        root = ET.fromstring(zf.read('word/document.xml'))
        body = root.find('w:body', NS)
        lines = []
        for child in body:
            tag = _local(child.tag)
            if tag == 'p':
                txt = render_paragraph(child, rels)
                if txt:
                    lines.append(txt)
            elif tag == 'tbl':
                rows = []
                max_cols = 0
                for tr in child.findall('w:tr', NS):
                    cells = []
                    for tc in tr.findall('w:tc', NS):
                        c_lines = [render_paragraph(p, rels) for p in tc.findall('w:p', NS)]
                        cells.append(' '.join(x for x in c_lines if x).replace('|', '\\|'))
                    max_cols = max(max_cols, len(cells))
                    rows.append(cells)
                if rows and max_cols > 0:
                    norm_rows = ['| ' + ' | '.join(r + [''] * (max_cols - len(r))) + ' |' for r in rows]
                    sep = '| ' + ' | '.join(['---'] * max_cols) + ' |'
                    lines.append(norm_rows[0])
                    lines.append(sep)
                    lines.extend(norm_rows[1:])
                    lines.append('')

    out_md = os.path.join(out_dir, 'extracted.md')
    with open(out_md, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines) + '\n')
    print(f'[OK] 提取完成：{out_md}（共 {len(lines)} 行，导出媒体 {img_count} 个至 {media_out}）')

    if render_pages or wmf_count > 0:
        if wmf_count > 0:
            print(
                f'[NOTICE] 检测到 {wmf_count} 个 WMF/EMF 对象（可能含 MathType 公式或旧版矢量图），'
                '自动渲染整页高清图 pages/page-*.png 以确保零盲区核对...'
            )
        _convert_office_to_pdf_pages(src_path, out_dir, dpi=dpi)

    return out_md


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='无损提取原卷/答案（DOCX/DOC/PDF）中的公式、图片锚点与整页预览')
    parser.add_argument('src_path', help='源 .docx / .doc / .pdf 文件路径')
    parser.add_argument('--out-dir', required=True, help='输出目录（自动生成 extracted.md、media/ 及按需 pages/）')
    parser.add_argument('--render-pages', action='store_true', help='强制渲染整页 PNG 预览图（PDF/含WMF时自动开启）')
    parser.add_argument('--dpi', type=int, default=160, help='整页预览渲染 DPI（默认 160）')
    args = parser.parse_args()
    extract_source(args.src_path, args.out_dir, render_pages=args.render_pages, dpi=args.dpi)
