"""
原卷 / 答案 DOCX 无损辅助提取脚本（仅依赖 Python 标准库）。

核心作用：
  1. 将 Word 原生 OMML 公式还原为可读 LaTeX 形式（保留分式、根号、上下标、括号等），
     彻底避免宿主纯文本提取压平公式（如把 sqrt(2) 变成 2、把 1/t^2 变成 1t2）；
  2. 将嵌入图片 word/media/* 导出至指定目录，并在文本中准确标注 [IMG: media/imageX.png] 锚点；
  3. 同步提取表格内容（如答案速查表）。

用法：
  python3 scripts/extract_source_docx.py 原卷.docx --out-dir /tmp/src_paper
"""
import argparse
import os
import xml.etree.ElementTree as ET
import zipfile

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
REL_EMBED = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed'


def _local(tag):
    return tag.split('}')[-1] if '}' in tag else tag


def render_omml(elem):
    """递归将 OMML 节点转写为可读 LaTeX 表达式。"""
    tag = _local(elem.tag)
    if tag == 't':
        return elem.text or ''
    if tag == 'f':
        num = elem.find('m:num', NS)
        den = elem.find('m:den', NS)
        n_s = ''.join(render_omml(c) for c in num) if num is not None else ''
        d_s = ''.join(render_omml(c) for c in den) if den is not None else ''
        return f'\\frac{{{n_s}}}{{{d_s}}}'
    if tag == 'rad':
        deg = elem.find('m:deg', NS)
        e = elem.find('m:e', NS)
        deg_s = ''.join(render_omml(c) for c in deg).strip() if deg is not None else ''
        e_s = ''.join(render_omml(c) for c in e) if e is not None else ''
        return f'\\sqrt[{deg_s}]{{{e_s}}}' if deg_s else f'\\sqrt{{{e_s}}}'
    if tag == 'sSup':
        e = elem.find('m:e', NS)
        sup = elem.find('m:sup', NS)
        return f"{''.join(render_omml(c) for c in e)}^{{{''.join(render_omml(c) for c in sup)}}}"
    if tag == 'sSub':
        e = elem.find('m:e', NS)
        sub = elem.find('m:sub', NS)
        return f"{''.join(render_omml(c) for c in e)}_{{{''.join(render_omml(c) for c in sub)}}}"
    if tag == 'sSubSup':
        e = elem.find('m:e', NS)
        sub = elem.find('m:sub', NS)
        sup = elem.find('m:sup', NS)
        return (
            f"{''.join(render_omml(c) for c in e)}"
            f"_{{{''.join(render_omml(c) for c in sub)}}}"
            f"^{{{''.join(render_omml(c) for c in sup)}}}"
        )
    if tag == 'd':
        dpr = elem.find('m:dPr', NS)
        beg, end = '(', ')'
        if dpr is not None:
            bc = dpr.find('m:begChr', NS)
            ec = dpr.find('m:endChr', NS)
            if bc is not None:
                beg = bc.attrib.get(f"{{{NS['m']}}}val", '(')
            if ec is not None:
                end = ec.attrib.get(f"{{{NS['m']}}}val", ')')
        inner = ', '.join(''.join(render_omml(c) for c in e) for e in elem.findall('m:e', NS))
        return f'{beg}{inner}{end}'
    if tag in ('rPr', 'fPr', 'radPr', 'dPr', 'ctrlPr'):
        return ''
    return ''.join(render_omml(c) for c in elem)


def render_paragraph(p, rels):
    parts = []
    for child in p:
        tag = _local(child.tag)
        if tag == 'r':
            for sub in child:
                stag = _local(sub.tag)
                if stag == 't':
                    parts.append(sub.text or '')
                elif stag == 'tab':
                    parts.append('\t')
                elif stag == 'drawing':
                    for blip in sub.findall('.//a:blip', NS):
                        rid = blip.attrib.get(REL_EMBED)
                        if rid and rid in rels:
                            parts.append(f' [IMG: {rels[rid]}] ')
        elif tag in ('oMath', 'oMathPara'):
            parts.append(f'${render_omml(child)}$')
        elif tag == 'hyperlink':
            parts.append(render_paragraph(child, rels))
    return ''.join(parts).strip()


def extract_docx(docx_path, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    media_out = os.path.join(out_dir, 'media')
    os.makedirs(media_out, exist_ok=True)

    with zipfile.ZipFile(docx_path, 'r') as zf:
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
                for tr in child.findall('w:tr', NS):
                    cells = []
                    for tc in tr.findall('w:tc', NS):
                        c_lines = [render_paragraph(p, rels) for p in tc.findall('w:p', NS)]
                        cells.append(' '.join(x for x in c_lines if x))
                    rows.append('| ' + ' | '.join(cells) + ' |')
                if rows:
                    sep = '| ' + ' | '.join(['---'] * len(rows[0].split('|')[1:-1])) + ' |'
                    lines.append(rows[0])
                    lines.append(sep)
                    lines.extend(rows[1:])
                    lines.append('')

    out_md = os.path.join(out_dir, 'extracted.md')
    with open(out_md, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines) + '\n')
    print(f'[OK] 提取完成：{out_md}（共 {len(lines)} 行，导出图片 {img_count} 张至 {media_out}）')
    return out_md


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='无损提取原卷/答案 DOCX 中的 OMML 公式与图片锚点')
    parser.add_argument('docx_path', help='源 .docx 文件路径')
    parser.add_argument('--out-dir', required=True, help='输出目录（自动生成 extracted.md 与 media/）')
    args = parser.parse_args()
    extract_docx(args.docx_path, args.out_dir)
