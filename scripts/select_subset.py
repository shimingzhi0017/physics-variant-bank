r"""
母库按错题题号抽题组卷脚本（仅依赖 Python 标准库）。

核心复用能力：
  1. 兼容任意大题标题（多项选择题 / 多选题 / 不定项 / 混排范围“8~10为多选” / 无分节讲义）与多字母答案，
     自动精准补齐多选题 `（多选）` 标记；
  2. 选择题超过 10 道时自动按每表最多 10 题分行折表，彻底防止 A4 横向列宽挤压变形；
  3. 表格行解析自动避开公式内的绝对值符号（如 $|\Delta\Phi|$），兼容加粗表头与二次抽题；
  4. 支持单题多图（如 fig_q12a.png / fig_q12_1.png）同步重映射；若母库目录已存在对应 PNG，
     自动拷贝重命名至输出目录，仅对缺失图片输出待补绘清单；
  5. 默认剥离无分节讲义中残余的单题分值（如 `（8 分）`）以统一体例，亦可通过 `--keep-scores` 保留。

用法：
  python3 scripts/select_subset.py \
      --student student_bank.md \
      --teacher teacher_bank.md \
      --pick 3,7,10,14-15 \
      --title "成都市实验外国语学校 2026 级秋季周考·错题靶向变式卷" \
      --out-dir ./subset_out
"""
import argparse
import os
import re
import shutil

MAX_CHOICES_PER_TABLE = 10


def parse_pick(pick_str):
    nums = []
    for part in pick_str.split(','):
        part = part.strip()
        if not part:
            continue
        if '-' in part:
            a, b = part.split('-', 1)
            nums.extend(range(int(a), int(b) + 1))
        else:
            nums.append(int(part))
    return sorted(set(nums))


def split_pipe_row(line):
    """按 Markdown 表格分隔符 '|' 切分单元格，自动忽略 $...$ 公式内部的绝对值竖线与转义 \\|。"""
    s = line.strip()
    if s.startswith('|'):
        s = s[1:]
    if s.endswith('|') and not s.endswith('\\|'):
        s = s[:-1]
    cells = []
    cur = []
    in_math = False
    i = 0
    while i < len(s):
        ch = s[i]
        if ch == '\\' and i + 1 < len(s):
            cur.append(ch)
            cur.append(s[i + 1])
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


def split_student_bank(text):
    """解析学生版母库，返回 {题号: {'block': str, 'section': str}}。"""
    lines = text.splitlines()
    items = {}
    cur_sec = ''
    cur_q = None
    buf = []

    def flush():
        nonlocal cur_q, buf
        if cur_q is not None:
            items[cur_q] = {'block': '\n'.join(buf).strip(), 'section': cur_sec}
        cur_q = None
        buf = []

    for line in lines:
        s = line.strip()
        if s.startswith('## '):
            flush()
            cur_sec = s
            continue
        m = re.match(r'^\*\*(\d+)\.\*\*(.*)$', s)
        if m:
            flush()
            cur_q = int(m.group(1))
            buf = [line]
            continue
        if cur_q is not None:
            if s == '---':
                flush()
            else:
                buf.append(line)
    flush()
    return items


def _extract_brief_from_detail(detail_text):
    """当第 1 节未显式列出某非选择题简答时，自动从第 3 节【参考答案】首行提取。"""
    for line in detail_text.splitlines():
        s = line.strip()
        if '【参考答案】' in s:
            after = re.sub(r'^.*?【参考答案】\**\s*', '', s).strip()
            if after:
                return after
    lines = detail_text.splitlines()
    for idx, line in enumerate(lines):
        if '【参考答案】' in line:
            acc = []
            for sub in lines[idx + 1:]:
                st = sub.strip()
                if not st:
                    if acc:
                        break
                    continue
                if st.startswith(('**【', '![', '<!--')):
                    break
                acc.append(st)
            if acc:
                return '　'.join(acc)
    return ''


def parse_teacher_bank(text):
    """解析教师版母库：提取选择题答案映射、非选择题简答映射、4列诊断映射表行、逐题详析块。"""
    lines = text.splitlines()
    choice_ans = {}
    nonchoice_ans = {}
    diag_rows = {}
    details = {}

    sec = 0
    cur_q = None
    buf = []

    def flush_detail():
        nonlocal cur_q, buf
        if cur_q is not None:
            details[cur_q] = '\n'.join(buf).strip()
        cur_q = None
        buf = []

    table_headers = None
    for line in lines:
        s = line.strip()
        if s.startswith('## 一、'):
            sec = 1
            continue
        if s.startswith('## 二、'):
            sec = 2
            continue
        if s.startswith('## 三、'):
            sec = 3
            continue

        if sec == 1:
            if s.startswith('|'):
                cells = split_pipe_row(s)
                c0 = cells[0].replace('*', '').strip() if cells else ''
                if c0 == '题号':
                    table_headers = [c.replace('*', '').strip() for c in cells[1:]]
                elif '答案' in c0 and table_headers:
                    for h, v in zip(table_headers, cells[1:]):
                        if h.isdigit():
                            choice_ans[int(h)] = v.replace('*', '').strip()
            else:
                m_nc = re.match(r'^(?:[-*+]\s*)?\*\*(\d+)\.?\*\*[:：]?\s*(.+)$', s)
                if m_nc:
                    nonchoice_ans[int(m_nc.group(1))] = m_nc.group(2).strip()

        elif sec == 2:
            if s.startswith('|'):
                cells = split_pipe_row(s)
                if len(cells) >= 4:
                    c0_clean = cells[0].replace('*', '').strip()
                    m_id = re.match(r'^(\d+)\b', c0_clean)
                    if m_id:
                        diag_rows[int(m_id.group(1))] = cells[1:4]

        elif sec == 3:
            m_q = re.match(r'^###\s*(?:\*\*(\d+)\.?\*\*|(\d+)\.)', s)
            if m_q:
                flush_detail()
                cur_q = int(m_q.group(1) or m_q.group(2))
                buf = []
                continue
            if cur_q is not None:
                if s == '---':
                    flush_detail()
                else:
                    buf.append(line)

    flush_detail()
    return choice_ans, nonchoice_ans, diag_rows, details


def _is_multi_choice(old_q, entry, choice_ans):
    sec = entry['section']
    first_line = entry['block'].splitlines()[0] if entry['block'] else ''
    if any(k in first_line for k in ('（多选）', '(多选)', '（不定项）', '(不定项)')):
        return True
    # 优先匹配混排标题中的精确题号范围，如“一、选择题（其中 1~8 题为单选，9~11 题为多选）”
    for m in re.finditer(r'(\d+)\s*[~～\-—至]\s*(\d+)[^，,。；;\n]*?(?:多选|不定项)', sec):
        if int(m.group(1)) <= old_q <= int(m.group(2)):
            return True
    if '单选' not in sec and any(k in sec for k in ('多项选择', '多选题', '不定项')):
        return True
    if old_q in choice_ans:
        letters = re.findall(r'[A-E]', choice_ans[old_q].upper())
        if len(letters) >= 2:
            return True
    return False


def renumber_and_sync_figs(block, old_q, new_q, src_root, out_dir):
    """
    将题块中的 fig_qXX / sol_qXX（含多子图后缀如 fig_q12a.png、fig_q12_1.png）重映射为新题号；
    若母库目录中已存在对应 PNG，自动复制重命名至 out_dir，否则记入待补绘列表。
    """
    pat = re.compile(rf'\b(fig|sol)_q0*{old_q}((?:_[a-zA-Z0-9]+|[a-z])?\.png)\b')
    todos = []
    seen = set()

    def _repl(m):
        kind, suffix = m.group(1), m.group(2)
        sub_folder = 'figs' if kind == 'fig' else 'figs_sol'
        old_rel = f'{sub_folder}/{kind}_q{old_q:02d}{suffix}'
        old_alt = f'{sub_folder}/{m.group(0)}'
        new_name = f'{kind}_q{new_q:02d}{suffix}'
        new_rel = f'{sub_folder}/{new_name}'

        if new_rel not in seen:
            seen.add(new_rel)
            src_candidates = [os.path.join(src_root, old_rel), os.path.join(src_root, old_alt)]
            src_hit = next((p for p in src_candidates if os.path.isfile(p)), None)
            if src_hit:
                dst_path = os.path.join(out_dir, new_rel)
                os.makedirs(os.path.dirname(dst_path), exist_ok=True)
                if os.path.abspath(src_hit) != os.path.abspath(dst_path):
                    shutil.copy2(src_hit, dst_path)
            else:
                todos.append(f'新第 {new_q:02d} 题（原第 {old_q:02d} 题）-> {new_rel}')
        return new_name

    new_block = pat.sub(_repl, block)
    return new_block, todos


def build_subset(student_path, teacher_path, picks, title, out_dir,
                 suggested_time='40 分钟', keep_scores=False):
    src_root = os.path.dirname(os.path.abspath(student_path))
    with open(student_path, 'r', encoding='utf-8') as f:
        s_items = split_student_bank(f.read())
    with open(teacher_path, 'r', encoding='utf-8') as f:
        choice_ans, nonchoice_ans, diag_rows, t_details = parse_teacher_bank(f.read())

    missing = [q for q in picks if q not in s_items or q not in t_details]
    if missing:
        raise ValueError(f'母库中未找到以下题号：{missing}')

    os.makedirs(out_dir, exist_ok=True)

    # 1. 生成学生版子集
    s_lines = [
        f'# {title}（学生版）',
        '',
        f'**建议用时：** {suggested_time}',
        '',
        '---',
        '',
    ]
    fig_todos = []
    for new_q, old_q in enumerate(picks, 1):
        entry = s_items[old_q]
        block = entry['block']
        is_multi = _is_multi_choice(old_q, entry, choice_ans)
        has_multi_tag = any(k in block.splitlines()[0] for k in ('（多选）', '(多选)', '（不定项）', '(不定项)'))
        prefix = f'**{new_q}.**' + ('（多选）' if is_multi and not has_multi_tag else '')
        if keep_scores:
            block = re.sub(r'^\*\*\d+\.\*\*', prefix, block, count=1)
        else:
            # 剥离单题残余分值（如 `（8 分）` / `(12分)`），使混编错题讲义体例一致
            block = re.sub(r'^\*\*\d+\.\*\*\s*(?:[（(]\s*\d+\s*分\s*[）)]\s*)?', prefix + ' ', block, count=1)
            block = re.sub(r'^(\*\*\d+\.\*\*(?:（多选）)?)\s+', r'\1 ', block, count=1)

        block, todos = renumber_and_sync_figs(block, old_q, new_q, src_root, out_dir)
        fig_todos.extend(todos)
        s_lines.append(block)
        s_lines.append('')

    out_s = os.path.join(out_dir, 'student_bank.md')
    with open(out_s, 'w', encoding='utf-8') as f:
        f.write('\n'.join(s_lines).rstrip() + '\n')

    # 2. 生成教师版子集
    t_lines = [
        f'# {title}（教师版）',
        '',
        '## 一、参考答案总览表',
        '',
    ]
    picked_choices = []
    picked_nonchoices = []
    for new_q, old_q in enumerate(picks, 1):
        if old_q in choice_ans:
            picked_choices.append((new_q, old_q))
        else:
            ans_str = nonchoice_ans.get(old_q) or _extract_brief_from_detail(t_details[old_q])
            picked_nonchoices.append((new_q, old_q, ans_str))

    # 选择题按每表最多 10 题折行，防止多题挤爆 A4 页面宽度
    for start in range(0, len(picked_choices), MAX_CHOICES_PER_TABLE):
        chunk = picked_choices[start:start + MAX_CHOICES_PER_TABLE]
        t_lines.append('| 题号 | ' + ' | '.join(str(nq) for nq, _ in chunk) + ' |')
        t_lines.append('| :---: | ' + ' | '.join(':---:' for _ in chunk) + ' |')
        t_lines.append('| **答案** | ' + ' | '.join(choice_ans[oq] for _, oq in chunk) + ' |')
        t_lines.append('')

    if picked_nonchoices:
        t_lines.append('**非选择题简答：**')
        t_lines.append('')
        for nq, _, ans_str in picked_nonchoices:
            sep = '' if ans_str.startswith(('（', '(')) else ' '
            t_lines.append(f'- **{nq}.**{sep}{ans_str}')
        t_lines.append('')

    t_lines.extend([
        '---',
        '',
        '## 二、原题考点诊断、答案核查与变式映射总览表',
        '',
        '| 题号 | 原题核心考点与高频易错点 | 原题参考答案核查 | 变式核心模型与重构切入点 |',
        '| :---: | :--- | :--- | :--- |',
    ])
    for nq, oq in enumerate(picks, 1):
        cols = diag_rows.get(oq, ['—', '—', '—'])
        t_lines.append(f'| **{nq}**（原第 {oq} 题） | {cols[0]} | {cols[1]} | {cols[2]} |')

    t_lines.extend([
        '',
        '---',
        '',
        '## 三、逐题参考答案与深度详析',
        '',
    ])
    sol_todos = []
    for nq, oq in enumerate(picks, 1):
        detail, todos = renumber_and_sync_figs(t_details[oq], oq, nq, src_root, out_dir)
        sol_todos.extend(todos)
        t_lines.append(f'### **{nq}.**')
        t_lines.append(detail)
        t_lines.append('')
        t_lines.append('---')
        t_lines.append('')

    out_t = os.path.join(out_dir, 'teacher_bank.md')
    with open(out_t, 'w', encoding='utf-8') as f:
        f.write('\n'.join(t_lines).rstrip() + '\n')

    print(f'[OK] 子集抽题完成（共 {len(picks)} 题：{picks}）')
    print(f'  - 学生版：{out_s}')
    print(f'  - 教师版：{out_t}')
    if fig_todos:
        print('  - 待补绘学生题图：')
        for item in fig_todos:
            print(f'      * {item}')
    if sol_todos:
        print('  - 待补绘教师解析图：')
        for item in sol_todos:
            print(f'      * {item}')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='从全卷母库按错题题号自动抽题、重编号并裁剪教师版')
    parser.add_argument('--student', required=True, help='母库 student_bank.md 路径')
    parser.add_argument('--teacher', required=True, help='母库 teacher_bank.md 路径')
    parser.add_argument('--pick', required=True, help='需抽取的题号列表，如 3,7,10,14-15')
    parser.add_argument('--title', required=True, help='子集试卷标题')
    parser.add_argument('--out-dir', required=True, help='输出目录')
    parser.add_argument('--time', default='40 分钟', help='建议用时文本')
    parser.add_argument('--keep-scores', action='store_true', help='保留原母库单题分值标注')
    args = parser.parse_args()
    build_subset(
        args.student, args.teacher, parse_pick(args.pick),
        args.title, args.out_dir, args.time, keep_scores=args.keep_scores
    )
