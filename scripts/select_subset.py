"""
母库按错题题号抽题组卷脚本（仅依赖 Python 标准库）。

适用场景：
  阶段一已完成整卷母库 student_bank.md 与 teacher_bank.md 编写；
  考后根据学生高频错题清单（如 --pick 3,7,10,14,15），自动抽取子集、连续重编号，
  同步裁剪教师版的“参考答案总览表”“原题考点映射总览表”与“逐题详析”，并输出待配图清单。

用法：
  python3 scripts/select_subset.py \
      --student student_bank.md \
      --teacher teacher_bank.md \
      --pick 3,7,10,14,15 \
      --title "成都市实验外国语学校 2026 级秋季周考·错题靶向变式卷" \
      --out-dir ./subset_out
"""
import argparse
import os
import re


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
    # 去重并保持原题序升序
    return sorted(set(nums))


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
                cells = [c.strip() for c in s.strip('|').split('|')]
                if cells and cells[0] == '题号':
                    table_headers = cells[1:]
                elif cells and '答案' in cells[0] and table_headers:
                    for h, v in zip(table_headers, cells[1:]):
                        if h.isdigit():
                            choice_ans[int(h)] = v
            else:
                m_nc = re.match(r'^-\s*\*\*(\d+)\.\*\*(.+)$', s)
                if m_nc:
                    nonchoice_ans[int(m_nc.group(1))] = m_nc.group(2).strip()

        elif sec == 2:
            if s.startswith('|'):
                cells = [c.strip() for c in s.strip('|').split('|')]
                if len(cells) >= 4:
                    m_id = re.match(r'^\*\*(\d+)\*\*$', cells[0]) or re.match(r'^(\d+)$', cells[0])
                    if m_id:
                        diag_rows[int(m_id.group(1))] = cells[1:4]

        elif sec == 3:
            m_q = re.match(r'^###\s*\*\*(\d+)\.\*\*$', s)
            if m_q:
                flush_detail()
                cur_q = int(m_q.group(1))
                buf = []
                continue
            if cur_q is not None:
                if s == '---':
                    flush_detail()
                else:
                    buf.append(line)

    flush_detail()
    return choice_ans, nonchoice_ans, diag_rows, details


def renumber_fig_refs(block, old_q, new_q):
    """将题块中的 fig_qXX / sol_qXX 图号同步映射为新题号。"""
    old_tag = f'q{old_q:02d}'
    new_tag = f'q{new_q:02d}'
    return block.replace(f'fig_{old_tag}.png', f'fig_{new_tag}.png').replace(f'sol_{old_tag}.png', f'sol_{new_tag}.png')


def build_subset(student_path, teacher_path, picks, title, out_dir, suggested_time='40 分钟'):
    with open(student_path, 'r', encoding='utf-8') as f:
        s_items = split_student_bank(f.read())
    with open(teacher_path, 'r', encoding='utf-8') as f:
        choice_ans, nonchoice_ans, diag_rows, t_details = parse_teacher_bank(f.read())

    missing = [q for q in picks if q not in s_items or q not in t_details]
    if missing:
        raise ValueError(f'母库中未找到以下题号：{missing}')

    os.makedirs(out_dir, exist_ok=True)

    # 1. 生成学生版子集（按连续编号 1..N，多选题若原本在多项选择题节则自动加（多选）标记）
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
        is_multi = '多项选择' in entry['section']
        # 替换题首题号
        prefix = f'**{new_q}.**' + ('（多选）' if is_multi and '（多选）' not in block.splitlines()[0] else '')
        block = re.sub(r'^\*\*\d+\.\*\*', prefix, block, count=1)
        block = renumber_fig_refs(block, old_q, new_q)
        s_lines.append(block)
        s_lines.append('')
        if f'fig_q{new_q:02d}.png' in block:
            fig_todos.append(f'新第 {new_q:02d} 题（原母库第 {old_q:02d} 题）-> figs/fig_q{new_q:02d}.png')

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
    picked_choices = [(new_q, old_q) for new_q, old_q in enumerate(picks, 1) if old_q in choice_ans]
    picked_nonchoices = [(new_q, old_q) for new_q, old_q in enumerate(picks, 1) if old_q in nonchoice_ans]

    if picked_choices:
        t_lines.append('| 题号 | ' + ' | '.join(str(nq) for nq, _ in picked_choices) + ' |')
        t_lines.append('| :---: | ' + ' | '.join(':---:' for _ in picked_choices) + ' |')
        t_lines.append('| **答案** | ' + ' | '.join(choice_ans[oq] for _, oq in picked_choices) + ' |')
        t_lines.append('')

    if picked_nonchoices:
        t_lines.append('**非选择题简答：**')
        t_lines.append('')
        for nq, oq in picked_nonchoices:
            t_lines.append(f'- **{nq}.**{nonchoice_ans[oq]}')
        t_lines.append('')

    t_lines.extend([
        '---',
        '',
        '## 二、原题考点诊断、答案核查与变式映射总览表',
        '',
        '| 题号 | 原题核心考点与高频易错点 | 原题参考答案核查 | 变式核心模型与思维升维切入点 |',
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
        detail = renumber_fig_refs(t_details[oq], oq, nq)
        t_lines.append(f'### **{nq}.**')
        t_lines.append(detail)
        t_lines.append('')
        t_lines.append('---')
        t_lines.append('')
        if f'sol_q{nq:02d}.png' in detail:
            sol_todos.append(f'新第 {nq:02d} 题（原母库第 {oq:02d} 题）-> figs_sol/sol_q{nq:02d}.png')

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
    args = parser.parse_args()
    build_subset(args.student, args.teacher, parse_pick(args.pick), args.title, args.out_dir, args.time)
