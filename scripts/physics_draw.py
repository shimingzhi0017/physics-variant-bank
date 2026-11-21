"""
高中物理试卷纯黑白工程制图复用组件库（300 DPI 防压缩确定性印刷版）
核心特性：
  1. 4 档固定画布预设 CANVAS_PRESETS、紧裁剪白边与对应 Word 嵌宽 WORD_WIDTH_CM；
  2. 防压缩加粗三级线宽（2.60/2.00/1.45 pt）与大字号（25/23.5/22/16.5 pt）；
  3. 统一大号实心燕尾箭头（SwallowtailArrow：主矢量/坐标轴、尺寸线、引出线，虚线箭杆配实心燕尾头）；
  4. 自动从多语言 TTC 中精确提取缓存标准中国大陆简体常规宋体 (Noto Serif CJK SC Regular)，
     绝不静默退化为日文 JP 0 号字形或粗体；
  5. draw_block_on_slope 支持 slope_dir='down_right' 与 'up_right' 双向斜面严格贴合。
"""
import os
import math
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.path import Path
from matplotlib import font_manager
from matplotlib.font_manager import FontProperties

# ==================== 防缩印/防压缩印刷常量 ====================
COLOR_BLACK = '#000000'
COLOR_GRAY  = '#EFEFEF'
COLOR_WHITE = '#FFFFFF'

# 三级加粗线宽 (pt) —— 即使教师在 Word 中二次压缩 30%~40% 或黑白复印仍清晰不断线
LW_THICK = 2.60   # 外轮廓、刚性轨道、地面/墙体基准、主曲线
LW_MED   = 2.00   # 坐标轴、力/速度/加速度矢量、次轮廓
LW_THIN  = 1.45   # 虚线投影、尺寸线、细绳、弹簧、剖面线、浮置加速度

# 防压缩大字号常量 (pt) —— 默认嵌宽纸面等效 16~19 pt；即使缩窄至 6~10 cm 仍 >= 11~13 pt
FS_SYM = 25.0     # 物体、几何点名、物理量主符号 (A, B, F, v_0, H)
FS_AX  = 23.5     # 坐标轴标题、尺寸线标签 (t/s, F/N, \Delta x)
FS_CAP = 23.5     # 子图标题 (图甲, 图乙)
FS_TIK = 22.0     # 刻度数字、角度、面积标注 (30°, 37°, 40, 50)
FS_CN  = 22.0     # 器件中文标注 (光电门, 电磁铁, 安全门)
FS_SUB = 16.5     # 中文下标小字号 (F_弹, v_共)

# 燕尾箭头默认参数（以 mutation_scale=10 pt 为基准，严格脱离字号影响）
ARROW_HL_MAIN   = 1.28   # 主箭头（坐标轴、物理矢量）头长 12.8 pt
ARROW_HW_MAIN   = 0.46   # 主箭头半宽 4.6 pt（全宽 9.2 pt）
ARROW_NOTCH     = 0.32   # 燕尾后缘内凹比例（经典人教版物理教材燕尾比例）
ARROW_HL_DIM    = 1.05   # 尺寸线双向燕尾箭头头长 10.5 pt
ARROW_HW_DIM    = 0.38   # 尺寸线双向燕尾箭头半宽 3.8 pt
ARROW_HL_LEADER = 1.02   # 教师解析图引出线燕尾箭头头长 10.2 pt
ARROW_HW_LEADER = 0.36   # 教师解析图引出线燕尾箭头半宽 3.6 pt

# 4 档固定画布预设与对应 Word 嵌宽（紧裁剪导图 + 防压缩留余量）
CANVAS_PRESETS = {
    'tall': (3.8, 4.5),   # 竖直构型（自由落体、频闪、竖直上抛）
    'std':  (5.0, 3.7),   # 标准单图（函数图像、斜面、矢量图）
    'wide': (6.4, 3.0),   # 宽幅单图（纸带、双车、通道俯视图）
    'dual': (8.0, 3.9),   # 双子图（图甲+图乙、左右解析图）
}
WORD_WIDTH_CM = {
    'tall': 7.5,
    'std':  10.2,
    'wide': 12.8,
    'dual': 15.5,
}

_CN_FONT_KEYWORDS = ('Noto Serif CJK SC', 'Noto Serif CJK', 'SimSun', 'STSong', 'Songti',
                     'Noto Sans CJK SC', 'Noto Sans CJK', 'WenQuanYi')
_CACHED_CN_FONT_PATH = None


def _extract_sc_regular_from_ttc(ttc_path, out_path):
    """从多语言 TTC 字库合集中精确提取中国大陆简体常规体 (SC Regular) 单字库。"""
    try:
        from fontTools.ttLib import TTCollection
        ttc = TTCollection(ttc_path)
        for font in ttc.fonts:
            name = (font['name'].getDebugName(1) or '') + ' ' + (font['name'].getDebugName(4) or '')
            sub = font['name'].getDebugName(2) or ''
            if 'SC' in name.upper() and 'BOLD' not in sub.upper():
                os.makedirs(os.path.dirname(out_path), exist_ok=True)
                font.save(out_path)
                return out_path
    except Exception:
        return None
    return None


def find_cn_font_path(strict=True):
    """
    确定性定位中国大陆标准简体常规宋体。
    1. 优先复用同目录缓存的 NotoSerifCJKsc-Regular.otf；
    2. 若系统有 NotoSerifCJK-Regular.ttc，自动提取 SC Regular 子字体缓存；
    3. 否则匹配系统非粗体常规中文字体（拒绝 Bold 及未拆包的日文索引 0 TTC）；
    4. strict=True 时找不到合格简体字体直接抛错，避免静默产出怪异字形。
    """
    global _CACHED_CN_FONT_PATH
    if _CACHED_CN_FONT_PATH and os.path.exists(_CACHED_CN_FONT_PATH):
        return _CACHED_CN_FONT_PATH

    here = os.path.dirname(os.path.abspath(__file__))
    local_sc = os.path.join(here, 'NotoSerifCJKsc-Regular.otf')
    if os.path.exists(local_sc):
        _CACHED_CN_FONT_PATH = local_sc
        return local_sc

    fonts = font_manager.fontManager.ttflist
    for f in fonts:
        base = os.path.basename(f.fname).lower()
        if 'notoserifcjk' in base and 'bold' not in base and base.endswith('.ttc'):
            extracted = _extract_sc_regular_from_ttc(f.fname, local_sc)
            if extracted:
                _CACHED_CN_FONT_PATH = extracted
                return extracted

    for kw in _CN_FONT_KEYWORDS:
        for f in fonts:
            base = os.path.basename(f.fname).lower()
            if 'bold' in base or getattr(f, 'weight', 400) >= 600:
                continue
            if base.endswith('.ttc') and 'noto' in base:
                continue  # 禁止直接加载未提取 SC 子字体的 Noto TTC（索引 0 为日文 JP）
            if kw.lower() in f.name.lower() or kw.replace(' ', '').lower() in base:
                _CACHED_CN_FONT_PATH = f.fname
                return f.fname

    if strict:
        raise RuntimeError(
            "未找到标准简体中文常规体字体！请安装 Noto Serif CJK / SimSun 或安装 fonttools 以自动提取 SC Regular。"
        )
    return None


def get_cn_font(size=FS_CN):
    path = find_cn_font_path(strict=True)
    return FontProperties(fname=path, size=size)


def setup_style():
    plt.rcParams['mathtext.fontset'] = 'stix'
    plt.rcParams['font.family'] = 'serif'
    plt.rcParams['font.serif'] = ['DejaVu Serif', 'Liberation Serif', 'Times New Roman']
    plt.rcParams['axes.edgecolor'] = COLOR_BLACK
    plt.rcParams['text.color'] = COLOR_BLACK
    plt.rcParams['axes.unicode_minus'] = False


def new_fig(preset='std', figsize=None, equal=True, dpi=300):
    """按固定预设创建单图纯白无边框画布（默认 preset='std' -> (5.4, 3.9)）。"""
    setup_style()
    sz = figsize if figsize is not None else CANVAS_PRESETS[preset]
    fig, ax = plt.subplots(figsize=sz, dpi=dpi)
    if equal:
        ax.set_aspect('equal')
    ax.axis('off')
    return fig, ax


def new_subplots(nrows=1, ncols=2, preset='dual', figsize=None, width_ratios=None, equal=True, dpi=300):
    """按固定预设创建多子图纯白无边框画布（默认 preset='dual' -> (8.4, 4.1)）。"""
    setup_style()
    sz = figsize if figsize is not None else CANVAS_PRESETS[preset]
    kw = {'width_ratios': width_ratios} if width_ratios else None
    fig, axes = plt.subplots(nrows, ncols, figsize=sz, gridspec_kw=kw, dpi=dpi)
    for ax in np.atleast_1d(axes).ravel():
        if equal:
            ax.set_aspect('equal')
        ax.axis('off')
    return fig, axes


def save_fig(fig, path, pad=0.10, w_pad=0.95, dpi=300, preset=None):
    """自动创建目录，去除多余白边并以 300 DPI 纯白底导出；可选按 preset 检查预期印刷缩放比。"""
    out_dir = os.path.dirname(path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    plt.tight_layout(pad=pad, w_pad=w_pad)
    fig.savefig(path, dpi=dpi, facecolor=COLOR_WHITE, bbox_inches='tight', pad_inches=0.06)
    plt.close(fig)
    if preset in WORD_WIDTH_CM:
        from PIL import Image
        with Image.open(path) as im:
            w_px, _ = im.size
        expected_scale = (WORD_WIDTH_CM[preset] / 2.54) / (w_px / float(dpi))
        if not (0.68 <= expected_scale <= 1.12):
            print(
                f"[WARN] {path}: preset='{preset}' 对应嵌宽 {WORD_WIDTH_CM[preset]}cm 时"
                f"预计 print_scale={expected_scale:.2f}（安全区间 0.65~1.15），建议调整 preset 或微调 xlim/figsize。"
            )


# ==================== 通用几何与物理原语 ====================
class SwallowtailArrow(patches.ArrowStyle._Base):
    """
    国标/人教版物理教材标准实心燕尾箭头（单向 -> / 双向 <->）。
    在显示像素坐标系下计算几何顶点，严格不受 equal=False 纵横比拉伸变形影响，
    且箭杆自动止于燕尾内凹槽内，绝不刺穿箭尖。
    """
    def __init__(self, head_length=ARROW_HL_MAIN, head_width=ARROW_HW_MAIN,
                 notch=ARROW_NOTCH, both=False, draw_shaft=True, draw_head=True):
        self.head_length = head_length
        self.head_width = head_width
        self.notch = notch
        self.both = both
        self.draw_shaft = draw_shaft
        self.draw_head = draw_head
        super().__init__()

    def _make_head(self, x_from, y_from, x_tip, y_tip, L, W, notch, lw):
        dx, dy = x_tip - x_from, y_tip - y_from
        dist = np.hypot(dx, dy)
        if dist == 0:
            return None, (x_tip, y_tip)
        ux, uy = dx / dist, dy / dist
        nx, ny = -uy, ux
        half_sin = W / np.hypot(L, W)
        pad = 0.5 * lw / max(half_sin, 0.25)
        tx, ty = x_tip - pad * ux, y_tip - pad * uy
        wx1, wy1 = tx - L * ux + W * nx, ty - L * uy + W * ny
        wx2, wy2 = tx - L * ux - W * nx, ty - L * uy - W * ny
        mx, my = tx - L * (1.0 - notch) * ux, ty - L * (1.0 - notch) * uy
        # 箭杆深入燕尾实心区内部少许，防止与内凹处脱节，同时绝不穿透箭尖
        sx, sy = tx - L * (1.0 - notch * 1.15) * ux, ty - L * (1.0 - notch * 1.15) * uy
        verts = [(tx, ty), (wx1, wy1), (mx, my), (wx2, wy2), (tx, ty)]
        codes = [Path.MOVETO, Path.LINETO, Path.LINETO, Path.LINETO, Path.CLOSEPOLY]
        return Path(verts, codes), (sx, sy)

    def transmute(self, path, mutation_size, linewidth):
        L = self.head_length * mutation_size
        W = self.head_width * mutation_size
        v = path.vertices.copy()
        codes = path.codes
        paths, fills = [], []

        if self.both:
            hA, pA = self._make_head(v[1, 0], v[1, 1], v[0, 0], v[0, 1], L, W, self.notch, linewidth)
            v[0] = pA
        else:
            hA = None

        hB, pB = self._make_head(v[-2, 0], v[-2, 1], v[-1, 0], v[-1, 1], L, W, self.notch, linewidth)
        v[-1] = pB

        if self.draw_shaft:
            paths.append(Path(v, codes))
            fills.append(False)
        if self.draw_head:
            if hA is not None:
                paths.append(hA)
                fills.append(True)
            if hB is not None:
                paths.append(hB)
                fills.append(True)
        return paths, fills


def leader_arrowprops(hl=ARROW_HL_LEADER, hw=ARROW_HW_LEADER, notch=ARROW_NOTCH,
                      lw=LW_THIN, ls='-', shrinkA=4, shrinkB=2):
    """返回统一的大号实心燕尾引出线 arrowprops（供 ax.annotate 标注面积/交点使用）。"""
    return dict(
        arrowstyle=SwallowtailArrow(head_length=hl, head_width=hw, notch=notch),
        mutation_scale=10.0,
        color=COLOR_BLACK, lw=lw, ls=ls,
        joinstyle='miter', shrinkA=shrinkA, shrinkB=shrinkB,
    )


def draw_arrow(ax, x0, y0, x1, y1, lw=LW_MED, hl=ARROW_HL_MAIN, hw=ARROW_HW_MAIN,
               notch=ARROW_NOTCH, ls='-', zorder=8):
    """标准大号实心燕尾箭头（坐标轴、矢量、浮置加速度通用；虚线箭杆自动配实心燕尾头）。"""
    if ls == '-':
        ax.add_patch(patches.FancyArrowPatch(
            (x0, y0), (x1, y1),
            arrowstyle=SwallowtailArrow(head_length=hl, head_width=hw, notch=notch),
            mutation_scale=10.0, color=COLOR_BLACK, lw=lw, ls='-',
            joinstyle='miter', shrinkA=0, shrinkB=0, zorder=zorder
        ))
    else:
        # 虚线/点画线矢量：箭杆用虚线，燕尾头保持锐利纯黑实心不破碎
        ax.add_patch(patches.FancyArrowPatch(
            (x0, y0), (x1, y1),
            arrowstyle=SwallowtailArrow(head_length=hl, head_width=hw, notch=notch,
                                        draw_shaft=True, draw_head=False),
            mutation_scale=10.0, color=COLOR_BLACK, lw=lw, ls=ls,
            shrinkA=0, shrinkB=0, zorder=zorder
        ))
        ax.add_patch(patches.FancyArrowPatch(
            (x0, y0), (x1, y1),
            arrowstyle=SwallowtailArrow(head_length=hl, head_width=hw, notch=notch,
                                        draw_shaft=False, draw_head=True),
            mutation_scale=10.0, color=COLOR_BLACK, lw=lw, ls='-',
            joinstyle='miter', shrinkA=0, shrinkB=0, zorder=zorder
        ))


def draw_axes(ax, x_min, x_max, y_min, y_max, xlabel='', ylabel='',
              origin_label=r'$\mathrm{O}$', x_lbl_offset=(0.0, -0.48),
              y_lbl_offset=(0.24, 0.08), o_offset=(-0.34, -0.48),
              x_lbl_ha='center', lw=LW_MED):
    """绘制标准纯黑白直角坐标系。"""
    draw_arrow(ax, x_min, 0, x_max, 0, lw=lw)
    draw_arrow(ax, 0, y_min, 0, y_max, lw=lw)
    if xlabel:
        ax.text(x_max + x_lbl_offset[0], x_lbl_offset[1], xlabel,
                fontsize=FS_AX, ha=x_lbl_ha, va='top')
    if ylabel:
        ax.text(y_lbl_offset[0], y_max + y_lbl_offset[1], ylabel,
                fontsize=FS_AX, ha='left', va='center')
    if origin_label:
        ax.text(o_offset[0], o_offset[1], origin_label,
                fontsize=FS_TIK, ha='center', va='center')


def draw_ground(ax, x_start, x_end, y, hatch_len=0.22, spacing=0.26, angle_deg=-135):
    """绘制水平地面基准粗线及 45° 细剖面线。"""
    ax.plot([x_start, x_end], [y, y], color=COLOR_BLACK, lw=LW_THICK, zorder=3)
    rad = math.radians(angle_deg)
    dx, dy = hatch_len * math.cos(rad), hatch_len * math.sin(rad)
    for x in np.arange(x_start + 0.12, x_end - 0.05, spacing):
        ax.plot([x, x + dx], [y, y + dy], color=COLOR_BLACK, lw=LW_THIN, zorder=2)


def draw_wall(ax, x, y_start, y_end, side='right', hatch_len=0.20, spacing=0.25):
    """绘制竖直墙面基准粗线及剖面斜线（side='right' 或 'left'）。"""
    ax.plot([x, x], [y_start, y_end], color=COLOR_BLACK, lw=LW_THICK, zorder=3)
    angle_deg = -45 if side == 'right' else -135
    rad = math.radians(angle_deg)
    dx, dy = hatch_len * math.cos(rad), hatch_len * math.sin(rad)
    for y in np.arange(y_start + 0.14, y_end - 0.05, spacing):
        ax.plot([x, x + dx], [y, y + dy], color=COLOR_BLACK, lw=LW_THIN, zorder=2)


def draw_spring(ax, p1, p2, coils=7, amp=0.12, lead=0.20, lw=LW_THIN):
    """绘制两端带直段的标准锯齿形轻质弹簧。"""
    x1, y1 = p1
    x2, y2 = p2
    dx, dy = x2 - x1, y2 - y1
    length = math.hypot(dx, dy)
    ux, uy = dx / length, dy / length
    nx, ny = -uy, ux
    pts = [(x1, y1), (x1 + ux * lead, y1 + uy * lead)]
    body_len = length - 2 * lead
    step = body_len / (2 * coils)
    for i in range(2 * coils):
        s = lead + (i + 0.5) * step
        sign = 1 if i % 2 == 0 else -1
        pts.append((x1 + ux * s + nx * sign * amp, y1 + uy * s + ny * sign * amp))
    pts.append((x1 + ux * (length - lead), y1 + uy * (length - lead)))
    pts.append((x2, y2))
    xs, ys = zip(*pts)
    ax.plot(xs, ys, color=COLOR_BLACK, lw=lw, solid_joinstyle='miter', zorder=4)


def draw_dim_h(ax, x1, x2, y, label, y_ref, y_ext=None, fs=FS_AX,
               hl=ARROW_HL_DIM, hw=ARROW_HW_DIM, notch=ARROW_NOTCH):
    """绘制水平外置尺寸线（尖端朝外双向实心燕尾箭头 + 丁字界线 + 白底遮罩标签）。"""
    if y_ext is None:
        y_ext = y - 0.14 if y < y_ref else y + 0.14
    ax.plot([x1, x1], [y_ref, y_ext], color=COLOR_BLACK, lw=LW_THIN, ls=':')
    ax.plot([x2, x2], [y_ref, y_ext], color=COLOR_BLACK, lw=LW_THIN, ls=':')
    ax.add_patch(patches.FancyArrowPatch(
        (x1, y), (x2, y),
        arrowstyle=SwallowtailArrow(head_length=hl, head_width=hw, notch=notch, both=True),
        mutation_scale=10.0, color=COLOR_BLACK, lw=LW_THIN,
        joinstyle='miter', shrinkA=0, shrinkB=0, zorder=6
    ))
    ax.text((x1 + x2) / 2, y, label, fontsize=fs, ha='center', va='center',
            bbox=dict(facecolor=COLOR_WHITE, edgecolor='none', pad=1.6), zorder=7)


def draw_dim_v(ax, x, y1, y2, label, x_ref, x_ext=None, fs=FS_AX,
               hl=ARROW_HL_DIM, hw=ARROW_HW_DIM, notch=ARROW_NOTCH):
    """绘制竖直外置尺寸线（尖端朝外双向实心燕尾箭头 + 丁字界线 + 白底遮罩标签）。"""
    if x_ext is None:
        x_ext = x + 0.14 if x > x_ref else x - 0.14
    ax.plot([x_ref, x_ext], [y1, y1], color=COLOR_BLACK, lw=LW_THIN, ls=':')
    ax.plot([x_ref, x_ext], [y2, y2], color=COLOR_BLACK, lw=LW_THIN, ls=':')
    ax.add_patch(patches.FancyArrowPatch(
        (x, y1), (x, y2),
        arrowstyle=SwallowtailArrow(head_length=hl, head_width=hw, notch=notch, both=True),
        mutation_scale=10.0, color=COLOR_BLACK, lw=LW_THIN,
        joinstyle='miter', shrinkA=0, shrinkB=0, zorder=6
    ))
    ax.text(x, (y1 + y2) / 2, label, fontsize=fs, ha='center', va='center',
            bbox=dict(facecolor=COLOR_WHITE, edgecolor='none', pad=1.6), zorder=7)


def draw_right_angle(ax, vertex, u1, u2, size=0.22, lw=LW_THIN, zorder=5):
    """在两正交单位向量 u1、u2 间绘制标准直角符号。"""
    v = np.asarray(vertex, dtype=float)
    e1 = np.asarray(u1, dtype=float)
    e2 = np.asarray(u2, dtype=float)
    e1 = e1 / np.linalg.norm(e1)
    e2 = e2 / np.linalg.norm(e2)
    p1 = v + size * e1
    p2 = v + size * (e1 + e2)
    p3 = v + size * e2
    ax.plot([p1[0], p2[0], p3[0]], [p1[1], p2[1], p3[1]],
            color=COLOR_BLACK, lw=lw, zorder=zorder)


def draw_angle_arc(ax, vertex, radius, theta1_deg, theta2_deg, label='',
                   label_r_scale=1.45, fs=FS_TIK, zorder=4):
    """绘制角度圆弧并自动沿角平分线外置角度标注。"""
    ax.add_patch(patches.Arc(
        vertex, 2 * radius, 2 * radius, angle=0,
        theta1=theta1_deg, theta2=theta2_deg,
        color=COLOR_BLACK, lw=LW_THIN, zorder=zorder
    ))
    if label:
        mid = math.radians(0.5 * (theta1_deg + theta2_deg))
        lx = vertex[0] + radius * label_r_scale * math.cos(mid)
        ly = vertex[1] + radius * label_r_scale * math.sin(mid)
        ax.text(lx, ly, label, fontsize=fs, ha='center', va='center', zorder=zorder + 1)


def draw_cn_sub(ax, x, y, main_math, cn_sub, dx=0.22, dy=-0.14,
                fs_main=FS_SYM, fs_sub=FS_SUB, ha='left', zorder=9):
    """安全渲染带中文下标的物理符号（如 F_弹、v_共），规避 mathtext 不支持 \\text{}。"""
    sub_fp = get_cn_font(size=fs_sub)
    if ha == 'left':
        ax.text(x, y, main_math, fontsize=fs_main, ha='left', va='center', zorder=zorder)
        ax.text(x + dx, y + dy, cn_sub, fontproperties=sub_fp, ha='left', va='center', zorder=zorder)
    else:
        ax.text(x - dx, y, main_math, fontsize=fs_main, ha='right', va='center', zorder=zorder)
        ax.text(x - dx, y + dy, cn_sub, fontproperties=sub_fp, ha='left', va='center', zorder=zorder)


def draw_block_on_slope(ax, x_origin, y_origin, theta_rad, s_start, length, height,
                        label='', fill=COLOR_WHITE, fs=FS_SYM, zorder=4,
                        slope_dir='down_right'):
    """
    在倾角 theta_rad 的斜面上绘制严格贴合物块，返回 (center, left_mid, right_mid)。
    slope_dir:
      - 'down_right': 斜面从左上 (x_origin, y_origin) 向右下倾斜，s_start 沿斜面向下计量；
      - 'up_right':   斜面从左下 (x_origin, y_origin) 向右上倾斜，s_start 沿斜面向上计量。
    """
    if slope_dir == 'up_right':
        ux, uy = math.cos(theta_rad), math.sin(theta_rad)
        nx, ny = -math.sin(theta_rad), math.cos(theta_rad)
    else:
        ux, uy = math.cos(theta_rad), -math.sin(theta_rad)
        nx, ny = math.sin(theta_rad), math.cos(theta_rad)
    p0 = (x_origin + ux * s_start, y_origin + uy * s_start)
    p1 = (p0[0] + ux * length, p0[1] + uy * length)
    p2 = (p1[0] + nx * height, p1[1] + ny * height)
    p3 = (p0[0] + nx * height, p0[1] + ny * height)
    ax.add_patch(patches.Polygon(
        [p0, p1, p2, p3], facecolor=fill, edgecolor=COLOR_BLACK, lw=LW_THICK, zorder=zorder
    ))
    center = ((p0[0] + p2[0]) / 2, (p0[1] + p2[1]) / 2)
    left_mid = ((p0[0] + p3[0]) / 2, (p0[1] + p3[1]) / 2)
    right_mid = ((p1[0] + p2[0]) / 2, (p1[1] + p2[1]) / 2)
    if label:
        ax.text(center[0], center[1], label, fontsize=fs, ha='center', va='center', zorder=zorder + 1)
    return center, left_mid, right_mid


def draw_pulley(ax, center, radius, fill=COLOR_GRAY, lw=LW_THICK, zorder=3):
    """绘制标准定滑轮及中心轴点。"""
    ax.add_patch(patches.Circle(
        center, radius, facecolor=fill, edgecolor=COLOR_BLACK, lw=lw, zorder=zorder
    ))
    ax.plot(center[0], center[1], 'ko', ms=2.6, zorder=zorder + 1)


def draw_rot_arc(ax, cx, cy, rx, ry, theta1=-150, theta2=-25,
                 label=r'$\omega$', lbl_pos=None, lw=LW_MED):
    """在竖直转轴或转盘上方绘制转动方向椭圆弧与实心燕尾箭头。"""
    ts = np.linspace(math.radians(theta1), math.radians(theta2), 60)
    xs = cx + rx * np.cos(ts)
    ys = cy + ry * np.sin(ts)
    ax.plot(xs[:-3], ys[:-3], color=COLOR_BLACK, lw=lw, zorder=7)
    draw_arrow(ax, xs[-6], ys[-6], xs[-1], ys[-1], lw=lw, hl=1.05, hw=0.38, zorder=8)
    if label:
        lx, ly = lbl_pos if lbl_pos else (cx + rx + 0.35, cy)
        ax.text(lx, ly, label, fontsize=FS_SYM, ha='left', va='center', zorder=9)
