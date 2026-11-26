"""
高中物理靶向变式题库技能环境自检与一键自愈脚本（仅依赖 Python 标准库）。

核心作用：
  1. 检测并自动创建/补齐专用 Python 虚拟环境（默认 ~/.local/share/physics-variant-bank-venv），
     按技能根目录 requirements.txt 安装 pillow/python-docx/lxml/matplotlib/numpy/fonttools；
  2. 在 Linux 环境下检测 `宋体/SimSun` 与 `黑体/SimHei` 的 fontconfig 映射，
     若回退到非宋体或日文字形，自动写入非破坏性配置
     ~/.config/fontconfig/conf.d/99-physics-variant-bank-cjk.conf 并刷新缓存；
  3. 检测 pandoc、libreoffice/soffice、pdftoppm 外部工具是否就绪，缺失时给出明确安装指引。

用法：
  python3 scripts/ensure_env.py          # 检测并自动修复 venv 与 fontconfig
  python3 scripts/ensure_env.py --check  # 仅只读检测，不写入任何文件
"""
import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

REQUIRED_IMPORTS = ('PIL', 'docx', 'lxml', 'matplotlib', 'numpy', 'fontTools')
SERIF_SC_CANDIDATES = (
    'Noto Serif CJK SC',
    'Source Han Serif SC',
    'Source Han Serif CN',
    'SimSun',
    'NSimSun',
    'Songti SC',
    'STSong',
    'FandolSong',
)
SANS_SC_CANDIDATES = (
    'Noto Sans CJK SC',
    'Source Han Sans SC',
    'Source Han Sans CN',
    'SimHei',
    'Heiti SC',
    'STHeiti',
    'WenQuanYi Zen Hei',
)


def _find_uv():
    uv = shutil.which('uv')
    if uv:
        return uv
    local_uv = Path.home() / '.local' / 'bin' / 'uv'
    return str(local_uv) if local_uv.is_file() else None


def _venv_python(venv_dir: Path) -> Path:
    if os.name == 'nt':
        return venv_dir / 'Scripts' / 'python.exe'
    return venv_dir / 'bin' / 'python'


def _missing_imports(py_bin: Path):
    if not py_bin.is_file():
        return list(REQUIRED_IMPORTS)
    code = (
        'import importlib, sys\n'
        f'mods = {REQUIRED_IMPORTS!r}\n'
        'missing = []\n'
        'for m in mods:\n'
        '    try:\n'
        '        importlib.import_module(m)\n'
        '    except Exception:\n'
        '        missing.append(m)\n'
        'print(",".join(missing))\n'
    )
    proc = subprocess.run([str(py_bin), '-c', code], capture_output=True, text=True)
    if proc.returncode != 0:
        return list(REQUIRED_IMPORTS)
    out = proc.stdout.strip()
    return [m for m in out.split(',') if m] if out else []


def ensure_venv(venv_dir: Path, req_file: Path, check_only: bool):
    py_bin = _venv_python(venv_dir)
    missing = _missing_imports(py_bin)
    if not missing:
        print(f'[OK] Python 虚拟环境就绪: {py_bin}')
        return True, py_bin

    if check_only:
        print(f'[MISSING] 虚拟环境未就绪或缺少依赖 {missing}: {venv_dir}')
        return False, py_bin

    if not req_file.is_file():
        print(f'[FAIL] 未找到依赖清单: {req_file}')
        return False, py_bin

    uv = _find_uv()
    venv_dir.parent.mkdir(parents=True, exist_ok=True)
    if uv:
        if not py_bin.is_file():
            print(f'[FIX] 使用 uv 创建虚拟环境: {venv_dir}')
            subprocess.check_call([uv, 'venv', str(venv_dir)])
        print(f'[FIX] 使用 uv 安装依赖: {req_file}')
        subprocess.check_call([uv, 'pip', 'install', '--python', str(py_bin), '-r', str(req_file)])
    else:
        if not py_bin.is_file():
            print(f'[FIX] 使用 python3 -m venv 创建虚拟环境: {venv_dir}')
            subprocess.check_call([sys.executable, '-m', 'venv', str(venv_dir)])
        print(f'[FIX] 使用 pip 安装依赖: {req_file}')
        subprocess.check_call([str(py_bin), '-m', 'pip', 'install', '-r', str(req_file)])

    missing_after = _missing_imports(py_bin)
    if missing_after:
        print(f'[FAIL] 安装后仍缺少模块: {missing_after}')
        return False, py_bin

    print(f'[OK] Python 虚拟环境已修复并就绪: {py_bin}')
    return True, py_bin


def _fc_match(family: str) -> str:
    proc = subprocess.run(['fc-match', family], capture_output=True, text=True)
    return proc.stdout.strip() if proc.returncode == 0 else ''


def _installed_zh_families():
    proc = subprocess.run(['fc-list', ':lang=zh', 'family'], capture_output=True, text=True)
    if proc.returncode != 0:
        return set()
    fams = set()
    for line in proc.stdout.splitlines():
        for part in line.split(','):
            p = part.strip()
            if p:
                fams.add(p)
    return fams


def ensure_linux_fontconfig(check_only: bool):
    if not sys.platform.startswith('linux'):
        print('[SKIP] 非 Linux 环境，跳过 fontconfig 字体映射检测。')
        return True

    if not shutil.which('fc-match') or not shutil.which('fc-list'):
        print('[WARN] 未检测到 fontconfig (fc-match/fc-list)，跳过中文字体映射检测。')
        return True

    song_match = _fc_match('宋体')
    hei_match = _fc_match('黑体')
    song_ok = any(k.lower() in song_match.lower() for k in SERIF_SC_CANDIDATES)
    hei_ok = any(k.lower() in hei_match.lower() for k in SANS_SC_CANDIDATES)

    if song_ok and hei_ok:
        print(f'[OK] Linux 中文字体映射正常: 宋体 -> {song_match} | 黑体 -> {hei_match}')
        return True

    installed = _installed_zh_families()
    pick_serif = next((c for c in SERIF_SC_CANDIDATES if c in installed), None)
    pick_sans = next((c for c in SANS_SC_CANDIDATES if c in installed), None)

    if not pick_serif:
        print(
            '[WARN] 系统未安装陆标简体宋体（如 Noto Serif CJK SC），LibreOffice 导出 PDF 将回退至非宋体。'
            '请安装系统字体包，例如: sudo apt install -y fonts-noto-cjk fonts-noto-cjk-extra'
        )
        return False

    if check_only:
        print(f'[MISSING] Linux 中文字体映射未配置: 宋体 -> {song_match} | 黑体 -> {hei_match}')
        return False

    conf_dir = Path.home() / '.config' / 'fontconfig' / 'conf.d'
    conf_dir.mkdir(parents=True, exist_ok=True)
    conf_path = conf_dir / '99-physics-variant-bank-cjk.conf'

    blocks = [
        '<?xml version="1.0"?>',
        '<!DOCTYPE fontconfig SYSTEM "urn:fontconfig:fonts.dtd">',
        '<fontconfig>',
        '  <!-- 由 physics-variant-bank/scripts/ensure_env.py 自动生成：映射 Windows 宋体/黑体至陆标简体 -->',
    ]
    for alias_name in ('宋体', 'SimSun', 'NSimSun'):
        blocks.extend([
            '  <alias>',
            f'    <family>{alias_name}</family>',
            f'    <prefer><family>{pick_serif}</family></prefer>',
            '  </alias>',
        ])
    if pick_sans:
        for alias_name in ('黑体', 'SimHei'):
            blocks.extend([
                '  <alias>',
                f'    <family>{alias_name}</family>',
                f'    <prefer><family>{pick_sans}</family></prefer>',
                '  </alias>',
            ])
    blocks.append('</fontconfig>')
    conf_path.write_text('\n'.join(blocks) + '\n', encoding='utf-8')
    print(f'[FIX] 已写入非破坏性 fontconfig 字体映射: {conf_path}')

    if shutil.which('fc-cache'):
        subprocess.run(['fc-cache', '-f'], check=False)

    song_after = _fc_match('宋体')
    hei_after = _fc_match('黑体')
    song_ok_after = any(k.lower() in song_after.lower() for k in SERIF_SC_CANDIDATES)
    if song_ok_after:
        print(f'[OK] Linux 中文字体映射已生效: 宋体 -> {song_after} | 黑体 -> {hei_after}')
        return True

    print(f'[WARN] 写入配置后宋体仍未命中陆标简体（当前: {song_after}），请检查 ~/.config/fontconfig/')
    return False


def check_system_tools():
    tools = {
        'pandoc': shutil.which('pandoc'),
        'libreoffice': shutil.which('libreoffice') or shutil.which('soffice'),
        'pdftoppm': shutil.which('pdftoppm'),
    }
    missing = [k for k, v in tools.items() if not v]
    if not missing:
        print('[OK] 外部文档工具就绪: pandoc, libreoffice, pdftoppm')
        return True
    print(f'[MISSING] 缺少系统级工具: {", ".join(missing)}')
    print('  Debian/Ubuntu 参考安装命令: sudo apt install -y pandoc libreoffice-writer poppler-utils')
    return False


def main():
    parser = argparse.ArgumentParser(description='physics-variant-bank 技能环境自检与一键自愈')
    parser.add_argument(
        '--venv',
        default=os.environ.get('PVB_VENV', str(Path.home() / '.local' / 'share' / 'physics-variant-bank-venv')),
        help='专用虚拟环境目录（默认 ~/.local/share/physics-variant-bank-venv）',
    )
    parser.add_argument('--check', action='store_true', help='仅检测，不自动安装或写入配置')
    args = parser.parse_args()

    skill_root = Path(__file__).resolve().parent.parent
    req_file = skill_root / 'requirements.txt'
    venv_dir = Path(args.venv).expanduser()

    ok_venv, py_bin = ensure_venv(venv_dir, req_file, args.check)
    ok_font = ensure_linux_fontconfig(args.check)
    ok_tools = check_system_tools()

    print(f'VENV_PY={py_bin}')
    if ok_venv and ok_font and ok_tools:
        print('[ENV PASS] 全部运行环境检查通过。')
        sys.exit(0)
    print('[ENV WARN] 存在未就绪的环境项，请按上方提示处理。')
    sys.exit(1)


if __name__ == '__main__':
    main()
